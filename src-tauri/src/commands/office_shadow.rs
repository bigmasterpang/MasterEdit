//! 结构操作的「影子工作簿」：结构改动只改内存，Ctrl+S 才落盘（0.20.0）。
//!
//! ## 为什么不直接写盘
//! 0.20.0 之前 `spreadsheet_structure` 是「执行即整簿写回」，用户按错一下行列就落盘了，
//! 中间没法反悔。现在结构操作只改内存里的影子工作簿，用户按 Ctrl+S（`spreadsheet_save`）
//! 才真正写文件 —— 于是「插入一行 → 觉得不对 → 不保存关掉」不会留下任何痕迹。
//! 保存**不再默认生成 `.bak`**（可显式要求），磁盘上的历史 `.bak` 也不会被删改。
//!
//! ## 生命周期
//! - **建立**：第一次对某个 `.xlsx` 做结构操作时从磁盘读入（企业加密文件在**内存中**
//!   解密，绝不落明文），同时记下「原始字节」（可选备份用、重新加密时当文件头）与
//!   磁盘文件指纹（大小 + 修改时间）。
//! - **累积**：之后的每一次结构操作都改同一个影子（`edits` 先应用、压快照、再应用结构 op），
//!   每次改完重新 `write_writer` 序列化一份字节缓存 —— 所有读路径
//!   （`spreadsheet_info` / `rows` / `find` / `stats`）直接拿这份缓存喂 calamine，
//!   因此**不新增 umya 读路径**，数字/日期显示口径与磁盘读取完全一致。
//! - **保存**：写回原文件成功后**保留影子与撤销栈**（用户要求「即使保存也不影响撤回」），
//!   只把「磁盘基线」挪到当前状态：`pending` 归 false、`original_bytes`/指纹更新。
//!   于是刚保存完也能一路 Ctrl+Z 撤回去，撤销后影子重新变脏、再按 Ctrl+S 又写回。
//! - **撤销/重做**：每做一次**结构改动**，就把「该改动之前」的整簿序列化字节压进撤销栈
//!   （见 [`ShadowDoc::push_undo`] 的容量策略）；[`undo`] / [`redo`] 用快照重建工作簿。
//!   快照时机是「`edits` 之后、结构 op 之前」，理由见 [`apply_op`] 的说明。
//! - **清空**：用户主动放弃 / 关闭文档 / 检测到外部改动时用 `spreadsheet_discard` 丢弃
//!   （撤销栈一并清空）。另外，**干净影子**（保存后、没有任何未保存改动）一旦发现磁盘
//!   被外部改过，会在读路径上自动作废（见 [`ShadowStore::bytes_of`]），避免读到旧内容。
//! - **容量**：最多 [`MAX_SHADOWS`] 个影子。新增时优先淘汰「不脏」的最久未用者；
//!   若全都带未落盘改动，直接报错让用户先保存或关掉一个，绝不悄悄丢掉别人的改动。
//!
//! ## 读路径如何看到影子
//! `office.rs::read_workbook_bytes` 一开始就问 [`shadow_bytes`]，命中就返回影子内容；
//! 解析缓存（`SheetCache`）的 key 里带上 [`revision`]（没有影子时恒为 0），
//! 因为结构操作不改磁盘 mtime，光靠 mtime 会命中改动前的旧解析结果。
//!
//! ## 内存代价
//! 每个影子 = 一份已解析的 umya 工作簿 + 一份序列化字节 + 撤销/重做快照。
//! 快照是整簿字节（一张 50 MB 的表每步约 50 MB），所以撤销栈受
//! [`MAX_UNDO_STEPS`]（30 步）与 [`MAX_UNDO_BYTES`]（512 MB，优先）双重约束。
//!
//! ## 并发取舍
//! 结构操作/撤销在**持锁**状态下完成「取影子 → 应用 → 序列化 → 放回」，期间其它表格命令会
//! 稍等（用户动作级别，代价可接受）；保存因为要写盘，先把影子**取出**再在锁外写，
//! 失败时放回（这期间读路径会临时看到磁盘内容）。

use crate::commands::office::{self, SheetMeta};
use crate::commands::office_ops::{
    apply_structure, ShadowEditResult, SpreadsheetState, StructureOp, StructureResult,
};
use crate::commands::office_write::{self, apply_edits, CellEdit, SaveResult};
use std::io::Cursor;
use std::path::Path;
use std::sync::atomic::{AtomicU64, Ordering};
use std::sync::{Mutex, MutexGuard};
use umya_spreadsheet::Workbook;

/// 同时存在的影子上限（每个影子 = 一份已解析的工作簿 + 一份序列化字节 + 撤销快照）
pub(crate) const MAX_SHADOWS: usize = 8;

/// 撤销栈的步数上限（超出丢最旧的）
pub(crate) const MAX_UNDO_STEPS: usize = 30;

/// 撤销快照的总字节上限（超出丢最旧的，但至少保留最近一步）。
///
/// **字节上限优先于步数上限**：快照是「整簿序列化字节」，一张 50 MB 的工作簿每步就是
/// 约 50 MB，30 步就是约 1.5 GB —— 所以先按 512 MB 截断，实际能回退的步数由工作簿大小
/// 决定（50 MB 的表约 10 步，几 MB 的表能占满 30 步）。上限对**每个影子**（每个文档）
/// 单独计算，所以「同时打开 8 个大表且每个都堆满快照」是理论最坏情况，远大于 512 MB。
pub(crate) const MAX_UNDO_BYTES: usize = 512 * 1024 * 1024;

/// 全局版本号：每次影子内容变化 +1。0 保留给「没有影子」，用于解析缓存的 key。
static NEXT_REVISION: AtomicU64 = AtomicU64::new(0);

fn next_revision() -> u64 {
    NEXT_REVISION.fetch_add(1, Ordering::SeqCst) + 1
}

/// 磁盘文件的指纹（大小 + 修改时间），用来识别「影子还干净时文件被外部改了」
#[derive(Clone, Copy, PartialEq, Eq, Debug)]
struct DiskStamp {
    len: u64,
    modified_ms: u64,
}

impl DiskStamp {
    fn of(path: &str) -> Option<DiskStamp> {
        let meta = std::fs::metadata(path).ok()?;
        Some(DiskStamp {
            len: meta.len(),
            modified_ms: crate::commands::file::modified_ms(&meta),
        })
    }
}

/// 一个文件的内存影子
pub(crate) struct ShadowDoc {
    /// 内存工作簿（结构操作与保存都改它）
    pub(crate) workbook: Workbook,
    /// 企业加密文件的 4096 字节头（非加密为 None）；保存时用它重新加密
    header: Option<Vec<u8>>,
    /// 建立影子（或上次保存）时磁盘上的**原始字节**（企业加密文件即密文）：做 `.bak` 用
    original_bytes: Vec<u8>,
    /// 当前影子的明文 xlsx 序列化结果：所有读路径直接拿它喂 calamine
    serialized: Vec<u8>,
    /// 撤销栈：每项是「某次改动**之前**」的整簿序列化快照（末尾为最近一步）
    undo: Vec<Vec<u8>>,
    /// 重做栈：撤销时把「撤销前」的快照压进来
    redo: Vec<Vec<u8>>,
    /// 「磁盘上的内容」对应撤销栈的哪个位置（0 = 建立影子时的状态）；
    /// None = 当前分支上已经没有那个状态了（必然有未落盘改动）
    saved_pos: Option<usize>,
    /// 建立影子/上次保存时磁盘文件的指纹：干净影子遇到外部改动会自动作废
    stamp: Option<DiskStamp>,
    /// 内容版本号（进解析缓存的 key）
    revision: u64,
}

impl ShadowDoc {
    /// 该影子是否加密文件（保存时决定要不要重新加密）
    fn encrypted(&self) -> bool {
        self.header.is_some()
    }

    /// 是否有未落盘的改动：当前状态是否就是「磁盘上那个状态」
    fn dirty(&self) -> bool {
        self.saved_pos != Some(self.undo.len())
    }

    /// 撤销栈位置（快照数）
    fn depth(&self) -> usize {
        self.undo.len()
    }

    /// 压入一步撤销快照，并按「步数 + 总字节」上限丢最旧的。
    ///
    /// 丢最旧的会把整个栈的位置整体前移，因此 `saved_pos` 也要跟着前移；如果连
    /// 「磁盘那个状态」都被丢掉了，就把 `saved_pos` 置空（永远算脏）。
    fn push_undo(&mut self, snapshot: Vec<u8>) {
        self.undo.push(snapshot);
        let dropped = trim_snapshots(&mut self.undo, MAX_UNDO_STEPS, MAX_UNDO_BYTES);
        if dropped > 0 {
            self.saved_pos = match self.saved_pos {
                Some(pos) if pos >= dropped => Some(pos - dropped),
                _ => None,
            };
        }
    }

    /// 重新序列化（把内存工作簿的当前状态同步到读路径用的字节缓存）。
    ///
    /// 失败时**保留旧缓存**：宁可读到稍旧的内容，也不能让「读到的字节」与「工作簿」
    /// 不一致到无法解释（下次成功改动会再同步一次）。
    fn resync(&mut self) {
        match office_write::serialize_workbook(&self.workbook) {
            Ok(bytes) => {
                self.serialized = bytes;
                self.revision = next_revision();
            }
            Err(_) => {}
        }
    }

    /// 用快照重建影子（撤销/重做共用）：反序列化回 umya 工作簿 + 更新字节缓存与版本号。
    /// 失败时把快照原样退回给调用方，方便整步回滚。
    fn restore(&mut self, snapshot: Vec<u8>) -> std::result::Result<(), (String, Vec<u8>)> {
        match umya_spreadsheet::reader::xlsx::read_reader(Cursor::new(snapshot.as_slice()), true) {
            Ok(workbook) => {
                self.workbook = workbook;
                self.serialized = snapshot;
                self.revision = next_revision();
                Ok(())
            }
            Err(e) => Err((format!("还原撤销快照失败：{e}"), snapshot)),
        }
    }
}

/// 影子内容（读路径要的字节 + 是否加密）
pub(crate) struct ShadowBytes {
    pub bytes: Vec<u8>,
    pub encrypted: bool,
}

/// 按上限裁剪快照栈（丢掉最旧的），返回丢掉的条数。
///
/// **字节上限优先于步数上限**：快照是整簿序列化字节，一张 50 MB 的表每步约 50 MB，
/// 30 步就是约 1.5 GB —— 所以先按 [`MAX_UNDO_BYTES`] 截断。至少保留最近一步，
/// 否则大表会完全无法撤销。
fn trim_snapshots(snapshots: &mut Vec<Vec<u8>>, max_steps: usize, max_bytes: usize) -> usize {
    let mut dropped = 0usize;
    while snapshots.len() > max_steps {
        snapshots.remove(0);
        dropped += 1;
    }
    let mut total: usize = snapshots.iter().map(Vec::len).sum();
    while snapshots.len() > 1 && total > max_bytes {
        total -= snapshots.remove(0).len();
        dropped += 1;
    }
    dropped
}

/// 影子表：`entries` 末尾为最近使用（LRU）
#[derive(Default)]
struct ShadowInner {
    entries: Vec<(String, ShadowDoc)>,
}

/// 影子仓库。生产用全局 [`SHADOWS`]；测试可以 `ShadowStore::new(n)` 造独立实例，
/// 因此容量淘汰、脏影子报错这些行为都能确定性地测。
pub(crate) struct ShadowStore {
    capacity: usize,
    inner: Mutex<ShadowInner>,
}

impl ShadowStore {
    /// `const` 构造函数：全局静态量需要它（`Mutex::new` 与 `Vec::new` 都是 const）
    pub(crate) const fn new(capacity: usize) -> Self {
        ShadowStore {
            capacity,
            inner: Mutex::new(ShadowInner {
                entries: Vec::new(),
            }),
        }
    }

    /// 锁（忽略 poison：某个测试 panic 不该让后续调用全部失败）
    fn lock(&self) -> MutexGuard<'_, ShadowInner> {
        self.inner.lock().unwrap_or_else(|e| e.into_inner())
    }

    fn position(entries: &[(String, ShadowDoc)], key: &str) -> Option<usize> {
        entries.iter().position(|(k, _)| k == key)
    }

    /// 是否有未落盘的结构改动
    pub(crate) fn is_pending(&self, path: &str) -> bool {
        let key = shadow_key(path);
        let inner = self.lock();
        Self::position(&inner.entries, &key)
            .map(|i| inner.entries[i].1.dirty())
            .unwrap_or(false)
    }

    /// 是否有可撤销 / 可重做的步骤（没有影子都是 false）
    pub(crate) fn undo_flags(&self, path: &str) -> (bool, bool) {
        let key = shadow_key(path);
        let inner = self.lock();
        Self::position(&inner.entries, &key)
            .map(|i| {
                let doc = &inner.entries[i].1;
                (!doc.undo.is_empty(), !doc.redo.is_empty())
            })
            .unwrap_or((false, false))
    }

    /// 是否有影子（不论脏不脏）
    pub(crate) fn has(&self, path: &str) -> bool {
        let key = shadow_key(path);
        let inner = self.lock();
        Self::position(&inner.entries, &key).is_some()
    }

    /// 影子内容版本号；没有影子时为 0
    pub(crate) fn revision_of(&self, path: &str) -> u64 {
        let key = shadow_key(path);
        let inner = self.lock();
        Self::position(&inner.entries, &key)
            .map(|i| inner.entries[i].1.revision)
            .unwrap_or(0)
    }

    /// 取影子内容（同时把它标记为最近使用）。
    ///
    /// 顺带做一件重要的事：**干净影子遇到「磁盘被外部改动」就自动作废**。保存之后影子会
    /// 保留（为了撤销栈），此时如果别的程序改了文件，继续读影子就会读到旧内容；这里比对
    /// 建立影子/上次保存时记下的文件指纹（大小 + 修改时间），不一致且影子是干净的
    /// （没有任何未保存改动、丢弃无损失）就把它丢掉，改读磁盘。脏影子不动 —— 那属于
    /// 用户的未保存改动，只能由 `spreadsheet_discard` 或保存来处理。
    pub(crate) fn bytes_of(&self, path: &str) -> Option<ShadowBytes> {
        let key = shadow_key(path);
        let mut inner = self.lock();
        let index = Self::position(&inner.entries, &key)?;
        let stale = {
            let doc = &inner.entries[index].1;
            !doc.dirty() && doc.stamp.is_some() && doc.stamp != DiskStamp::of(path)
        };
        if stale {
            inner.entries.remove(index);
            return None;
        }
        let entry = inner.entries.remove(index);
        let out = ShadowBytes {
            bytes: entry.1.serialized.clone(),
            encrypted: entry.1.encrypted(),
        };
        inner.entries.push(entry);
        Some(out)
    }

    /// 影子里的工作表列表（前端刷新底部标签用）
    pub(crate) fn sheets_of(&self, path: &str) -> Option<Vec<SheetMeta>> {
        let key = shadow_key(path);
        let inner = self.lock();
        let index = Self::position(&inner.entries, &key)?;
        Some(sheet_metas(&inner.entries[index].1.workbook))
    }

    /// 取出影子（后续在锁外做重活）。之后必须 `insert` 放回或明确丢弃。
    fn take(&self, path: &str) -> Option<ShadowDoc> {
        let key = shadow_key(path);
        let mut inner = self.lock();
        let index = Self::position(&inner.entries, &key)?;
        Some(inner.entries.remove(index).1)
    }

    /// 占用一个位置前先问一句：要么本来就有影子，要么还有空位，要么能淘汰一个不脏的。
    /// 这样「操作做完才发现放不下」不会发生 —— 用户的改动不会白做。
    fn ensure_slot(&self, path: &str) -> Result<(), String> {
        let key = shadow_key(path);
        let inner = self.lock();
        if Self::position(&inner.entries, &key).is_some() {
            return Ok(());
        }
        if inner.entries.len() < self.capacity || inner.entries.iter().any(|(_, d)| !d.dirty()) {
            return Ok(());
        }
        Err(format!(
            "同时有未保存结构改动的表格过多（上限 {0} 个），请先保存或关闭其中一个",
            self.capacity
        ))
    }

    /// 放回影子。容量满时优先淘汰最久未用的「不脏」条目；**若全都有未落盘改动，仍然放回** ——
    /// 宁可短暂超出上限，也不能悄悄丢掉用户还没保存的改动。真正拦人的是 [`Self::ensure_slot`]：
    /// 它在动影子之前就把「太多表格有未保存改动」报给用户，用户的这次操作根本还没开始。
    fn insert(&self, path: &str, doc: ShadowDoc) {
        let key = shadow_key(path);
        let mut inner = self.lock();
        if let Some(index) = Self::position(&inner.entries, &key) {
            inner.entries.remove(index);
        }
        if inner.entries.len() >= self.capacity {
            if let Some(index) = inner.entries.iter().position(|(_, d)| !d.dirty()) {
                inner.entries.remove(index);
            }
        }
        inner.entries.push((key, doc));
    }

    /// 丢弃影子（返回是否真的丢掉了东西）
    pub(crate) fn discard(&self, path: &str) -> bool {
        let key = shadow_key(path);
        let mut inner = self.lock();
        match Self::position(&inner.entries, &key) {
            Some(index) => {
                inner.entries.remove(index);
                true
            }
            None => false,
        }
    }

    /// 当前影子数量（测试用）
    #[cfg(test)]
    fn len(&self) -> usize {
        self.lock().entries.len()
    }

    /// 某个影子的撤销栈深度（测试用）
    #[cfg(test)]
    fn undo_len(&self, path: &str) -> Option<usize> {
        let key = shadow_key(path);
        let inner = self.lock();
        Self::position(&inner.entries, &key).map(|i| inner.entries[i].1.undo.len())
    }

    /// 直接塞一个影子进去（测试用：构造「不脏」的条目来验证淘汰策略）
    #[cfg(test)]
    fn insert_for_test(&self, path: &str, doc: ShadowDoc) {
        self.insert(path, doc)
    }
}

/// 生产用全局影子仓库
pub(crate) static SHADOWS: ShadowStore = ShadowStore::new(MAX_SHADOWS);

/* ------------------------------------------------------------------ */
/* 路径规范化                                                          */
/* ------------------------------------------------------------------ */

/// 影子 key：规范化路径（Windows 大小写不敏感、分隔符统一），避免同一文件两份影子。
///
/// 先 `canonicalize` 解析真实路径（能消掉 `..`、8.3 短名、大小写差异与不同拼法），
/// 文件不存在（例如刚被删掉）时退回字符串规整。
pub(crate) fn shadow_key(path: &str) -> String {
    let p = Path::new(path);
    let resolved = p.canonicalize().unwrap_or_else(|_| p.to_path_buf());
    let text = resolved
        .to_string_lossy()
        .replace('/', std::path::MAIN_SEPARATOR_STR);
    if cfg!(windows) {
        text.to_lowercase()
    } else {
        text
    }
}

/* ------------------------------------------------------------------ */
/* 全局读接口（office.rs 的读路径直接用这几个）                          */
/* ------------------------------------------------------------------ */

/// 影子内容（没有影子返回 None）
pub(crate) fn shadow_bytes(path: &str) -> Option<ShadowBytes> {
    SHADOWS.bytes_of(path)
}

/// 影子内容版本号（没有影子为 0）；进 calamine 解析缓存的 key
pub(crate) fn revision(path: &str) -> u64 {
    SHADOWS.revision_of(path)
}

/// 是否有未落盘的结构改动
pub(crate) fn pending(path: &str) -> bool {
    SHADOWS.is_pending(path)
}

/// 丢弃影子（返回是否真的丢掉了东西）。同时清空撤销/重做栈（影子整个没了）。
pub(crate) fn discard(path: &str) -> bool {
    SHADOWS.discard(path)
}

/* ------------------------------------------------------------------ */
/* 影子建立与结构操作                                                   */
/* ------------------------------------------------------------------ */

/// 工作簿里的工作表名 → `SheetMeta` 列表
fn sheet_metas(workbook: &Workbook) -> Vec<SheetMeta> {
    workbook
        .sheet_collection_no_check()
        .iter()
        .map(|s| SheetMeta {
            name: s.name().to_string(),
        })
        .collect()
}

/// 从磁盘建立影子：只有 `.xlsx` 可以（整簿重写会丢 xlsm 的宏、xls/xlsb/ods 结构不同），
/// 企业加密文件在内存中解密，同时记下原始字节与加密头。
fn build_shadow(path: &str) -> Result<ShadowDoc, String> {
    office_write::editable_required(path)?;
    let p = Path::new(path);
    let meta = std::fs::metadata(p).map_err(|e| format!("读取文件信息失败: {e}"))?;
    if !meta.is_file() {
        return Err("目标不是有效的文件".to_string());
    }
    let (raw, plain, encrypted) = office::read_raw_and_plain(p)?;
    let header = if encrypted {
        Some(raw[..raw.len().min(4096)].to_vec())
    } else {
        None
    };
    let workbook = umya_spreadsheet::reader::xlsx::read_reader(Cursor::new(plain.as_slice()), true)
        .map_err(|e| office_write::describe_open_failure(&plain, &e.to_string()))?;
    // 先序列化一份：即使这次操作最后是空操作，读路径也能立刻安全地走影子
    let serialized = office_write::serialize_workbook(&workbook)?;
    Ok(ShadowDoc {
        workbook,
        header,
        original_bytes: raw,
        serialized,
        // 撤销栈从空开始，当前位置 0 就是「磁盘上的内容」
        undo: Vec::new(),
        redo: Vec::new(),
        saved_pos: Some(0),
        stamp: DiskStamp::of(path),
        revision: next_revision(),
    })
}

/// 结构操作：`edits` 先应用到影子，再应用结构 op，**全程不碰磁盘**。
///
/// ## 快照时机（决定撤销语义是否与前端对齐）
/// 顺序是 **应用 `edits` → 压快照 → 应用结构 op**：
/// - 前端把单元格编辑也记在自己的同一条撤销栈里（`kind: "cells"` 的步骤，由前端自己回退）；
/// - 所以「撤销这次结构操作」必须回到「**单元格编辑已生效、结构没做**」的状态。若快照在
///   `edits` 之前压，撤销会把那些编辑一起回退，而前端栈里的编辑步骤还在 → 再按一次 Ctrl+Z
///   就会重复撤销、状态错乱。两条栈因此严格一一对应：后端每收到一次结构改动压一步。
/// - 反过来说：只带 `edits` 而结构 op 是空操作（越界插入/删除）时**不压结构快照** ——
///   那一步属于前端的「cells」步骤，后端替它压栈反而会错位。
///
/// ## 其它
/// - 越界的插入/删除是空操作（见 `office_ops::apply_structure`）：既不留影子也不报 pending、
///   也不压栈；只带 `edits` 时内容确实变了，影子会留下并报 `pending: true`；
/// - 失败时报错，且**新建的影子直接丢掉**（磁盘与内存都没改过结构）；已有影子放回原处；
/// - 容量不足（已有 [`MAX_SHADOWS`] 个脏影子）在动影子之前就报错，用户的改动不会白做。
pub(crate) fn apply_op(
    store: &ShadowStore,
    path: &str,
    op: &StructureOp,
    edits: &[CellEdit],
) -> Result<StructureResult, String> {
    let had = store.has(path);
    if !had {
        store.ensure_slot(path)?;
    }
    let mut doc = match store.take(path) {
        Some(doc) => doc,
        None => build_shadow(path)?,
    };

    // 1) 先应用本次请求带进来的单元格编辑（前端把它们记成独立的「cells」撤销步）
    let edited = match apply_edits(&mut doc.workbook, edits) {
        Ok(edited) => edited,
        Err(err) => {
            if !had {
                // 新影子：直接丢掉，工作簿与磁盘都保持原样
                return Err(err);
            }
            // 已有影子：把内存工作簿重新序列化（edits 可能已经落进去了），放回原处。
            // 只要带了 edits 就不能再声称「和磁盘一样」（可能已部分生效）
            if !edits.is_empty() {
                doc.saved_pos = None;
            }
            doc.resync();
            store.insert(path, doc);
            return Err(err);
        }
    };
    if edited > 0 {
        // 让字节缓存反映「edits 之后」的状态：下面的快照必须是这个状态，
        // 读路径也应当立刻看到这些编辑
        doc.resync();
    }

    // 2) 结构 op 之前的快照：用「移动」而不是「克隆」，空操作时再原样放回
    let snapshot = std::mem::take(&mut doc.serialized);

    // 3) 应用结构 op
    let changed = match apply_structure(&mut doc.workbook, op) {
        Ok(changed) => changed,
        Err(err) => {
            if !had {
                return Err(err);
            }
            // 结构没改成，但 edits 可能已经生效：不能再声称「和磁盘一样」，并把缓存同步回来
            if edited > 0 {
                doc.saved_pos = None;
            }
            doc.resync();
            store.insert(path, doc);
            return Err(err);
        }
    };

    let sheets = sheet_metas(&doc.workbook);
    if changed {
        // 「磁盘那个状态」还在当前分支上吗？
        // - 本次带了 edits（内容在快照之前就已经变了）→ 任何撤销位置都不等于磁盘内容；
        // - 撤销过之后又做了新操作 → 那个未来被截断了，再也回不去。
        // 两种情况都必须让影子保持「脏」，否则撤销回来会误报「与磁盘一致」。
        if edited > 0 || matches!(doc.saved_pos, Some(pos) if pos > doc.depth()) {
            doc.saved_pos = None;
        }
        doc.push_undo(snapshot);
        doc.redo.clear();
        // 同步读路径用的字节缓存（失败也只是让预览暂时落后一步，保存写的仍是工作簿）
        doc.resync();
        store.insert(path, doc);
    } else if had || edited > 0 {
        // 空操作：快照（= edits 之后的字节缓存）原样放回
        doc.serialized = snapshot;
        if edited > 0 {
            // 结构没改，但 edits 改了内容：影子确实与磁盘不同
            doc.saved_pos = None;
        }
        let pending = doc.dirty();
        store.insert(path, doc);
        return Ok(StructureResult {
            path: path.to_string(),
            sheets,
            saved: false,
            pending,
        });
    }
    // 空操作 + 没有编辑 + 本来没有影子 → 不留痕
    Ok(StructureResult {
        path: path.to_string(),
        sheets,
        saved: false,
        pending: changed,
    })
}

/* ------------------------------------------------------------------ */
/* 撤销 / 重做                                                         */
/* ------------------------------------------------------------------ */

/// 撤销/重做共用的实现：`forward == false` 撤销，`true` 重做。
///
/// 把当前状态压进另一个栈、再从目标栈弹出快照重建影子（工作簿 + 字节缓存 + 版本号）。
/// 栈空了就报「没有可撤销/可重做的改动」，并且**不动影子**（原样放回）。
/// 快照损坏时整步回滚（两个栈都恢复原状），不会把影子弄丢。
fn step(store: &ShadowStore, path: &str, forward: bool) -> Result<ShadowEditResult, String> {
    let empty_msg = if forward {
        "没有可重做的改动"
    } else {
        "没有可撤销的改动"
    };
    let Some(mut doc) = store.take(path) else {
        return Err(empty_msg.to_string());
    };

    // 当前状态先进「另一个栈」，再从来源栈取快照
    let snapshot = if forward {
        if doc.redo.is_empty() {
            store.insert(path, doc);
            return Err(empty_msg.to_string());
        }
        let snapshot = doc.redo.pop().unwrap_or_default();
        doc.undo.push(std::mem::take(&mut doc.serialized));
        snapshot
    } else {
        if doc.undo.is_empty() {
            store.insert(path, doc);
            return Err(empty_msg.to_string());
        }
        let snapshot = doc.undo.pop().unwrap_or_default();
        doc.redo.push(std::mem::take(&mut doc.serialized));
        snapshot
    };

    if let Err((err, snapshot)) = doc.restore(snapshot) {
        // 快照坏了：两个栈恢复原状（刚才移出去的那个状态从另一个栈取回）
        if forward {
            doc.redo.push(snapshot);
            doc.serialized = doc.undo.pop().unwrap_or_default();
        } else {
            doc.undo.push(snapshot);
            doc.serialized = doc.redo.pop().unwrap_or_default();
        }
        store.insert(path, doc);
        return Err(err);
    }

    let result = ShadowEditResult {
        path: path.to_string(),
        sheets: sheet_metas(&doc.workbook),
        pending: doc.dirty(),
        can_undo: !doc.undo.is_empty(),
        can_redo: !doc.redo.is_empty(),
    };
    store.insert(path, doc);
    Ok(result)
}

/// 撤销上一次改动（结构操作 / 随请求带进来的单元格编辑）
pub(crate) fn undo(store: &ShadowStore, path: &str) -> Result<ShadowEditResult, String> {
    step(store, path, false)
}

/// 重做上一次被撤销的改动
pub(crate) fn redo(store: &ShadowStore, path: &str) -> Result<ShadowEditResult, String> {
    step(store, path, true)
}

/// 当前未落盘状态 + 工作表列表 + 可撤销/可重做标志（同步实现，测试直接调用）
pub(crate) fn state(store: &ShadowStore, path: &str) -> Result<SpreadsheetState, String> {
    let sheets = match store.sheets_of(path) {
        Some(sheets) => sheets,
        None => office::info_impl(path)?.sheets,
    };
    let (can_undo, can_redo) = store.undo_flags(path);
    Ok(SpreadsheetState {
        path: path.to_string(),
        sheets,
        pending: store.is_pending(path),
        can_undo,
        can_redo,
    })
}

/* ------------------------------------------------------------------ */
/* 保存（Ctrl+S）                                                      */
/* ------------------------------------------------------------------ */

/// 命令入口（同步）：`backup` 传 `None` 时**不生成 `.bak`**（0.20.0 起的默认）。
///
/// 单独抽出来是为了让「默认不备份」这条约定本身可测（`spreadsheet_save` 命令体只做
/// `spawn_blocking`，测试跑不了 Tauri 运行时）。
pub(crate) fn save_requested(
    store: &ShadowStore,
    path: &str,
    target: Option<&str>,
    edits: &[CellEdit],
    backup: Option<bool>,
) -> Result<SaveResult, String> {
    save(store, path, target, edits, backup.unwrap_or(false))
}

/// 保存：**有影子就用影子**（先把本次 `edits` 应用到影子，再走原有写盘流水线：
/// `.bak`（可选，默认不生成）= 上次保存/建立影子时磁盘上的原始字节、临时文件 → calamine
/// 校验 → 原子替换、企业加密文件按原 4096 字节头重新加密）；没有影子时完全保持原行为。
///
/// 写回原文件成功后**保留影子与撤销栈**（用户明确要求「即使保存也不影响撤回」）：
/// 只把磁盘基线挪到当前状态（`pending` 归 false、`original_bytes`/指纹更新），
/// 于是保存之后还能一路撤销回去，撤销后影子重新变脏、再按 Ctrl+S 又写回去。
/// 另存为（写到别的文件）时影子保持脏状态，因为原文件仍有未落盘的改动。
pub(crate) fn save(
    store: &ShadowStore,
    path: &str,
    target: Option<&str>,
    edits: &[CellEdit],
    backup: bool,
) -> Result<SaveResult, String> {
    let Some(mut doc) = store.take(path) else {
        // 没有影子：走原有磁盘流水线（读盘 → 应用编辑 → 写出）
        return office_write::save_pipeline(path, target, backup, |workbook| {
            apply_edits(workbook, edits)
        });
    };

    let source = Path::new(path);
    let result = (|| -> Result<(SaveResult, bool), String> {
        office_write::editable_required(path)?;
        let dest = office_write::resolve_target(source, target)?;
        // resolve_target 在「目标就是原文件」时原样返回 source
        let in_place = dest == source;
        let saved_cells = apply_edits(&mut doc.workbook, edits)?;
        let saved = office_write::write_back_workbook(
            &dest,
            &doc.workbook,
            saved_cells,
            if backup { Some(source) } else { None },
            &doc.original_bytes,
            doc.encrypted(),
        )?;
        Ok((saved, in_place))
    })();

    match result {
        Ok((saved, true)) => {
            // 写回原文件成功：**保留影子与撤销栈**（用户要求「即使保存也不影响撤回」），
            // 只把「磁盘基线」挪到当前状态 —— 重新记下刚写出的原始字节（下次做 .bak 用）
            // 与文件指纹（干净影子遇到外部改动会自动作废），撤销栈原地不动。
            if let Ok(bytes) = std::fs::read(source) {
                doc.original_bytes = bytes;
            }
            doc.stamp = DiskStamp::of(path);
            doc.saved_pos = Some(doc.depth());
            store.insert(path, doc);
            Ok(saved)
        }
        Ok((saved, false)) => {
            // 另存为：目标写好了，但原文件仍有未落盘的改动，影子保持脏状态放回去
            doc.resync();
            store.insert(path, doc);
            Ok(saved)
        }
        Err(err) => {
            // 保存失败：影子放回，后续重试仍能写出完整内容（edits 可能已部分进工作簿，
            // 前端仍持有 sheetEdits 列表，重试是幂等的）
            if !edits.is_empty() {
                doc.saved_pos = None;
            }
            doc.resync();
            store.insert(path, doc);
            Err(err)
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::commands::office_write::CellEdit;
    use crate::commands::office_ops::StructureOp;
    use rust_xlsxwriter::{Format, Workbook as XlsxWriter};

    fn test_dir(name: &str) -> std::path::PathBuf {
        let dir = std::env::temp_dir()
            .join("masteredit-office-shadow")
            .join(name);
        let _ = std::fs::remove_dir_all(&dir);
        std::fs::create_dir_all(&dir).expect("创建测试目录");
        dir
    }

    fn make_sample(path: &Path) {
        let mut wb = XlsxWriter::new();
        let sheet = wb.add_worksheet();
        sheet.set_name("数据").unwrap();
        let bold = Format::new().set_bold();
        sheet.write_string_with_format(0, 0, "名称", &bold).unwrap();
        sheet.write_string(1, 0, "苹果").unwrap();
        sheet.write_number(1, 1, 12.5).unwrap();
        sheet.write_string(2, 0, "香蕉").unwrap();
        wb.save(path).expect("生成样本");
    }

    fn edit(sheet: &str, row: u32, col: u32, kind: &str, value: &str) -> CellEdit {
        CellEdit {
            sheet: sheet.to_string(),
            row,
            col,
            kind: kind.to_string(),
            value: value.to_string(),
        }
    }

    fn insert_rows(sheet: &str, at: u32, count: u32) -> StructureOp {
        StructureOp::InsertRows {
            sheet: sheet.to_string(),
            at,
            count,
        }
    }

    fn delete_rows(sheet: &str, at: u32, count: u32) -> StructureOp {
        StructureOp::DeleteRows {
            sheet: sheet.to_string(),
            at,
            count,
        }
    }

    /// 影子当前的序列化字节（撤销/重做逐步逐字节比对用）
    fn store_bytes(store: &ShadowStore, path: &Path) -> Vec<u8> {
        store
            .bytes_of(path.to_str().unwrap())
            .expect("应有影子")
            .bytes
    }

    /// 造一个「不脏」的影子（直接塞进仓库，用来验证淘汰策略）
    fn seed_clean(store: &ShadowStore, path: &Path) {
        let mut doc = build_shadow(path.to_str().unwrap()).expect("建立影子");
        doc.saved_pos = Some(doc.depth());
        store.insert_for_test(path.to_str().unwrap(), doc);
    }

    /// 键规范化：大小写与分隔符不同也认为是同一个文件（避免同一文件两份影子）
    #[test]
    fn shadow_key_normalizes_case_and_separators() {
        let dir = test_dir("key");
        let path = dir.join("Sample.XLSX");
        make_sample(&path);
        let a = shadow_key(path.to_str().unwrap());
        let b = shadow_key(&path.to_string_lossy().to_uppercase());
        let c = shadow_key(&path.to_string_lossy().replace('\\', "/"));
        assert_eq!(a, b, "大小写不同应是同一个 key");
        assert_eq!(a, c, "分隔符不同应是同一个 key");
        assert!(a.contains("sample.xlsx"), "key 应是小写规范形式：{a}");
    }

    /// 结构操作只改影子：磁盘字节一字不动、pending = true、影子内容已变
    #[test]
    fn structure_changes_shadow_only() {
        let dir = test_dir("shadow-only");
        let path = dir.join("a.xlsx");
        make_sample(&path);
        let store = ShadowStore::new(MAX_SHADOWS);
        let before = std::fs::read(&path).unwrap();

        let result = apply_op(&store, path.to_str().unwrap(), &insert_rows("数据", 1, 1), &[])
            .expect("插入行应成功");
        assert_eq!(result.saved, false, "结构操作不再直接落盘");
        assert!(result.pending, "有未落盘改动");
        assert_eq!(result.sheets.len(), 1);

        assert_eq!(std::fs::read(&path).unwrap(), before, "磁盘字节必须一字不动");
        assert!(store.is_pending(path.to_str().unwrap()));

        // 影子内容：A2 的「苹果」下移到 A3
        let bytes = store
            .bytes_of(path.to_str().unwrap())
            .expect("应有影子")
            .bytes;
        let mut wb = calamine::open_workbook_auto_from_rs(Cursor::new(bytes)).unwrap();
        use calamine::Reader;
        let range = wb.worksheet_range("数据").unwrap();
        assert_eq!(
            range.get_value((2, 0)).cloned(),
            Some(calamine::Data::String("苹果".into()))
        );
        assert_eq!(
            range.get_value((1, 0)).cloned().unwrap_or(calamine::Data::Empty),
            calamine::Data::Empty,
            "插入的空行"
        );
    }

    /// 空操作（越界插入）不留影子、不报 pending
    #[test]
    fn noop_structure_leaves_no_shadow() {
        let dir = test_dir("noop");
        let path = dir.join("a.xlsx");
        make_sample(&path);
        let store = ShadowStore::new(MAX_SHADOWS);

        let result = apply_op(&store, path.to_str().unwrap(), &insert_rows("数据", 100, 1), &[])
            .expect("越界插入应成功");
        assert!(!result.pending, "什么都没改，不该报 pending");
        assert_eq!(result.saved, false);
        assert!(!store.has(path.to_str().unwrap()), "不该留下影子");
        assert_eq!(store.len(), 0);
    }

    /// 失败的结构操作不建立影子，磁盘也不动
    #[test]
    fn failed_structure_leaves_nothing_behind() {
        let dir = test_dir("failed");
        let path = dir.join("a.xlsx");
        make_sample(&path);
        let store = ShadowStore::new(MAX_SHADOWS);
        let before = std::fs::read(&path).unwrap();

        let op = StructureOp::RenameSheet {
            sheet: "数据".to_string(),
            name: "数据".repeat(100),
        };
        let err = apply_op(&store, path.to_str().unwrap(), &op, &[]).expect_err("非法表名应报错");
        assert!(err.contains("过长"), "实际提示：{err}");
        assert!(!store.has(path.to_str().unwrap()), "失败不该留下影子");
        assert_eq!(std::fs::read(&path).unwrap(), before);
    }

    /// 保存：写回影子内容、`.bak`（显式要求时才生成）是操作前的磁盘字节；
    /// **保存后影子与撤销栈都保留**，只把磁盘基线挪到当前状态。
    #[test]
    fn save_lands_content_and_keeps_shadow_for_undo() {
        let dir = test_dir("save");
        let path = dir.join("a.xlsx");
        make_sample(&path);
        let store = ShadowStore::new(MAX_SHADOWS);
        let original = std::fs::read(&path).unwrap();

        apply_op(&store, path.to_str().unwrap(), &insert_rows("数据", 1, 1), &[]).unwrap();
        let saved = save(&store, path.to_str().unwrap(), None, &[], true).expect("保存应成功");
        assert_eq!(saved.saved_cells, 0);
        assert!(store.has(path.to_str().unwrap()), "保存后影子保留（撤销栈在里面）");
        assert!(!store.is_pending(path.to_str().unwrap()), "保存后不再是未落盘状态");
        assert_eq!(
            store.undo_flags(path.to_str().unwrap()),
            (true, false),
            "撤销栈必须保留（用户要求保存不影响撤回）"
        );
        assert_eq!(
            std::fs::read(dir.join("a.xlsx.bak")).unwrap(),
            original,
            ".bak 应是操作前的磁盘字节"
        );

        // 磁盘内容 = 影子内容
        let bytes = std::fs::read(&path).unwrap();
        let mut wb = calamine::open_workbook_auto_from_rs(Cursor::new(bytes.clone())).unwrap();
        use calamine::Reader;
        let range = wb.worksheet_range("数据").unwrap();
        assert_eq!(
            range.get_value((2, 0)).cloned(),
            Some(calamine::Data::String("苹果".into()))
        );
        drop(wb);

        // 保存之后仍能撤销：影子回到插入行之前，磁盘保持刚保存的内容
        let result = undo(&store, path.to_str().unwrap()).expect("保存后仍应能撤销");
        assert!(result.pending, "撤销后重新变脏");
        assert!(!result.can_undo, "只有一步");
        assert!(result.can_redo);
        assert_eq!(
            std::fs::read(&path).unwrap(),
            bytes,
            "撤销不写盘：磁盘还是刚保存的内容"
        );
        let shadow = store.bytes_of(path.to_str().unwrap()).unwrap().bytes;
        let mut wb = calamine::open_workbook_auto_from_rs(Cursor::new(shadow)).unwrap();
        let range = wb.worksheet_range("数据").unwrap();
        assert_eq!(
            range.get_value((1, 0)).cloned(),
            Some(calamine::Data::String("苹果".into())),
            "撤销后「苹果」回到第 2 行"
        );

        // 再重做：回到保存后的状态，且重新变干净
        let result = redo(&store, path.to_str().unwrap()).expect("应能重做");
        assert!(!result.pending, "重做到保存时的状态 → 不再脏");
        assert!(result.can_undo && !result.can_redo);
    }

    /// `.bak` 默认不生成（`save_requested(None)`）；显式要求才生成
    #[test]
    fn backup_is_optional_and_off_by_default() {
        let dir = test_dir("backup");
        let path = dir.join("a.xlsx");
        make_sample(&path);
        let store = ShadowStore::new(MAX_SHADOWS);
        let bak = dir.join("a.xlsx.bak");

        apply_op(&store, path.to_str().unwrap(), &insert_rows("数据", 1, 1), &[]).unwrap();
        let saved = save_requested(&store, path.to_str().unwrap(), None, &[], None)
            .expect("保存应成功");
        assert!(saved.backup_path.is_none(), "不传 backup 时不应备份");
        assert!(!bak.exists(), "默认不该生成 .bak");

        // 显式要求：生成，内容是「这批改动之前」的磁盘字节
        let before = std::fs::read(&path).unwrap();
        apply_op(&store, path.to_str().unwrap(), &insert_rows("数据", 0, 1), &[]).unwrap();
        let saved = save_requested(&store, path.to_str().unwrap(), None, &[], Some(true))
            .expect("保存应成功");
        assert_eq!(
            saved.backup_path.as_deref(),
            Some(bak.to_string_lossy().as_ref())
        );
        assert_eq!(std::fs::read(&bak).unwrap(), before);
        assert_ne!(std::fs::read(&path).unwrap(), before, "文件本身已更新");
    }

    /* ------------------------------ 撤销/重做 ------------------------------ */

    /// 删除行 → 撤销：内容完全恢复；再重做：又回到删除后的状态
    #[test]
    fn undo_restores_deleted_rows_and_redo_replays() {
        let dir = test_dir("undo-delete");
        let path = dir.join("a.xlsx");
        make_sample(&path);
        let store = ShadowStore::new(MAX_SHADOWS);

        // 先建立影子并退回「初始状态」，拿到与磁盘一致的那份影子字节
        apply_op(&store, path.to_str().unwrap(), &insert_rows("数据", 0, 0), &[]).unwrap();
        undo(&store, path.to_str().unwrap()).unwrap();
        let before = store_bytes(&store, &path);
        assert!(!store.is_pending(path.to_str().unwrap()), "退回初始状态 → 与磁盘一致");

        apply_op(&store, path.to_str().unwrap(), &delete_rows("数据", 1, 1), &[])
            .expect("删除行应成功");
        let after_delete = store_bytes(&store, &path);
        assert_ne!(before, after_delete, "删除确实改了内容");

        let result = undo(&store, path.to_str().unwrap()).expect("应能撤销");
        assert!(!result.pending, "撤销回到初始状态 → 与磁盘一致");
        assert!(!result.can_undo && result.can_redo);
        assert_eq!(store_bytes(&store, &path), before, "撤销应逐字节回到操作前");
        // 被删行的内容回来了（A2 原本是「苹果」）
        let mut wb = calamine::open_workbook_auto_from_rs(Cursor::new(before.clone())).unwrap();
        use calamine::Reader;
        assert_eq!(
            wb.worksheet_range("数据").unwrap().get_value((1, 0)).cloned(),
            Some(calamine::Data::String("苹果".into()))
        );
        drop(wb);

        let result = redo(&store, path.to_str().unwrap()).expect("应能重做");
        assert!(result.pending);
        assert!(result.can_undo && !result.can_redo);
        assert_eq!(store_bytes(&store, &path), after_delete);
        // 重做后被删行的位置变成「香蕉」（原来第 3 行）
        let mut wb =
            calamine::open_workbook_auto_from_rs(Cursor::new(after_delete.clone())).unwrap();
        assert_eq!(
            wb.worksheet_range("数据").unwrap().get_value((1, 0)).cloned(),
            Some(calamine::Data::String("香蕉".into()))
        );
    }

    /// 连续三次操作 → 连续三次撤销（逐步逐字节比对）→ 连续三次重做回到最终状态
    #[test]
    fn three_steps_walk_back_and_forth() {
        let dir = test_dir("undo-three");
        let path = dir.join("a.xlsx");
        make_sample(&path);
        let store = ShadowStore::new(MAX_SHADOWS);

        // 先建立影子再退回初始状态，拿到 s0
        apply_op(&store, path.to_str().unwrap(), &insert_rows("数据", 0, 0), &[]).unwrap();
        undo(&store, path.to_str().unwrap()).unwrap();
        let s0 = store_bytes(&store, &path);

        apply_op(&store, path.to_str().unwrap(), &insert_rows("数据", 1, 1), &[]).unwrap();
        let s1 = store_bytes(&store, &path);
        apply_op(&store, path.to_str().unwrap(), &insert_rows("数据", 0, 1), &[]).unwrap();
        let s2 = store_bytes(&store, &path);
        apply_op(&store, path.to_str().unwrap(), &delete_rows("数据", 2, 1), &[]).unwrap();
        let s3 = store_bytes(&store, &path);
        // 三步内容互不相同（否则测不出「逐步」）
        assert!(s0 != s1 && s1 != s2 && s2 != s3);

        let result = undo(&store, path.to_str().unwrap()).unwrap();
        assert_eq!(store_bytes(&store, &path), s2);
        assert!(result.can_undo && result.can_redo);
        undo(&store, path.to_str().unwrap()).unwrap();
        assert_eq!(store_bytes(&store, &path), s1);
        let result = undo(&store, path.to_str().unwrap()).unwrap();
        assert_eq!(store_bytes(&store, &path), s0);
        assert!(!result.can_undo, "已经退回最初状态");
        assert!(result.can_redo);
        assert!(!result.pending, "回到最初状态 → 与磁盘一致");

        // 三次重做回到最终状态
        redo(&store, path.to_str().unwrap()).unwrap();
        assert_eq!(store_bytes(&store, &path), s1);
        redo(&store, path.to_str().unwrap()).unwrap();
        assert_eq!(store_bytes(&store, &path), s2);
        let result = redo(&store, path.to_str().unwrap()).unwrap();
        assert_eq!(store_bytes(&store, &path), s3);
        assert!(result.can_undo && !result.can_redo);

        // 栈空时报错
        let err = redo(&store, path.to_str().unwrap()).expect_err("重做栈空应报错");
        assert!(err.contains("没有可重做"), "实际提示：{err}");
    }

    /// 撤销栈步数上限：35 次操作只能撤销 30 次
    #[test]
    fn undo_stack_is_capped_at_thirty_steps() {
        let dir = test_dir("undo-cap");
        let path = dir.join("a.xlsx");
        make_sample(&path);
        let store = ShadowStore::new(MAX_SHADOWS);

        for i in 0..35 {
            apply_op(&store, path.to_str().unwrap(), &insert_rows("数据", 0, 1), &[])
                .unwrap_or_else(|e| panic!("第 {i} 次插入应成功：{e}"));
        }
        assert_eq!(
            store.undo_len(path.to_str().unwrap()),
            Some(MAX_UNDO_STEPS),
            "栈里只应留最近 {MAX_UNDO_STEPS} 步"
        );

        for i in 0..MAX_UNDO_STEPS {
            undo(&store, path.to_str().unwrap())
                .unwrap_or_else(|e| panic!("第 {} 次撤销应成功：{e}", i + 1));
        }
        let err = undo(&store, path.to_str().unwrap()).expect_err("第 31 次撤销应报错");
        assert_eq!(err, "没有可撤销的改动");
        // 最旧的 5 步已经回不去了：剩下的内容仍是「前 5 次插入之后」的状态
        assert!(store.is_pending(path.to_str().unwrap()));
    }

    /// 快照裁剪：字节上限优先于步数上限，且至少保留最近一步
    #[test]
    fn snapshot_trimming_prefers_byte_budget() {
        // 步数上限宽松、总字节上限 100 → 5 条 40 字节只留最近 2 条（40+40=80 ≤ 100）
        let mut stack: Vec<Vec<u8>> = (0..5).map(|_| vec![0u8; 40]).collect();
        assert_eq!(trim_snapshots(&mut stack, 30, 100), 3, "丢掉最旧的 3 条");
        assert_eq!(stack.len(), 2);

        // 单条就超预算时也要保留最近一步（否则大表完全无法撤销）
        let mut huge: Vec<Vec<u8>> = (0..2).map(|_| vec![0u8; 10_000]).collect();
        trim_snapshots(&mut huge, 30, 100);
        assert_eq!(huge.len(), 1);

        // 步数上限生效（字节不设限）
        let mut steps: Vec<Vec<u8>> = (0..10).map(|_| vec![0u8; 1]).collect();
        assert_eq!(trim_snapshots(&mut steps, 3, usize::MAX), 7);
        assert_eq!(steps.len(), 3);

        // 都没超：一个都不丢
        let mut few: Vec<Vec<u8>> = (0..2).map(|_| vec![0u8; 8]).collect();
        assert_eq!(trim_snapshots(&mut few, 30, 512), 0);
        assert_eq!(few.len(), 2);
    }

    /// 空操作（越界插入）不压栈：没有可撤销的步骤
    #[test]
    fn noop_structure_pushes_no_undo_step() {
        let dir = test_dir("undo-noop");
        let path = dir.join("a.xlsx");
        make_sample(&path);
        let store = ShadowStore::new(MAX_SHADOWS);

        apply_op(&store, path.to_str().unwrap(), &insert_rows("数据", 100, 1), &[]).unwrap();
        assert_eq!(store.undo_flags(path.to_str().unwrap()), (false, false));
        let err = undo(&store, path.to_str().unwrap()).expect_err("没有可撤销的改动");
        assert_eq!(err, "没有可撤销的改动");
        // 连影子都没有（什么都没改）
        assert!(!store.has(path.to_str().unwrap()));
    }

    /// 丢弃影子同时清空撤销栈
    #[test]
    fn discard_clears_undo_stack() {
        let dir = test_dir("undo-discard");
        let path = dir.join("a.xlsx");
        make_sample(&path);
        let store = ShadowStore::new(MAX_SHADOWS);

        apply_op(&store, path.to_str().unwrap(), &insert_rows("数据", 1, 1), &[]).unwrap();
        assert_eq!(store.undo_flags(path.to_str().unwrap()), (true, false));
        assert!(store.discard(path.to_str().unwrap()), "丢掉了影子");
        assert_eq!(store.undo_flags(path.to_str().unwrap()), (false, false));
        assert_eq!(
            undo(&store, path.to_str().unwrap()).expect_err("丢弃后没有可撤销的"),
            "没有可撤销的改动"
        );
    }

    /// 撤销/重做只改内存：磁盘始终不动
    #[test]
    fn undo_never_touches_disk() {
        let dir = test_dir("undo-disk");
        let path = dir.join("a.xlsx");
        make_sample(&path);
        let store = ShadowStore::new(MAX_SHADOWS);
        let disk = std::fs::read(&path).unwrap();

        apply_op(&store, path.to_str().unwrap(), &insert_rows("数据", 1, 1), &[]).unwrap();
        undo(&store, path.to_str().unwrap()).unwrap();
        redo(&store, path.to_str().unwrap()).unwrap();
        undo(&store, path.to_str().unwrap()).unwrap();
        assert_eq!(std::fs::read(&path).unwrap(), disk, "撤销/重做不碰磁盘");
    }

    /// 保存后影子会保留，所以必须有安全网：**干净影子**遇到外部改动时自动作废（回落磁盘），
    /// 脏影子不动（那是用户还没保存的改动，只能由 discard 或保存处理）。
    #[test]
    fn clean_shadow_is_dropped_when_file_changes_externally() {
        let dir = test_dir("external-change");
        let path = dir.join("a.xlsx");
        make_sample(&path);
        let store = ShadowStore::new(MAX_SHADOWS);

        // 脏影子：外部改动不作废
        apply_op(&store, path.to_str().unwrap(), &insert_rows("数据", 1, 1), &[]).unwrap();
        std::fs::write(&path, b"external change").unwrap();
        assert!(store.is_pending(path.to_str().unwrap()));
        assert!(
            store.bytes_of(path.to_str().unwrap()).is_some(),
            "脏影子不该被外部改动作废"
        );

        // 保存后（干净影子）：再被外部改动一次 → 读路径自动作废
        let other = dir.join("b.xlsx");
        make_sample(&other);
        let store2 = ShadowStore::new(MAX_SHADOWS);
        apply_op(&store2, other.to_str().unwrap(), &insert_rows("数据", 1, 1), &[]).unwrap();
        save(&store2, other.to_str().unwrap(), None, &[], false).unwrap();
        assert!(!store2.is_pending(other.to_str().unwrap()));
        assert!(store2.bytes_of(other.to_str().unwrap()).is_some(), "刚保存完读影子");

        let mut changed = std::fs::read(&other).unwrap();
        changed.extend_from_slice(b"external");
        std::fs::write(&other, &changed).unwrap();
        assert!(
            store2.bytes_of(other.to_str().unwrap()).is_none(),
            "干净影子遇到外部改动应自动作废"
        );
        assert!(!store2.has(other.to_str().unwrap()), "作废后影子没了");
    }

    /// 保存时带上的 edits 也只在保存那一刻落盘
    #[test]
    fn edits_land_only_on_save() {
        let dir = test_dir("edits");
        let path = dir.join("a.xlsx");
        make_sample(&path);
        let store = ShadowStore::new(MAX_SHADOWS);
        let before = std::fs::read(&path).unwrap();
        let edits = [edit("数据", 0, 0, "text", "标题")];

        apply_op(
            &store,
            path.to_str().unwrap(),
            &insert_rows("数据", 1, 0),
            &edits,
        )
        .unwrap();
        assert_eq!(std::fs::read(&path).unwrap(), before, "编辑也不该立刻落盘");

        save(&store, path.to_str().unwrap(), None, &edits, false).expect("保存应成功");
        let bytes = std::fs::read(&path).unwrap();
        let mut wb = calamine::open_workbook_auto_from_rs(Cursor::new(bytes)).unwrap();
        use calamine::Reader;
        let range = wb.worksheet_range("数据").unwrap();
        assert_eq!(
            range.get_value((0, 0)).cloned(),
            Some(calamine::Data::String("标题".into()))
        );
    }

    /// 丢弃影子：读回磁盘内容，pending 归 false
    #[test]
    fn discard_restores_disk_view() {
        let dir = test_dir("discard");
        let path = dir.join("a.xlsx");
        make_sample(&path);
        let store = ShadowStore::new(MAX_SHADOWS);

        apply_op(&store, path.to_str().unwrap(), &insert_rows("数据", 1, 1), &[]).unwrap();
        assert!(store.is_pending(path.to_str().unwrap()));
        assert!(store.discard(path.to_str().unwrap()), "确实丢掉了东西");
        assert!(!store.is_pending(path.to_str().unwrap()));
        assert!(store.bytes_of(path.to_str().unwrap()).is_none());
        assert!(!store.discard(path.to_str().unwrap()), "再丢一次就是 false");
    }

    /// 容量：前 N 个「不脏」时可淘汰最久未用者并继续操作；全都脏则明确报错
    #[test]
    fn capacity_evicts_clean_lru_and_reports_when_all_dirty() {
        let dir = test_dir("capacity");
        let store = ShadowStore::new(2);
        let paths: Vec<std::path::PathBuf> = (0..3)
            .map(|i| {
                let p = dir.join(format!("s{i}.xlsx"));
                make_sample(&p);
                p
            })
            .collect();

        // 两个「不脏」的影子占满容量
        seed_clean(&store, &paths[0]);
        seed_clean(&store, &paths[1]);
        assert_eq!(store.len(), 2);

        // 第三个：淘汰最久未用的 s0，操作正常完成
        let result = apply_op(
            &store,
            paths[2].to_str().unwrap(),
            &insert_rows("数据", 1, 1),
            &[],
        )
        .expect("淘汰不脏条目后应能继续操作");
        assert!(result.pending);
        assert_eq!(store.len(), 2, "上限仍是 2");
        assert!(!store.has(paths[0].to_str().unwrap()), "最久未用者被淘汰");
        assert!(store.has(paths[1].to_str().unwrap()));

        // 现在把 s1 也改脏：两个都是脏的了
        apply_op(
            &store,
            paths[1].to_str().unwrap(),
            &insert_rows("数据", 1, 1),
            &[],
        )
        .expect("同一文件继续操作不该被容量拦住");
        assert_eq!(store.len(), 2);

        // 再来一个新文件：全都有未落盘改动 → 明确报错，且不丢任何人的改动
        let third = dir.join("s3.xlsx");
        make_sample(&third);
        let err = apply_op(
            &store,
            third.to_str().unwrap(),
            &insert_rows("数据", 1, 1),
            &[],
        )
        .expect_err("全都有未落盘改动时应报错");
        assert!(err.contains("过多"), "实际提示：{err}");
        assert!(!store.has(third.to_str().unwrap()));
        assert!(store.has(paths[1].to_str().unwrap()) && store.has(paths[2].to_str().unwrap()));
        assert_eq!(store.len(), 2, "报错时不该动已有影子");
    }

    /// 同一文件重复操作只占一个位置（键规范化生效）
    #[test]
    fn same_file_uses_one_slot() {
        let dir = test_dir("one-slot");
        let path = dir.join("a.xlsx");
        make_sample(&path);
        let store = ShadowStore::new(2);

        for _ in 0..3 {
            apply_op(&store, path.to_str().unwrap(), &insert_rows("数据", 1, 1), &[]).unwrap();
        }
        assert_eq!(store.len(), 1, "同一文件只应有一个影子");
    }

    /// 另存为：写到别的文件，原文件仍有未落盘的改动 → 影子保留且仍是脏的
    #[test]
    fn save_as_keeps_the_shadow_pending() {
        let dir = test_dir("save-as");
        let path = dir.join("a.xlsx");
        make_sample(&path);
        let store = ShadowStore::new(MAX_SHADOWS);
        let original = std::fs::read(&path).unwrap();
        let target = dir.join("copy.xlsx");

        apply_op(&store, path.to_str().unwrap(), &insert_rows("数据", 1, 1), &[]).unwrap();
        let saved = save(
            &store,
            path.to_str().unwrap(),
            Some(target.to_str().unwrap()),
            &[],
            true,
        )
        .expect("另存为应成功");
        assert_eq!(saved.path, target.to_string_lossy());

        // 目标文件有影子内容
        let mut wb = calamine::open_workbook_auto_from_rs(Cursor::new(
            std::fs::read(&target).unwrap(),
        ))
        .unwrap();
        use calamine::Reader;
        assert_eq!(
            wb.worksheet_range("数据").unwrap().get_value((2, 0)).cloned(),
            Some(calamine::Data::String("苹果".into()))
        );
        // 原文件没动，影子还在（原文件仍有未落盘改动）
        assert_eq!(std::fs::read(&path).unwrap(), original);
        assert!(store.is_pending(path.to_str().unwrap()));
        assert_eq!(
            std::fs::read(dir.join("a.xlsx.bak")).unwrap(),
            original,
            "另存为备份的是源文件"
        );
    }

    /// 保存失败（编辑非法）时影子放回：结构改动不丢，重试仍能写出
    #[test]
    fn failed_save_keeps_the_shadow() {
        let dir = test_dir("save-fail");
        let path = dir.join("a.xlsx");
        make_sample(&path);
        let store = ShadowStore::new(MAX_SHADOWS);
        let disk = std::fs::read(&path).unwrap();

        apply_op(&store, path.to_str().unwrap(), &insert_rows("数据", 1, 1), &[]).unwrap();
        let bad = [edit("数据", 0, 0, "number", "不是数字")];
        let err = save(&store, path.to_str().unwrap(), None, &bad, true)
            .expect_err("非法编辑应让保存失败");
        assert!(err.contains("不是数字"), "实际提示：{err}");
        assert_eq!(std::fs::read(&path).unwrap(), disk, "失败时磁盘不得改变");
        assert!(
            store.is_pending(path.to_str().unwrap()),
            "结构改动必须还在影子里"
        );

        // 重新保存（这次不带非法编辑）：结构改动照常落盘
        save(&store, path.to_str().unwrap(), None, &[], false).expect("重试应成功");
        let mut wb = calamine::open_workbook_auto_from_rs(Cursor::new(
            std::fs::read(&path).unwrap(),
        ))
        .unwrap();
        use calamine::Reader;
        assert_eq!(
            wb.worksheet_range("数据").unwrap().get_value((2, 0)).cloned(),
            Some(calamine::Data::String("苹果".into()))
        );
    }
}
