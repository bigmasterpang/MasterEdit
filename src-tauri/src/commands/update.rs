//! 内置在线更新：查询软件中心（master.dapang.wang / 国内加速节点）并下载安装包。
//!
//! 说明：
//! - 软件中心 API 未开启 CORS，WebView 内无法直接 fetch，因此统一走 WinHTTP；
//! - 不引入第三方 HTTP 库，保持安装包体积；
//! - 下载过程中向前端推送 `update-progress` 事件，并实时计算 SHA-256 供校验。

#[cfg(windows)]
use std::ffi::c_void;
#[cfg(windows)]
use std::io::Write;
#[cfg(windows)]
use std::sync::atomic::{AtomicU64, Ordering};

#[cfg(windows)]
use serde::Serialize;
#[cfg(windows)]
use tauri::{AppHandle, Emitter};
#[cfg(windows)]
use windows::core::PCWSTR;
#[cfg(windows)]
use windows::Win32::Networking::WinHttp::*;

/// 软件中心节点（主站优先，失败回退国内纯 IP 节点）
pub const PORTAL_NODES: [&str; 2] = ["https://master.dapang.wang", "http://106.14.225.57"];
/// 软件中心中的应用标识与平台
pub const APP_ID: &str = "masteredit";
pub const PLATFORM: &str = "windows";

/* ------------------------------------------------------------------ */
/* WinHTTP 封装                                                        */
/* ------------------------------------------------------------------ */

#[cfg(windows)]
struct Handle(*mut c_void);

#[cfg(windows)]
impl Drop for Handle {
    fn drop(&mut self) {
        if !self.0.is_null() {
            unsafe {
                let _ = WinHttpCloseHandle(self.0);
            }
        }
    }
}

#[cfg(windows)]
fn wide(text: &str) -> Vec<u16> {
    text.encode_utf16().chain(std::iter::once(0)).collect()
}

#[cfg(windows)]
fn wrap(handle: *mut c_void, what: &str) -> Result<Handle, String> {
    if handle.is_null() {
        let code = unsafe { windows::Win32::Foundation::GetLastError().0 };
        Err(format!("{what} 失败（Win32 错误 {code}）"))
    } else {
        Ok(Handle(handle))
    }
}

/// 解析 URL 为 (是否 HTTPS, 主机, 端口, 路径)
#[cfg(windows)]
fn parse_url(url: &str) -> Result<(bool, String, u16, String), String> {
    let (scheme, rest) = url.split_once("://").ok_or("URL 缺少协议头")?;
    let secure = match scheme.to_ascii_lowercase().as_str() {
        "https" => true,
        "http" => false,
        other => return Err(format!("不支持的协议: {other}")),
    };
    let (host_port, path) = match rest.find('/') {
        Some(index) => (&rest[..index], &rest[index..]),
        None => (rest, "/"),
    };
    let (host, port) = match host_port.rsplit_once(':') {
        Some((h, p)) => (
            h.to_string(),
            p.parse::<u16>().map_err(|_| "端口号无效".to_string())?,
        ),
        None => (host_port.to_string(), if secure { 443 } else { 80 }),
    };
    if host.is_empty() {
        return Err("URL 缺少主机名".to_string());
    }
    Ok((secure, host, port, path.to_string()))
}

/// 网络代理模式
#[cfg(windows)]
#[derive(Clone, Copy, PartialEq)]
pub enum ProxyMode {
    /// 跟随系统代理（默认）
    System,
    /// 直连，忽略系统代理
    Direct,
}

#[cfg(windows)]
pub struct HttpResponse {
    pub status: u32,
    pub body: Vec<u8>,
}

/// 执行一次 GET 请求，`on_chunk` 返回 false 可提前中断；
/// `collect` 为 false 时只交给回调处理，不保留响应体（下载大文件时用）。
#[cfg(windows)]
pub fn http_get<F>(
    url: &str,
    timeout_ms: i32,
    proxy: ProxyMode,
    collect: bool,
    mut on_chunk: F,
) -> Result<HttpResponse, String>
where
    F: FnMut(&[u8]) -> bool,
{
    let (secure, host, port, path) = parse_url(url)?;

    let agent = wide("MasterEdit-Updater/1.0");
    let access_type = match proxy {
        ProxyMode::System => WINHTTP_ACCESS_TYPE_DEFAULT_PROXY,
        ProxyMode::Direct => WINHTTP_ACCESS_TYPE_NO_PROXY,
    };
    let session = wrap(
        unsafe {
            WinHttpOpen(
                PCWSTR(agent.as_ptr()),
                access_type,
                PCWSTR::null(),
                PCWSTR::null(),
                0,
            )
        },
        "WinHttpOpen",
    )?;

    unsafe {
        let _ = WinHttpSetTimeouts(session.0, timeout_ms, timeout_ms, timeout_ms, timeout_ms);
    }

    let host_w = wide(&host);
    let connect = wrap(
        unsafe { WinHttpConnect(session.0, PCWSTR(host_w.as_ptr()), port, 0) },
        "WinHttpConnect",
    )?;

    let verb = wide("GET");
    let path_w = wide(&path);
    let flags = if secure {
        WINHTTP_FLAG_SECURE
    } else {
        WINHTTP_OPEN_REQUEST_FLAGS(0)
    };
    let request = wrap(
        unsafe {
            WinHttpOpenRequest(
                connect.0,
                PCWSTR(verb.as_ptr()),
                PCWSTR(path_w.as_ptr()),
                PCWSTR::null(),
                PCWSTR::null(),
                std::ptr::null(),
                flags,
            )
        },
        "WinHttpOpenRequest",
    )?;

    // 证书校验交给 WinHTTP 默认行为处理：旧实现曾用 WinHttpSetOption 打开
    // SECURITY_FLAG_IGNORE_* 一味跳过校验（未知 CA / 证书过期 / 域名不匹配全放过），
    // 而更新元数据（版本号 + SHA-256）与安装包都来自同一响应，
    // 关掉校验等于让中间人同时伪造两者，静默安装任意程序。
    // 当前软件中心证书链正常，保持严格校验。
    if secure {
        let _ = &request;
    }

    unsafe {
        WinHttpSendRequest(request.0, None, None, 0, 0, 0)
            .map_err(|e| format!("发送请求失败: {e}"))?;
        WinHttpReceiveResponse(request.0, std::ptr::null_mut())
            .map_err(|e| format!("接收响应失败: {e}"))?;
    }

    // 状态码
    let mut status: u32 = 0;
    let mut status_len = std::mem::size_of::<u32>() as u32;
    unsafe {
        let _ = WinHttpQueryHeaders(
            request.0,
            WINHTTP_QUERY_STATUS_CODE | WINHTTP_QUERY_FLAG_NUMBER,
            PCWSTR::null(),
            Some(&mut status as *mut u32 as *mut c_void),
            &mut status_len,
            std::ptr::null_mut(),
        );
    }

    // 响应体（分块读取）
    let mut body: Vec<u8> = Vec::new();
    let mut buffer = vec![0u8; 64 * 1024];
    loop {
        let mut read: u32 = 0;
        let ok = unsafe {
            WinHttpReadData(
                request.0,
                buffer.as_mut_ptr() as *mut c_void,
                buffer.len() as u32,
                &mut read,
            )
        };
        if let Err(error) = ok {
            // 读取失败不能当成正常结束：否则半截安装包会被当作下载成功并被执行
            return Err(format!("读取响应数据失败: {error}"));
        }
        if read == 0 {
            break;
        }
        let chunk = &buffer[..read as usize];
        if !on_chunk(chunk) {
            break;
        }
        if collect {
            body.extend_from_slice(chunk);
        }
    }

    Ok(HttpResponse { status, body })
}

/* ------------------------------------------------------------------ */
/* SHA-256（纯 Rust 实现，避免额外依赖）                                */
/* ------------------------------------------------------------------ */

const SHA256_K: [u32; 64] = [
    0x428a2f98, 0x71374491, 0xb5c0fbcf, 0xe9b5dba5, 0x3956c25b, 0x59f111f1, 0x923f82a4, 0xab1c5ed5,
    0xd807aa98, 0x12835b01, 0x243185be, 0x550c7dc3, 0x72be5d74, 0x80deb1fe, 0x9bdc06a7, 0xc19bf174,
    0xe49b69c1, 0xefbe4786, 0x0fc19dc6, 0x240ca1cc, 0x2de92c6f, 0x4a7484aa, 0x5cb0a9dc, 0x76f988da,
    0x983e5152, 0xa831c66d, 0xb00327c8, 0xbf597fc7, 0xc6e00bf3, 0xd5a79147, 0x06ca6351, 0x14292967,
    0x27b70a85, 0x2e1b2138, 0x4d2c6dfc, 0x53380d13, 0x650a7354, 0x766a0abb, 0x81c2c92e, 0x92722c85,
    0xa2bfe8a1, 0xa81a664b, 0xc24b8b70, 0xc76c51a3, 0xd192e819, 0xd6990624, 0xf40e3585, 0x106aa070,
    0x19a4c116, 0x1e376c08, 0x2748774c, 0x34b0bcb5, 0x391c0cb3, 0x4ed8aa4a, 0x5b9cca4f, 0x682e6ff3,
    0x748f82ee, 0x78a5636f, 0x84c87814, 0x8cc70208, 0x90befffa, 0xa4506ceb, 0xbef9a3f7, 0xc67178f2,
];

/// 增量式 SHA-256：下载过程中逐块喂入，避免二次读盘
pub struct Sha256 {
    state: [u32; 8],
    buffer: Vec<u8>,
    length: u64,
}

impl Sha256 {
    pub fn new() -> Self {
        Self {
            state: [
                0x6a09e667, 0xbb67ae85, 0x3c6ef372, 0xa54ff53a, 0x510e527f, 0x9b05688c, 0x1f83d9ab,
                0x5be0cd19,
            ],
            buffer: Vec::with_capacity(64),
            length: 0,
        }
    }

    fn compress(&mut self, block: &[u8]) {
        let mut w = [0u32; 64];
        for i in 0..16 {
            w[i] = u32::from_be_bytes([
                block[i * 4],
                block[i * 4 + 1],
                block[i * 4 + 2],
                block[i * 4 + 3],
            ]);
        }
        for i in 16..64 {
            let s0 = w[i - 15].rotate_right(7) ^ w[i - 15].rotate_right(18) ^ (w[i - 15] >> 3);
            let s1 = w[i - 2].rotate_right(17) ^ w[i - 2].rotate_right(19) ^ (w[i - 2] >> 10);
            w[i] = w[i - 16]
                .wrapping_add(s0)
                .wrapping_add(w[i - 7])
                .wrapping_add(s1);
        }
        let [mut a, mut b, mut c, mut d, mut e, mut f, mut g, mut h] = self.state;
        for i in 0..64 {
            let s1 = e.rotate_right(6) ^ e.rotate_right(11) ^ e.rotate_right(25);
            let ch = (e & f) ^ ((!e) & g);
            let temp1 = h
                .wrapping_add(s1)
                .wrapping_add(ch)
                .wrapping_add(SHA256_K[i])
                .wrapping_add(w[i]);
            let s0 = a.rotate_right(2) ^ a.rotate_right(13) ^ a.rotate_right(22);
            let maj = (a & b) ^ (a & c) ^ (b & c);
            let temp2 = s0.wrapping_add(maj);
            h = g;
            g = f;
            f = e;
            e = d.wrapping_add(temp1);
            d = c;
            c = b;
            b = a;
            a = temp1.wrapping_add(temp2);
        }
        self.state[0] = self.state[0].wrapping_add(a);
        self.state[1] = self.state[1].wrapping_add(b);
        self.state[2] = self.state[2].wrapping_add(c);
        self.state[3] = self.state[3].wrapping_add(d);
        self.state[4] = self.state[4].wrapping_add(e);
        self.state[5] = self.state[5].wrapping_add(f);
        self.state[6] = self.state[6].wrapping_add(g);
        self.state[7] = self.state[7].wrapping_add(h);
    }

    pub fn update(&mut self, data: &[u8]) {
        self.length += data.len() as u64;
        let mut offset = 0;
        while offset < data.len() {
            let need = 64 - self.buffer.len();
            let take = need.min(data.len() - offset);
            self.buffer.extend_from_slice(&data[offset..offset + take]);
            offset += take;
            if self.buffer.len() == 64 {
                let block = std::mem::take(&mut self.buffer);
                self.compress(&block);
            }
        }
    }

    pub fn finish_hex(mut self) -> String {
        let bit_len = self.length * 8;
        self.buffer.push(0x80);
        if self.buffer.len() > 56 {
            while self.buffer.len() < 64 {
                self.buffer.push(0);
            }
            let block = std::mem::take(&mut self.buffer);
            self.compress(&block);
        }
        while self.buffer.len() < 56 {
            self.buffer.push(0);
        }
        self.buffer.extend_from_slice(&bit_len.to_be_bytes());
        let block = std::mem::take(&mut self.buffer);
        self.compress(&block);

        self.state
            .iter()
            .map(|word| format!("{word:08x}"))
            .collect::<String>()
    }
}

/* ------------------------------------------------------------------ */
/* 软件中心 API                                                        */
/* ------------------------------------------------------------------ */

#[derive(Serialize, Clone)]
#[serde(rename_all = "camelCase")]
pub struct PortalRelease {
    pub base_url: String,
    pub version: String,
    pub filename: String,
    pub size: u64,
    pub human_size: String,
    pub sha256: String,
    pub uploaded_at: String,
    pub release_notes: String,
    pub download_url: String,
    pub has_update: bool,
    pub current_version: String,
    /// 当前程序类型：portable（绿色版）/ installed（安装版）
    pub install_kind: String,
}

/// 判断当前是否为「安装版」：
/// 1. 安装目录存在安装标记文件（由 NSIS 安装钩子写入）；
/// 2. 或同目录存在 uninstall.exe（NSIS 卸载程序）。
/// 绿色版（便携版）没有这些文件，因此走自替换升级。
#[cfg(windows)]
pub fn is_installed_build() -> bool {
    let Ok(current) = std::env::current_exe() else {
        return false;
    };
    let Some(dir) = current.parent() else {
        return false;
    };
    dir.join("installed.marker").exists() || dir.join("uninstall.exe").exists()
}

#[cfg(windows)]
fn install_kind() -> &'static str {
    if is_installed_build() {
        "installer"
    } else {
        "portable"
    }
}

fn version_parts(value: &str) -> Vec<u64> {
    value
        .trim()
        .trim_start_matches(['v', 'V'])
        .split(['.', '-', '+'])
        .map(|part| part.parse::<u64>().unwrap_or(0))
        .collect()
}

/// 语义化版本比较：remote 比 local 新则返回 true
pub fn is_newer(remote: &str, local: &str) -> bool {
    let r = version_parts(remote);
    let l = version_parts(local);
    let len = r.len().max(l.len());
    for i in 0..len {
        let rv = r.get(i).copied().unwrap_or(0);
        let lv = l.get(i).copied().unwrap_or(0);
        if rv > lv {
            return true;
        }
        if rv < lv {
            return false;
        }
    }
    false
}

/// 按节点顺序尝试请求并解析 JSON
fn fetch_latest(current: &str, variant: &str) -> Result<PortalRelease, String> {
    let mut last_error = String::from("所有更新服务器均不可用");
    for base in PORTAL_NODES {
        let url = format!("{base}/api/apps/{APP_ID}/{PLATFORM}/latest?variant={variant}");
        for proxy in [ProxyMode::System, ProxyMode::Direct] {
            match http_get(&url, 4000, proxy, true, |_| true) {
                Ok(response) if response.status == 200 => {
                    let text = String::from_utf8_lossy(&response.body).to_string();
                    let json: serde_json::Value = serde_json::from_str(&text)
                        .map_err(|e| format!("响应解析失败: {e}"))?;
                    let version = json
                        .get("version")
                        .and_then(|v| v.as_str())
                        .unwrap_or_default()
                        .to_string();
                    if version.is_empty() {
                        last_error = "服务器未返回版本号".to_string();
                        continue;
                    }
                    let download_path = json
                        .get("download_url")
                        .and_then(|v| v.as_str())
                        .unwrap_or_default()
                        .to_string();
                    return Ok(PortalRelease {
                        base_url: base.to_string(),
                        has_update: is_newer(&version, current),
                        version,
                        filename: json
                            .get("filename")
                            .and_then(|v| v.as_str())
                            .unwrap_or_default()
                            .to_string(),
                        size: json.get("size").and_then(|v| v.as_u64()).unwrap_or(0),
                        human_size: json
                            .get("human_size")
                            .and_then(|v| v.as_str())
                            .unwrap_or_default()
                            .to_string(),
                        sha256: json
                            .get("sha256")
                            .and_then(|v| v.as_str())
                            .unwrap_or_default()
                            .to_string(),
                        uploaded_at: json
                            .get("uploaded_at")
                            .and_then(|v| v.as_str())
                            .unwrap_or_default()
                            .to_string(),
                        release_notes: json
                            .get("release_notes")
                            .and_then(|v| v.as_str())
                            .unwrap_or_default()
                            .to_string(),
                        download_url: format!("{base}{download_path}"),
                        current_version: current.to_string(),
                        install_kind: install_kind().to_string(),
                    });
                }
                Ok(response) if response.status == 404 => {
                    last_error = "软件中心暂无 MasterEdit 的发布版本".to_string();
                }
                Ok(response) => {
                    last_error = format!("服务器返回状态码 {}", response.status);
                }
                Err(error) => {
                    last_error = error;
                }
            }
        }
    }
    Err(last_error)
}

/* ------------------------------------------------------------------ */
/* Tauri 命令                                                          */
/* ------------------------------------------------------------------ */

/// 查询软件中心最新版本
#[cfg(windows)]
#[tauri::command]
pub async fn check_update(app: AppHandle) -> Result<PortalRelease, String> {
    let current = app.package_info().version.to_string();
    let variant = install_kind().to_string();
    tauri::async_runtime::spawn_blocking(move || fetch_latest(&current, &variant))
        .await
        .map_err(|e| format!("检查更新失败: {e}"))?
}

#[cfg(windows)]
#[derive(Serialize, Clone)]
#[serde(rename_all = "camelCase")]
pub struct DownloadResult {
    pub path: String,
    pub size: u64,
    pub sha256: String,
}

#[cfg(windows)]
#[derive(Serialize, Clone)]
#[serde(rename_all = "camelCase")]
struct ProgressPayload {
    received: u64,
    total: u64,
}

/// 下载安装包到临时目录，边下边算 SHA-256，并推送进度事件
#[cfg(windows)]
#[tauri::command]
pub async fn download_update(
    app: AppHandle,
    url: String,
    dest: String,
    expect_sha256: Option<String>,
    expect_size: Option<u64>,
) -> Result<DownloadResult, String> {
    let dest_for_task = dest.clone();
    let app_for_task = app.clone();
    let expected = expect_sha256.unwrap_or_default().to_lowercase();
    let total_hint = expect_size.unwrap_or(0);

    let result = tauri::async_runtime::spawn_blocking(move || -> Result<DownloadResult, String> {
        let received = AtomicU64::new(0);
        let mut hasher = Sha256::new();
        let mut file = std::fs::File::create(&dest_for_task)
            .map_err(|e| format!("无法创建临时文件: {e}"))?;

        let mut total = total_hint;
        let mut last_emit = std::time::Instant::now();
        // 写盘失败必须冒泡：否则会留下半截文件却报告下载成功
        let mut write_error: Option<String> = None;

        let response = http_get(&url, 8000, ProxyMode::System, false, |chunk| {
            if let Err(error) = file.write_all(chunk) {
                write_error = Some(error.to_string());
                return false;
            }
            hasher.update(chunk);
            let now = received.fetch_add(chunk.len() as u64, Ordering::Relaxed) + chunk.len() as u64;
            if now > total {
                total = now;
            }
            // 每 120ms 推送一次进度，避免事件风暴
            if last_emit.elapsed().as_millis() >= 120 {
                last_emit = std::time::Instant::now();
                let _ = app_for_task.emit(
                    "update-progress",
                    ProgressPayload {
                        received: now,
                        total,
                    },
                );
            }
            true
        })?;

        if let Some(error) = write_error {
            let _ = std::fs::remove_file(&dest_for_task);
            return Err(format!("写入安装包失败: {error}"));
        }

        if response.status != 200 {
            let _ = std::fs::remove_file(&dest_for_task);
            return Err(format!("下载失败，服务器返回状态码 {}", response.status));
        }

        let size = received.load(Ordering::Relaxed);
        let sha256 = hasher.finish_hex();
        let _ = app_for_task.emit(
            "update-progress",
            ProgressPayload {
                received: size,
                total: if total > 0 { total } else { size },
            },
        );

        // 大小不符说明传输被截断（校验值缺失时这是唯一的完整性兜底）
        if total_hint > 0 && size != total_hint {
            let _ = std::fs::remove_file(&dest_for_task);
            return Err(format!(
                "安装包不完整（期望 {total_hint} 字节，实际 {size} 字节），请重试"
            ));
        }

        if !expected.is_empty() && expected != sha256 {
            let _ = std::fs::remove_file(&dest_for_task);
            return Err(format!(
                "安装包校验失败（SHA-256 不匹配）\n期望: {expected}\n实际: {sha256}"
            ));
        }

        Ok(DownloadResult {
            path: dest_for_task,
            size,
            sha256,
        })
    })
    .await
    .map_err(|e| format!("下载任务失败: {e}"))?;

    result
}

/// 运行下载好的安装程序（NSIS 安装包模式）
#[cfg(windows)]
#[tauri::command]
pub async fn run_installer(path: String) -> Result<(), String> {
    if !std::path::Path::new(&path).exists() {
        return Err("安装包不存在，请重新下载".to_string());
    }
    std::process::Command::new(&path)
        .spawn()
        .map_err(|e| format!("启动安装包失败: {e}"))?;
    Ok(())
}

/// 便携版自更新：用下载好的新版本替换当前可执行文件并重启。
/// Windows 允许重命名正在运行的 exe，因此流程为：
/// 当前 exe → 改名为 .old → 新 exe 就位 → 启动新版本 → 前端退出。
#[cfg(windows)]
#[tauri::command]
pub async fn apply_update(path: String) -> Result<(), String> {
    let source = std::path::PathBuf::from(&path);
    if !source.exists() {
        return Err("更新文件不存在，请重新下载".to_string());
    }

    let current = std::env::current_exe().map_err(|e| format!("无法定位当前程序: {e}"))?;
    let dir = current
        .parent()
        .ok_or_else(|| "无法确定程序所在目录".to_string())?
        .to_path_buf();
    let old = dir.join(format!(
        "{}.old",
        current
            .file_name()
            .map(|n| n.to_string_lossy().to_string())
            .unwrap_or_else(|| "MasterEdit.exe".to_string())
    ));

    // 清理上一次更新留下的备份
    let _ = std::fs::remove_file(&old);

    std::fs::rename(&current, &old).map_err(|e| format!("替换程序失败（改名阶段）: {e}"))?;
    if let Err(error) = std::fs::copy(&source, &current) {
        // 回滚，避免程序不可用
        let _ = std::fs::rename(&old, &current);
        return Err(format!("替换程序失败（写入阶段）: {error}"));
    }

    std::process::Command::new(&current)
        .spawn()
        .map_err(|e| format!("启动新版本失败: {e}"))?;
    Ok(())
}

/// 启动时清理上次更新遗留的 .old 备份
#[cfg(windows)]
pub fn cleanup_old_binary() {
    let Ok(current) = std::env::current_exe() else {
        return;
    };
    let name = current
        .file_name()
        .map(|n| n.to_string_lossy().to_string())
        .unwrap_or_else(|| "MasterEdit.exe".to_string());
    if let Some(dir) = current.parent() {
        let _ = std::fs::remove_file(dir.join(format!("{name}.old")));
    }
}

/// 退出应用（自更新替换完成后调用，确保旧进程立即结束）
/// 安装版静默升级：
/// 生成一个临时脚本，等待当前进程退出 → 静默运行安装包（/S）→ 启动新版 → 自删。
/// 由前端在下载完成后调用，随后立即退出应用。
///
/// 两个必须注意的点：
/// 1. 安装路径只能通过命令行参数传入脚本，**不能**拼进脚本文本：
///    路径里出现单引号（如 `C:\Users\O'Brien\...`）会让脚本语法错误甚至被注入；
/// 2. 静默安装会先卸载旧版本，可执行文件名也可能随版本变化（MasterMD.exe → MasterEdit.exe），
///    因此重启时按候选名逐个探测，找不到就放弃重启（由用户手动启动）。
#[cfg(windows)]
#[tauri::command]
pub async fn apply_installer_update(path: String) -> Result<(), String> {
    let setup = std::path::PathBuf::from(&path);
    if !setup.exists() {
        return Err("安装包不存在，请重新下载".to_string());
    }
    let current = std::env::current_exe().map_err(|e| format!("无法定位当前程序: {e}"))?;
    let pid = std::process::id();
    let script = std::env::temp_dir().join(format!("masteredit-update-{pid}.ps1"));
    let body = format!(
        "$ErrorActionPreference = 'SilentlyContinue'\r\n\
         $setup = $args[0]\r\n\
         $exe = $args[1]\r\n\
         while (Get-Process -Id {pid} -ErrorAction SilentlyContinue) {{ Start-Sleep -Milliseconds 700 }}\r\n\
         Start-Sleep -Milliseconds 400\r\n\
         Start-Process -FilePath $setup -ArgumentList '/S' -Wait\r\n\
         Start-Sleep -Milliseconds 600\r\n\
         $dir = Split-Path -Parent $exe\r\n\
         $candidates = @($exe, (Join-Path $dir 'MasterEdit.exe'), (Join-Path $dir 'masteredit.exe'), (Join-Path $dir 'MasterMD.exe'), (Join-Path $dir 'mastermd.exe'))\r\n\
         foreach ($candidate in $candidates) {{ if (Test-Path -LiteralPath $candidate) {{ Start-Process -FilePath $candidate; break }} }}\r\n\
         Remove-Item -LiteralPath $MyInvocation.MyCommand.Path -Force\r\n",
        pid = pid
    );
    std::fs::write(&script, body).map_err(|e| format!("无法写入升级脚本: {e}"))?;

    use std::os::windows::process::CommandExt;
    const CREATE_NO_WINDOW: u32 = 0x0800_0000;
    std::process::Command::new("powershell")
        .args([
            "-NoProfile",
            "-ExecutionPolicy",
            "Bypass",
            "-WindowStyle",
            "Hidden",
            "-File",
        ])
        .arg(&script)
        .arg(&setup)
        .arg(&current)
        .creation_flags(CREATE_NO_WINDOW)
        .spawn()
        .map_err(|e| format!("启动升级脚本失败: {e}"))?;
    Ok(())
}

#[cfg(windows)]
#[tauri::command]
pub fn quit_app(app: tauri::AppHandle) {
    app.exit(0);
}

/// 当前程序所在路径（关于面板展示，便于用户确认便携版位置）
#[cfg(windows)]
#[tauri::command]
pub fn current_exe_path() -> Option<String> {
    std::env::current_exe()
        .ok()
        .map(|p| p.to_string_lossy().to_string())
}
