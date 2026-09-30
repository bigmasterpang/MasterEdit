/**
 * 块数据层：窗口化取块 + 缓存 + 后台预取（`document_blocks` 的前端封装）。
 *
 * ── 取块策略 ──────────────────────────────────────────────────────────
 * 1. **打开即取第一窗**（120 块）：拿到 `total` 与 `encrypted`，并立刻能渲染首屏；
 * 2. **后台预取整篇**（上限 `MAX_PRELOAD` 块）：块模型很小（样本 184 段的文档也就几十 KB），
 *    预取完整篇的最大好处是「所有块的高度估算一次算好」，滚动时不会因为陆续加载而上下跳动，
 *    同时大纲与「复制全部文本」也不需要再等；
 * 3. 超过预取上限的巨型文档：虚拟滚动按窗口（对齐到 `WINDOW_SIZE`）按需取，
 *    取过的窗口进缓存，失败窗口记录在案**不重试**（否则会在渲染循环里打出无请求风暴）。
 *
 * 后端命令无状态、`from` 原样回显，所以「窗口对齐 + 只认当前文档代次（generation）」就足够丢弃过期响应。
 */
import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { invoke } from "@tauri-apps/api/core";
import type { DocBlock, DocBlockPage, DocPageSetup } from "../../types";

/** 一次取的块数（后端建议 50~200：首次是"解密+解包+建块"，之后是切片克隆） */
const WINDOW_SIZE = 120;
/** 打开后后台预取的块数上限（超过它的巨型文档转为按需窗口取） */
const MAX_PRELOAD = 3000;
/** 「复制全部文本」/ 大纲最多回头补取多少块（防极端文档把内存打满） */
const MAX_LOAD_ALL = 20000;

export interface DocxBlocksApi {
  /** 文档真实块总数（还没取到第一窗时为 0） */
  total: number;
  encrypted: boolean;
  loading: boolean;
  /** 取块失败的中文原因（窗口路失败时给出，界面显示提示条 + 重试按钮） */
  error: string | null;
  /** 已加载内容的版本号：缓存变化时 +1，供渲染层重建高度索引 */
  version: number;
  /** 取某一块（未加载返回 undefined） */
  blockAt(index: number): DocBlock | undefined;
  /** 确保 [from, to] 区间已加载（虚拟滚动窗口调用；重复调用安全） */
  ensure(from: number, to: number): void;
  /**
   * 取回整篇文档的块（复制全部文本 / 大纲用）。
   * 返回**按下标对齐的数组**（未取到的位置是 undefined）；`complete` 说明是否完整。
   */
  loadAll(): Promise<{ blocks: Array<DocBlock | undefined>; complete: boolean }>;
  /** 清掉失败记录并重新取当前需要的窗口 */
  retry(): void;
  /**
   * 页面设置（分页视图用，来自后端 `w:sectPr`）。
   * 后端字段还没落地时是 null，渲染层按 A4 + 2.54cm 页边距兜底。
   */
  page: DocPageSetup | null;
}

export function useDocxBlocks(path: string | null, modifiedAt: number): DocxBlocksApi {
  const [total, setTotal] = useState(0);
  const [encrypted, setEncrypted] = useState(false);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [version, setVersion] = useState(0);
  const [page, setPage] = useState<DocPageSetup | null>(null);

  /** 块缓存：下标 → 块 */
  const cacheRef = useRef(new Map<number, DocBlock>());
  /** 取失败的窗口（key = 窗口起始下标）：记录后不再自动重试，避免请求风暴 */
  const failedRef = useRef(new Set<number>());
  /** 正在飞的窗口：同一个窗口不会被重复请求 */
  const inflightRef = useRef(new Set<number>());
  /** 文档代次：换文档/外部改动后，旧响应全部丢弃 */
  const generationRef = useRef(0);
  const totalRef = useRef(0);
  const pathRef = useRef<string | null>(path);

  const bump = useCallback(() => setVersion((value) => value + 1), []);

  const windowStart = useCallback((index: number) => Math.floor(index / WINDOW_SIZE) * WINDOW_SIZE, []);

  /** 取一个窗口并写进缓存；返回是否成功 */
  const fetchWindow = useCallback(
    async (start: number, generation: number): Promise<boolean> => {
      const currentPath = pathRef.current;
      if (!currentPath) return false;
      if (inflightRef.current.has(start) || failedRef.current.has(start)) return false;
      inflightRef.current.add(start);
      try {
        const page = await invoke<DocBlockPage>("document_blocks", {
          path: currentPath,
          from: start,
          count: WINDOW_SIZE,
        });
        if (generationRef.current !== generation) return false;
        // 后端是无状态的：`from` 会被原样回显，用返回的 from 定位（不信本地 start）
        const from = Number.isFinite(page.from) ? page.from : start;
        page.blocks.forEach((block, offset) => cacheRef.current.set(from + offset, block));
        totalRef.current = page.total;
        setTotal(page.total);
        setEncrypted(page.encrypted);
        // 页面设置：后端还在补这个字段，缺了就当 null（渲染层按 A4 兜底）
        if (page.page) setPage(page.page);
        setError(null);
        bump();
        return true;
      } catch (reason) {
        if (generationRef.current !== generation) return false;
        failedRef.current.add(start);
        setError(typeof reason === "string" && reason ? reason : String(reason));
        return false;
      } finally {
        inflightRef.current.delete(start);
      }
    },
    [bump],
  );

  /* 打开文档（或外部改动）后重建缓存并取第一窗 + 后台预取 */
  useEffect(() => {
    pathRef.current = path;
    const generation = generationRef.current + 1;
    generationRef.current = generation;
    cacheRef.current = new Map();
    failedRef.current = new Set();
    inflightRef.current = new Set();
    totalRef.current = 0;
    setTotal(0);
    setEncrypted(false);
    setPage(null);
    setError(null);
    if (!path) {
      setLoading(false);
      bump();
      return;
    }
    setLoading(true);
    void (async () => {
      const ok = await fetchWindow(0, generation);
      if (generationRef.current !== generation) return;
      setLoading(false);
      if (!ok) return;
      // 后台把后续窗口一路取完（转发为上限）
      const limit = Math.min(totalRef.current, MAX_PRELOAD);
      for (let start = WINDOW_SIZE; start < limit; start += WINDOW_SIZE) {
        if (generationRef.current !== generation) return;
        await fetchWindow(start, generation);
      }
    })();
  }, [path, modifiedAt, fetchWindow, bump]);

  const ensure = useCallback(
    (from: number, to: number) => {
      const currentPath = pathRef.current;
      if (!currentPath || totalRef.current <= 0) return;
      const generation = generationRef.current;
      const start = windowStart(Math.max(0, from));
      const end = windowStart(Math.min(totalRef.current - 1, Math.max(from, to)));
      for (let index = start; index <= end; index += WINDOW_SIZE) {
        // 窗口内已全部命中缓存 → 跳过
        let missing = false;
        for (let i = index; i < index + WINDOW_SIZE && i < totalRef.current; i += 1) {
          if (!cacheRef.current.has(i)) {
            missing = true;
            break;
          }
        }
        if (!missing) continue;
        void fetchWindow(index, generation);
      }
    },
    [fetchWindow, windowStart],
  );

  /** 一次性补齐整篇（复制 / 大纲用），返回按下标对齐的数组 */
  const loadAll = useCallback(async (): Promise<{ blocks: Array<DocBlock | undefined>; complete: boolean }> => {
    const generation = generationRef.current;
    // 第一窗都还没到（用户点得很快）：先补上，否则 total 还是 0
    if (totalRef.current <= 0) await fetchWindow(0, generation);
    if (generationRef.current !== generation) return { blocks: [], complete: false };
    const limit = Math.min(totalRef.current, MAX_LOAD_ALL);
    // 先清掉失败记录：用户主动要整篇文本，值得重试一次
    failedRef.current.clear();
    for (let start = 0; start < limit; start += WINDOW_SIZE) {
      if (generationRef.current !== generation) break;
      if (inflightRef.current.has(start)) continue;
      let missing = false;
      for (let i = start; i < start + WINDOW_SIZE && i < limit; i += 1) {
        if (!cacheRef.current.has(i)) {
          missing = true;
          break;
        }
      }
      if (!missing) continue;
      await fetchWindow(start, generation);
    }
    const blocks: Array<DocBlock | undefined> = [];
    for (let index = 0; index < limit; index += 1) blocks.push(cacheRef.current.get(index));
    return { blocks, complete: limit >= totalRef.current };
  }, [fetchWindow]);

  const retry = useCallback(() => {
    // 只清失败记录：渲染层的窗口 effect 会立刻按当前可视范围重新 ensure
    failedRef.current.clear();
    setError(null);
    bump();
  }, [bump]);

  return useMemo(
    () => ({
      total,
      encrypted,
      loading,
      error,
      version,
      page,
      blockAt: (index: number) => cacheRef.current.get(index),
      ensure,
      loadAll,
      retry,
    }),
    [total, encrypted, loading, error, version, page, ensure, loadAll, retry],
  );
}
