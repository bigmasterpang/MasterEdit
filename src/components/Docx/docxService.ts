import { useEffect, useState } from "react";
import type { OutlineItem } from "./docxGrid";

export interface DocxOutlineState {
  items: OutlineItem[];
  loading: boolean;
  partial: boolean;
}

const docxOutlineMap = new Map<string, DocxOutlineState>();
const outlineListeners = new Set<() => void>();

export function setDocxOutline(docId: string, state: DocxOutlineState): void {
  docxOutlineMap.set(docId, state);
  outlineListeners.forEach((l) => l());
}

export function clearDocxOutline(docId: string): void {
  docxOutlineMap.delete(docId);
  outlineListeners.forEach((l) => l());
}

export function useDocxOutline(docId: string | null | undefined): DocxOutlineState {
  const [, setTick] = useState(0);
  useEffect(() => {
    const handler = () => setTick((t) => t + 1);
    outlineListeners.add(handler);
    return () => {
      outlineListeners.delete(handler);
    };
  }, []);
  return docId
    ? docxOutlineMap.get(docId) ?? { items: [], loading: false, partial: false }
    : { items: [], loading: false, partial: false };
}

export function jumpToDocxBlock(docId: string, blockIndex: number): void {
  window.dispatchEvent(
    new CustomEvent("docx-jump-to-block", {
      detail: { docId, blockIndex },
    }),
  );
}
