import { useEffect, useRef, useState } from "react";
import * as pdfjsLib from "pdfjs-dist";
// @ts-ignore
import workerUrl from "pdfjs-dist/build/pdf.worker.min.mjs?url";
import { PDFDocument } from "pdf-lib";
import { Highlighter, MessageSquare, Pen, Download, Save } from "lucide-react";
import { saveFileDialog } from "../../platform";
import type { OpenedFile } from "../../platform";
import { InlinePrompt } from "../../components/InlinePrompt";

pdfjsLib.GlobalWorkerOptions.workerSrc = workerUrl;

type Ann =
  | { type: "highlight"; page: number; x: number; y: number; w: number; h: number; color: string }
  | { type: "comment"; page: number; x: number; y: number; text: string }
  | { type: "freehand"; page: number; points: { x: number; y: number }[]; color: string };

type Mode = "none" | "highlight" | "comment" | "freehand";

export function PdfViewer({ file }: { file: OpenedFile }) {
  const [numPages, setNumPages] = useState(0);
  const [mode, setMode] = useState<Mode>("none");
  const [anns, setAnns] = useState<Ann[]>([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string>("");
  const [saving, setSaving] = useState(false);
  const [commentPrompt, setCommentPrompt] = useState<null | { page: number; x: number; y: number; nx: number; ny: number }>(null);
  const containerRef = useRef<HTMLDivElement>(null);
  // mode 的镜像 ref：bindOverlay 绑定的事件闭包必须读到最新值，否则批注工具全部失效
  const modeRef = useRef<Mode>("none");
  modeRef.current = mode;
  const pagesRef = useRef<{ canvas: HTMLCanvasElement; viewport: any }[]>([]);
  const drawRef = useRef<{ page: number; startX: number; startY: number; points: { x: number; y: number }[] } | null>(null);

  useEffect(() => {
    let cancelled = false;
    (async () => {
      setLoading(true);
      setError("");
      try {
        const data = new Uint8Array(file.buffer.slice(0));
        const doc = await pdfjsLib.getDocument({ data }).promise;
        if (cancelled) return;
        setNumPages(doc.numPages);
        pagesRef.current = [];
        const container = containerRef.current;
        if (!container) return;
        // 注意：containerRef 指向的 div 在 JSX 中没有任何 React 子节点，
        // 此处清空/追加 canvas 不会与 React 协调器冲突（白屏根因已修复）
        container.innerHTML = "";
        const scale = 1.6;
        for (let i = 1; i <= doc.numPages; i++) {
          const page = await doc.getPage(i);
          const viewport = page.getViewport({ scale });
          const canvas = document.createElement("canvas");
          canvas.width = viewport.width;
          canvas.height = viewport.height;
          canvas.style.width = "100%";
          canvas.style.maxWidth = `${viewport.width}px`;
          canvas.style.margin = "0 auto 16px";
          canvas.style.display = "block";
          canvas.dataset.page = String(i - 1);
          const ctx = canvas.getContext("2d")!;
          await page.render({ canvasContext: ctx, viewport }).promise;

          const wrap = document.createElement("div");
          wrap.style.position = "relative";
          wrap.style.width = "100%";
          wrap.style.maxWidth = `${viewport.width}px`;
          wrap.style.margin = "0 auto 16px";
          const overlay = document.createElement("div");
          overlay.style.position = "absolute";
          overlay.style.inset = "0";
          overlay.style.cursor = modeRef.current === "none" ? "default" : "crosshair";
          wrap.appendChild(canvas);
          wrap.appendChild(overlay);
          container.appendChild(wrap);
          pagesRef.current.push({ canvas, viewport });
          bindOverlay(overlay, i - 1);
        }
      } catch (e) {
        if (!cancelled) setError("解析失败：" + (e as Error).message + "（加密或损坏的 PDF 暂不支持）");
      } finally {
        if (!cancelled) setLoading(false);
      }
    })();
    return () => {
      cancelled = true;
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [file]);

  function toNorm(el: HTMLElement, clientX: number, clientY: number) {
    const rect = el.getBoundingClientRect();
    return { x: (clientX - rect.left) / rect.width, y: (clientY - rect.top) / rect.height };
  }

  function bindOverlay(overlay: HTMLElement, page: number) {
    overlay.onmousedown = (e) => {
      if (modeRef.current === "none") return;
      const p = toNorm(overlay, e.clientX, e.clientY);
      if (modeRef.current === "highlight") {
        drawRef.current = { page, startX: p.x, startY: p.y, points: [] };
      } else if (modeRef.current === "comment") {
        setCommentPrompt({ page, x: e.clientX, y: e.clientY, nx: p.x, ny: p.y });
        setMode("none");
      } else if (modeRef.current === "freehand") {
        drawRef.current = { page, startX: p.x, startY: p.y, points: [p] };
      }
    };
    overlay.onmousemove = (e) => {
      if (!drawRef.current || drawRef.current.page !== page) return;
      const p = toNorm(overlay, e.clientX, e.clientY);
      if (modeRef.current === "highlight") {
        const s = drawRef.current;
        drawTemp(page, { type: "highlight", page, x: Math.min(s.startX, p.x), y: Math.min(s.startY, p.y), w: Math.abs(p.x - s.startX), h: Math.abs(p.y - s.startY), color: "#FFE066" });
      } else if (mode === "freehand") {
        drawRef.current.points.push(p);
        drawTemp(page, { type: "freehand", page, points: drawRef.current.points, color: "#FF3B30" });
      }
    };
    overlay.onmouseup = (e) => {
      if (!drawRef.current || drawRef.current.page !== page) return;
      const p = toNorm(overlay, e.clientX, e.clientY);
      if (modeRef.current === "highlight") {
        const s = drawRef.current;
        const w = Math.abs(p.x - s.startX), h = Math.abs(p.y - s.startY);
        if (w > 0.01 && h > 0.01) setAnns((a) => [...a, { type: "highlight", page, x: Math.min(s.startX, p.x), y: Math.min(s.startY, p.y), w, h, color: "#FFE066" }]);
      } else if (modeRef.current === "freehand") {
        setAnns((a) => [...a, { type: "freehand", page, points: drawRef.current!.points, color: "#FF3B30" }]);
      }
      drawRef.current = null;
      clearTemp(page);
      setMode("none");
    };
  }

  // 临时绘制（拖拽时预览）
  function drawTemp(page: number, ann: Ann) {
    const ctx = pagesRef.current[page]?.canvas.getContext("2d");
    if (!ctx) return;
    clearTemp(page);
    paintAnn(ctx, page, ann);
  }
  function clearTemp(page: number) {
    const item = pagesRef.current[page];
    if (!item) return;
    // 重绘整页（含已保存批注）
    const canvas = item.canvas;
    const saved = anns.filter((a) => a.page === page);
    const tmp = document.createElement("canvas");
    tmp.width = canvas.width; tmp.height = canvas.height;
    tmp.getContext("2d")!.drawImage(canvas, 0, 0);
    const ctx = tmp.getContext("2d")!;
    saved.forEach((a) => paintAnn(ctx, page, a));
    const cctx = canvas.getContext("2d")!;
    cctx.clearRect(0, 0, canvas.width, canvas.height);
    cctx.drawImage(tmp, 0, 0);
  }

  function paintAnn(ctx: CanvasRenderingContext2D, page: number, ann: Ann) {
    const vw = pagesRef.current[page]?.viewport.width || 1;
    const vh = pagesRef.current[page]?.viewport.height || 1;
    if (ann.type === "highlight") {
      ctx.save();
      ctx.globalAlpha = 0.4;
      ctx.fillStyle = ann.color;
      ctx.fillRect(ann.x * vw, ann.y * vh, ann.w * vw, ann.h * vh);
      ctx.restore();
    } else if (ann.type === "comment") {
      ctx.save();
      ctx.fillStyle = "#2E5BF0";
      ctx.beginPath();
      ctx.arc(ann.x * vw, ann.y * vh, 6, 0, Math.PI * 2);
      ctx.fill();
      ctx.fillStyle = "#fff";
      ctx.font = "10px sans-serif";
      ctx.fillText("!", ann.x * vw - 3, ann.y * vh + 3);
      ctx.fillStyle = "#1A2233";
      ctx.font = "13px sans-serif";
      ctx.fillText(ann.text, ann.x * vw + 10, ann.y * vh + 4);
      ctx.restore();
    } else if (ann.type === "freehand") {
      ctx.save();
      ctx.strokeStyle = ann.color;
      ctx.lineWidth = 2;
      ctx.beginPath();
      ann.points.forEach((pt, i) => {
        const X = pt.x * vw, Y = pt.y * vh;
        if (i === 0) ctx.moveTo(X, Y); else ctx.lineTo(X, Y);
      });
      ctx.stroke();
      ctx.restore();
    }
  }

  // 重绘所有页（批注变更时）
  useEffect(() => {
    if (loading) return;
    anns.forEach((a) => clearTemp(a.page));
    anns.forEach((a) => paintAnn(pagesRef.current[a.page]?.canvas.getContext("2d"), a.page, a));
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [anns, loading]);

  async function handleSave() {
    setSaving(true);
    try {
      const existing = await PDFDocument.load(new Uint8Array(file.buffer.slice(0)));
      // 把每页当前 canvas（含批注）作为图片，重绘到新 PDF 对应页尺寸
      for (let i = 0; i < pagesRef.current.length; i++) {
        const item = pagesRef.current[i];
        const canvas = item.canvas;
        const pngDataUrl = canvas.toDataURL("image/png");
        const pngBytes = await fetch(pngDataUrl).then((r) => r.arrayBuffer());
        const img = await existing.embedPng(pngBytes);
        const page = existing.getPage(i);
        const { width, height } = page.getSize();
        // 清空原内容并贴上带批注的图（批注已烘焙进图）
        page.drawImage(img, { x: 0, y: 0, width, height });
      }
      const bytes = await existing.save();
      await saveFileDialog(file.name.replace(/\.pdf$/i, "_annotated.pdf"), new Uint8Array(bytes));
    } finally {
      setSaving(false);
    }
  }

  return (
    <div style={{ height: "100%", display: "flex", flexDirection: "column" }}>
      <div style={{ display: "flex", gap: 8, alignItems: "center", padding: "8px 12px", background: "var(--card)", borderBottom: "1px solid var(--border)" }}>
        <button className={`icon-btn ${mode === "highlight" ? "active" : ""}`} title="高亮" onClick={() => setMode(mode === "highlight" ? "none" : "highlight")}>
          <Highlighter size={18} />
        </button>
        <button className={`icon-btn ${mode === "comment" ? "active" : ""}`} title="评论批注" onClick={() => setMode(mode === "comment" ? "none" : "comment")}>
          <MessageSquare size={18} />
        </button>
        <button className={`icon-btn ${mode === "freehand" ? "active" : ""}`} title="自由绘制" onClick={() => setMode(mode === "freehand" ? "none" : "freehand")}>
          <Pen size={18} />
        </button>
        <span className="separator" />
        <button className="btn primary" disabled={saving} onClick={handleSave}>
          <Download size={16} /> {saving ? "保存中…" : "保存批注 PDF"}
        </button>
        <div style={{ flex: 1 }} />
        <span style={{ fontSize: 12, color: "var(--muted-fg)" }}>{numPages} 页 · 批注 {anns.length}</span>
      </div>
      <div style={{ flex: 1, position: "relative", overflowY: "auto", background: "var(--muted)" }}>
        {/* canvas 容器：React 不渲染任何子节点，全部子内容由命令式代码管理（避免 removeChild 冲突） */}
        <div ref={containerRef} style={{ padding: 16 }} />
        {loading && (
          <div style={{ position: "absolute", top: 24, left: 0, right: 0, textAlign: "center", color: "var(--muted-fg)" }}>
            正在解析 PDF…
          </div>
        )}
        {!loading && !!error && (
          <div style={{ position: "absolute", top: 24, left: 0, right: 0, padding: "0 24px", textAlign: "center", color: "var(--warning)", fontSize: 13 }}>
            {error}
          </div>
        )}
        {commentPrompt && (
          <InlinePrompt
            x={commentPrompt.x}
            y={commentPrompt.y}
            placeholder="批注内容（Enter 确定）"
            onOk={(text) => {
              setAnns((a) => [...a, { type: "comment", page: commentPrompt.page, x: commentPrompt.nx, y: commentPrompt.ny, text }]);
              setCommentPrompt(null);
            }}
            onCancel={() => setCommentPrompt(null)}
          />
        )}
      </div>
    </div>
  );
}
