import { useEffect, useRef, useState } from "react";
import JSZip from "jszip";
import { Save, MessageSquare, ChevronLeft, ChevronRight, AlertTriangle } from "lucide-react";
import { saveFileDialog } from "../../platform";
import type { OpenedFile } from "../../platform";

interface SlideText { runs: string[]; xml: string; }

export function PptxViewer({ file }: { file: OpenedFile }) {
  const [slides, setSlides] = useState<SlideText[]>([]);
  const [page, setPage] = useState(0);
  const [notes, setNotes] = useState<{ page: number; x: number; y: number; text: string }[]>([]);
  const [saving, setSaving] = useState(false);
  const [loading, setLoading] = useState(true);
  const zipRef = useRef<JSZip | null>(null);

  useEffect(() => {
    (async () => {
      setLoading(true);
      try {
        const zip = await JSZip.loadAsync(new Uint8Array(file.buffer.slice(0)));
        zipRef.current = zip;
        const names = Object.keys(zip.files)
          .filter((n) => /^ppt\/slides\/slide\d+\.xml$/.test(n))
          .sort((a, b) => {
            const na = +a.match(/slide(\d+)/)![1];
            const nb = +b.match(/slide(\d+)/)![1];
            return na - nb;
          });
        const parsed: SlideText[] = [];
        for (const n of names) {
          const xml = await zip.file(n)!.async("string");
          const runs = Array.from(xml.matchAll(/<a:t>([\s\S]*?)<\/a:t>/g)).map((m) => m[1]);
          parsed.push({ runs, xml });
        }
        setSlides(parsed);
      } catch (e) {
        setSlides([{ runs: ["解析失败：" + (e as Error).message], xml: "" }]);
      } finally {
        setLoading(false);
      }
    })();
  }, [file]);

  function addNote(e: React.MouseEvent) {
    const text = window.prompt("批注内容（PPT 批注为会话内显示，不写回文件）：");
    if (!text) return;
    const rect = (e.currentTarget as HTMLElement).getBoundingClientRect();
    setNotes((n) => [...n, { page, x: (e.clientX - rect.left) / rect.width, y: (e.clientY - rect.top) / rect.height, text }]);
  }

  // 受限文本替换：把编辑后的 runs 写回对应 slide XML（仅替换 <a:t> 文本，不改排版）
  async function handleSave() {
    if (!zipRef.current) return;
    setSaving(true);
    try {
      for (let i = 0; i < slides.length; i++) {
        const slideName = `ppt/slides/slide${i + 1}.xml`;
        const f = zipRef.current.file(slideName);
        if (!f) continue;
        let xml = await f.async("string");
        let idx = 0;
        xml = xml.replace(/<a:t>([\s\S]*?)<\/a:t>/g, (_m, _g1) => {
          const rep = slides[i].runs[idx] ?? "";
          idx++;
          return `<a:t>${escapeXml(rep)}</a:t>`;
        });
        zipRef.current.file(slideName, xml);
      }
      const out = await zipRef.current.generateAsync({ type: "uint8array" });
      await saveFileDialog(file.name, out);
    } finally {
      setSaving(false);
    }
  }

  function escapeXml(s: string) {
    return s.replace(/[<>&'"]/g, (c) => ({ "<": "&lt;", ">": "&gt;", "&": "&amp;", "'": "&apos;", '"': "&quot;" }[c]!));
  }

  const slide = slides[page];

  return (
    <div style={{ height: "100%", display: "flex", flexDirection: "column" }}>
      <div style={{ display: "flex", gap: 8, alignItems: "center", padding: "8px 12px", background: "var(--card)", borderBottom: "1px solid var(--border)" }}>
        <button className="btn primary" disabled={saving} onClick={handleSave}>
          <Save size={16} /> {saving ? "保存中…" : "保存文本修改"}
        </button>
        <button className="icon-btn" title="批注" onClick={addNote as any}><MessageSquare size={18} /></button>
        <span style={{ display: "flex", gap: 4, alignItems: "center", fontSize: 12, color: "var(--warning)" }}>
          <AlertTriangle size={14} /> 降级能力：查看+批注+文本替换，排版不保真
        </span>
        <div style={{ flex: 1 }} />
        <button className="icon-btn" onClick={() => setPage((p) => Math.max(0, p - 1))}><ChevronLeft size={18} /></button>
        <span style={{ fontSize: 12, color: "var(--muted-fg)" }}>{page + 1} / {slides.length}</span>
        <button className="icon-btn" onClick={() => setPage((p) => Math.min(slides.length - 1, p + 1))}><ChevronRight size={18} /></button>
      </div>
      <div style={{ flex: 1, overflowY: "auto", padding: 24, background: "var(--muted)", display: "flex", gap: 24 }}>
        {/* 幻灯片内容视图 */}
        <div style={{ flex: 1, display: "flex", justifyContent: "center" }}>
          <div
            onClick={addNote}
            style={{
              width: "100%", maxWidth: 720, aspectRatio: "16 / 9", background: "#fff", color: "#1A2233",
              borderRadius: 8, boxShadow: "var(--shadow)", padding: 32, overflow: "auto", position: "relative", cursor: "crosshair",
            }}
          >
            {slide?.runs.map((t, i) => (
              <div key={i} style={{ marginBottom: 10, fontSize: i === 0 ? 22 : 15, fontWeight: i === 0 ? 700 : 400 }}>{t || " "}</div>
            ))}
            {notes.filter((n) => n.page === page).map((n, i) => (
              <div key={i} style={{ position: "absolute", left: `${n.x * 100}%`, top: `${n.y * 100}%`, background: "#2E5BF0", color: "#fff", borderRadius: 12, padding: "2px 8px", fontSize: 11, maxWidth: 160, transform: "translate(-50%,-50%)" }}>
                {n.text}
              </div>
            ))}
          </div>
        </div>
        {/* 文本替换面板 */}
        <div style={{ width: 300, flexShrink: 0, background: "var(--card)", borderRadius: 8, padding: 12, overflowY: "auto" }}>
          <div style={{ fontWeight: 600, fontSize: 13, marginBottom: 8 }}>文本替换（第 {page + 1} 页）</div>
          {slide?.runs.map((t, i) => (
            <textarea
              key={i}
              value={t}
              onChange={(e) => {
                const v = e.target.value;
                setSlides((prev) => {
                  const next = prev.slice();
                  next[page] = { ...next[page], runs: next[page].runs.map((r, j) => (j === i ? v : r)) };
                  return next;
                });
              }}
              rows={2}
              style={{ width: "100%", marginBottom: 8, background: "var(--muted)", color: "var(--fg)", border: "1px solid var(--border)", borderRadius: 6, padding: 6, font: "inherit", resize: "vertical" }}
            />
          ))}
        </div>
      </div>
    </div>
  );
}
