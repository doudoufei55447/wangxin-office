import { useEffect, useRef, useState } from "react";
import JSZip from "jszip";
import { Save, MessageSquare, ChevronLeft, ChevronRight, AlertTriangle, Image as ImageIcon } from "lucide-react";
import { saveFileDialog } from "../../platform";
import type { OpenedFile } from "../../platform";

interface Shape {
  x: number; y: number; cx: number; cy: number;
  isPic: boolean;
  textHtml: string;
  color: string;
  fontSize: number; // sz (hundredths of point)
  bold: boolean;
  align: "left" | "center" | "right";
  img: string;
}
interface Slide { W: number; H: number; bgColor: string; shapes: Shape[]; runs: string[]; xml: string; }

const A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";
const P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main";

function getByLocal(root: Document | Element, name: string): Element[] {
  const all = root.getElementsByTagName("*");
  const res: Element[] = [];
  for (let i = 0; i < all.length; i++) if (all[i].localName === name) res.push(all[i]);
  return res;
}

function resolveMedia(slideName: string, target: string): string {
  // slideName: ppt/slides/slide1.xml ; target: ../media/image1.png
  if (target.startsWith("http")) return target;
  if (target.startsWith("../")) return "ppt/" + target.slice(3);
  if (target.startsWith("media/")) return "ppt/" + target;
  return "ppt/slides/" + target;
}

async function parseSlide(zip: JSZip, name: string): Promise<Slide> {
  const xml = await zip.file(name)!.async("string");
  const relName = name.replace("slides/", "slides/_rels/") + ".rels";
  const relsXml = (await zip.file(relName)?.async("string")) || "<Relationships/>";
  const relDoc = new DOMParser().parseFromString(relsXml, "application/xml");
  const relMap: Record<string, string> = {};
  relDoc.getElementsByTagName("Relationship").forEach((r) => {
    relMap[r.getAttribute("Id") || ""] = r.getAttribute("Target") || "";
  });

  const doc = new DOMParser().parseFromString(xml, "application/xml");
  const sldSz = getByLocal(doc, "sldSz")[0];
  const W = sldSz ? +sldSz.getAttribute("cx")! : 12192000;
  const H = sldSz ? +sldSz.getAttribute("cy")! : 6858000;

  let bgColor = "#FFFFFF";
  const bg = getByLocal(doc, "bg")[0];
  if (bg) {
    const sf = getByLocal(bg, "srgbClr")[0];
    if (sf) bgColor = "#" + sf.getAttribute("val")!;
  }

  const shapes: Shape[] = [];
  const sps = [...getByLocal(doc, "sp"), ...getByLocal(doc, "pic")];
  for (const sp of sps) {
    const spPr = getByLocal(sp, "spPr")[0];
    const xfrm = spPr ? getByLocal(spPr, "xfrm")[0] : null;
    if (!xfrm) continue;
    const off = getByLocal(xfrm, "off")[0];
    const ext = getByLocal(xfrm, "ext")[0];
    if (!off || !ext) continue;
    const geo = {
      x: +off.getAttribute("x")!,
      y: +off.getAttribute("y")!,
      cx: +ext.getAttribute("cx")!,
      cy: +ext.getAttribute("cy")!,
    };
    const isPic = sp.localName === "pic";
    let textHtml = "";
    let color = "#1A2233";
    let fontSize = 1800;
    let bold = false;
    let align: "left" | "center" | "right" = "left";

    const txBody = getByLocal(sp, "txBody")[0];
    if (txBody) {
      const ps = getByLocal(txBody, "p");
      const lines: string[] = [];
      for (const p of ps) {
        const rs = getByLocal(p, "r");
        let line = "";
        for (const r of rs) {
          const t = getByLocal(r, "t")[0];
          line += t ? t.textContent || "" : "";
        }
        if (rs.length) {
          const rPr = getByLocal(rs[0], "rPr")[0];
          if (rPr) {
            const sz = rPr.getAttribute("sz");
            if (sz) fontSize = +sz;
            if (rPr.getAttribute("b") === "1") bold = true;
            const sf = getByLocal(rPr, "srgbClr")[0];
            if (sf) color = "#" + sf.getAttribute("val")!;
          }
          const pPr = getByLocal(p, "pPr")[0];
          if (pPr) {
            const a = pPr.getAttribute("algn");
            if (a === "ctr") align = "center";
            else if (a === "r") align = "right";
          }
        }
        lines.push(line);
      }
      textHtml = lines
        .join("\n")
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/\n/g, "<br/>");
    }

    let img = "";
    if (isPic) {
      const blip = getByLocal(sp, "blip")[0];
      const rid = blip?.getAttribute("r:embed");
      if (rid && relMap[rid]) {
        const mediaPath = resolveMedia(name, relMap[rid]);
        const data = await zip.file(mediaPath)?.async("base64");
        const ext = mediaPath.split(".").pop() || "png";
        if (data) img = `data:image/${ext};base64,${data}`;
      }
    }
    shapes.push({ ...geo, isPic, textHtml, color, fontSize, bold, align, img });
  }

  const runs: string[] = [];
  for (const sp of sps) {
    const txBody = getByLocal(sp, "txBody")[0];
    if (txBody) for (const t of getByLocal(txBody, "t")) runs.push(t.textContent || "");
  }
  return { W, H, bgColor, shapes, runs, xml };
}

export function PptxViewer({ file }: { file: OpenedFile }) {
  const [slides, setSlides] = useState<Slide[]>([]);
  const [page, setPage] = useState(0);
  const [notes, setNotes] = useState<{ page: number; x: number; y: number; text: string }[]>([]);
  const [saving, setSaving] = useState(false);
  const [loading, setLoading] = useState(true);
  const zipRef = useRef<JSZip | null>(null);
  const stageRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    (async () => {
      setLoading(true);
      try {
        const zip = await JSZip.loadAsync(new Uint8Array(file.buffer.slice(0)));
        zipRef.current = zip;
        const names = Object.keys(zip.files)
          .filter((n) => /^ppt\/slides\/slide\d+\.xml$/.test(n))
          .sort((a, b) => +a.match(/slide(\d+)/)![1] - +b.match(/slide(\d+)/)![1]);
        const parsed: Slide[] = [];
        for (const n of names) parsed.push(await parseSlide(zip, n));
        setSlides(parsed);
      } catch (e) {
        setSlides([{ W: 12192000, H: 6858000, bgColor: "#fff", shapes: [], runs: ["解析失败：" + (e as Error).message], xml: "" }]);
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

  // 受限文本替换：把编辑后的 runs 按文档顺序写回对应 slide 的 <a:t>
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
        xml = xml.replace(/<a:t>([\s\S]*?)<\/a:t>/g, () => {
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
          <AlertTriangle size={14} /> 查看+改字保排版（纯前端，不改布局）
        </span>
        <div style={{ flex: 1 }} />
        <button className="icon-btn" onClick={() => setPage((p) => Math.max(0, p - 1))}><ChevronLeft size={18} /></button>
        <span style={{ fontSize: 12, color: "var(--muted-fg)" }}>{page + 1} / {slides.length}</span>
        <button className="icon-btn" onClick={() => setPage((p) => Math.min(slides.length - 1, p + 1))}><ChevronRight size={18} /></button>
      </div>
      <div style={{ flex: 1, overflowY: "auto", padding: 24, background: "var(--muted)", display: "flex", justifyContent: "center" }}>
        <div
          ref={stageRef}
          onClick={addNote}
          style={{ width: "100%", maxWidth: 960, position: "relative", cursor: "crosshair", boxShadow: "var(--shadow)", background: slide?.bgColor }}
        >
          {loading && <div style={{ padding: 24, color: "var(--muted-fg)" }}>正在解析演示文稿…</div>}
          {slide && (
            <svg viewBox={`0 0 ${slide.W} ${slide.H}`} style={{ width: "100%", display: "block", background: slide.bgColor }}>
              {slide.shapes.map((s, i) =>
                s.isPic ? (
                  <image key={i} x={s.x} y={s.y} width={s.cx} height={s.cy} href={s.img} preserveAspectRatio="none" />
                ) : (
                  <foreignObject key={i} x={s.x} y={s.y} width={s.cx} height={s.cy}>
                    <div
                      xmlns="http://www.w3.org/1999/xhtml"
                      style={{
                        width: "100%", height: "100%", boxSizing: "border-box",
                        display: "flex", flexDirection: "column", justifyContent: "center",
                        padding: "2%", overflow: "hidden",
                        color: s.color, fontWeight: s.bold ? 700 : 400,
                        fontSize: (s.fontSize / 100) * 12700, lineHeight: 1.2,
                        textAlign: s.align, fontFamily: "PingFang SC, Microsoft YaHei, sans-serif",
                        whiteSpace: "pre-wrap", wordBreak: "break-word",
                      }}
                      dangerouslySetInnerHTML={{ __html: s.textHtml }}
                    />
                  </foreignObject>
                )
              )}
            </svg>
          )}
          {notes.filter((n) => n.page === page).map((n, i) => (
            <div
              key={i}
              style={{ position: "absolute", left: `${n.x * 100}%`, top: `${n.y * 100}%`, background: "#2E5BF0", color: "#fff", borderRadius: 12, padding: "2px 8px", fontSize: 11, maxWidth: 180, transform: "translate(-50%,-50%)", pointerEvents: "none" }}
            >
              {n.text}
            </div>
          ))}
        </div>
      </div>
      {/* 文本替换面板 */}
      <div style={{ width: 320, flexShrink: 0, background: "var(--card)", borderLeft: "1px solid var(--border)", padding: 12, overflowY: "auto" }}>
        <div style={{ fontWeight: 600, fontSize: 13, marginBottom: 8 }}>文本替换（第 {page + 1} 页 · 共 {slide?.runs.length ?? 0} 段）</div>
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
        {(!slide || slide.runs.length === 0) && <div style={{ fontSize: 12, color: "var(--muted-fg)", display: "flex", gap: 4 }}><ImageIcon size={14} /> 本页无文本（仅图片）</div>}
      </div>
    </div>
  );
}
