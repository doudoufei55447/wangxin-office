import { useEffect, useRef, useState } from "react";
import JSZip from "jszip";
import { Save, MessageSquare, ChevronLeft, ChevronRight, AlertTriangle, Image as ImageIcon } from "lucide-react";
import { saveFileDialog, isLegacyBinaryFormat } from "../../platform";
import { InlinePrompt } from "../../components/InlinePrompt";
import type { OpenedFile } from "../../platform";

interface TableData { colW: number[]; rows: string[][]; }
interface Shape {
  x: number; y: number; cx: number; cy: number;
  isPic: boolean;
  lines: string[]; // 文本行（原生 SVG text/tspan 渲染，不用 foreignObject——Chromium 154 起存在不绘制回归）
  color: string;
  fontSize: number; // sz (hundredths of point)
  bold: boolean;
  align: "left" | "center" | "right";
  img: string;
  table?: TableData;
  placeholder?: string;
}
interface Slide { W: number; H: number; bgColor: string; shapes: Shape[]; runs: string[]; xml: string; }

interface Geo { x: number; y: number; cx: number; cy: number; }
// 子坐标系 → 幻灯片坐标系映射：slide = d + v * s
interface SpaceMap { dx: number; dy: number; sx: number; sy: number; }
const IDENTITY: SpaceMap = { dx: 0, dy: 0, sx: 1, sy: 1 };

interface ParseCtx {
  zip: JSZip;
  phCache: Map<string, Map<string, Geo>>;
  W: number; H: number;
}

const A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";
const P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main";
const REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships";
const REL_BASE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

function getByLocal(root: Document | Element, name: string): Element[] {
  const all = root.getElementsByTagName("*");
  const res: Element[] = [];
  for (let i = 0; i < all.length; i++) if (all[i].localName === name) res.push(all[i]);
  return res;
}

function num(el: Element | null | undefined, attr: string, fallback = 0): number {
  if (!el) return fallback;
  const v = el.getAttribute(attr);
  return v === null ? fallback : +v || fallback;
}

// 解析 part 的 .rels：byId（r:embed 引用）与 byType（slideLayout/slideMaster 引用），Target 归一化为 zip 路径
function resolveRelPath(partDir: string, target: string): string {
  if (/^https?:/.test(target) || target.startsWith("/")) {
    return target.startsWith("/") ? target.slice(1) : target;
  }
  const parts = (partDir ? partDir + "/" : "") + target;
  const out: string[] = [];
  for (const seg of parts.split("/")) {
    if (seg === "..") out.pop();
    else if (seg !== "." && seg !== "") out.push(seg);
  }
  return out.join("/");
}

async function partRels(zip: JSZip, partPath: string): Promise<{ byId: Record<string, string>; byType: Record<string, string> }> {
  const i = partPath.lastIndexOf("/");
  const dir = i >= 0 ? partPath.slice(0, i) : "";
  const base = i >= 0 ? partPath.slice(i + 1) : partPath;
  const relPath = `${dir}/_rels/${base}.rels`;
  const relsXml = (await zip.file(relPath)?.async("string")) || "<Relationships/>";
  const doc = new DOMParser().parseFromString(relsXml, "application/xml");
  const byId: Record<string, string> = {};
  const byType: Record<string, string> = {};
  // XML 文档上 getElementsByTagName 返回 HTMLCollection（无 forEach），必须先转数组
  Array.from(doc.getElementsByTagName("Relationship")).forEach((r) => {
    const id = r.getAttribute("Id") || "";
    const type = r.getAttribute("Type") || "";
    const target = resolveRelPath(dir, r.getAttribute("Target") || "");
    if (id) byId[id] = target;
    const kind = type.slice(type.lastIndexOf("/") + 1);
    if (kind && !byType[kind]) byType[kind] = target;
  });
  return { byId, byType };
}

function phKey(ph: Element): string {
  return (ph.getAttribute("type") || "body") + "#" + (ph.getAttribute("idx") || "");
}

function lookupPh(map: Map<string, Geo> | undefined, key: string): Geo | null {
  if (!map) return null;
  const [t, idx] = key.split("#");
  return map.get(key) || map.get(t + "#") || (idx ? map.get("body#" + idx) : null) || null;
}

// 收集 layout/master 中占位符的几何；layout 缺失的再向 master 继承
async function phMapFor(ctx: ParseCtx, partPath: string, isLayout: boolean): Promise<Map<string, Geo>> {
  const cached = ctx.phCache.get(partPath);
  if (cached) return cached;
  const map = new Map<string, Geo>();
  ctx.phCache.set(partPath, map); // 先占位防循环引用
  const xml = await ctx.zip.file(partPath)?.async("string");
  if (!xml) return map;
  const doc = new DOMParser().parseFromString(xml, "application/xml");
  const entries: { key: string; geo: Geo | null }[] = [];
  for (const sp of getByLocal(doc, "sp")) {
    const ph = getByLocal(sp, "ph")[0];
    if (!ph) continue;
    const xfrm = getByLocal(sp, "xfrm")[0];
    let geo: Geo | null = null;
    if (xfrm) {
      const off = getByLocal(xfrm, "off")[0];
      const ext = getByLocal(xfrm, "ext")[0];
      if (off && ext) {
        geo = { x: num(off, "x"), y: num(off, "y"), cx: num(ext, "cx"), cy: num(ext, "cy") };
      }
    }
    entries.push({ key: phKey(ph), geo });
  }
  if (isLayout) {
    const rels = await partRels(ctx.zip, partPath);
    const masterPath = rels.byType["slideMaster"];
    const masterMap = masterPath ? await phMapFor(ctx, masterPath, false) : null;
    for (const it of entries) {
      if (it.geo) map.set(it.key, it.geo);
      else {
        const inh = lookupPh(masterMap, it.key);
        if (inh) map.set(it.key, inh);
      }
    }
  } else {
    for (const it of entries) if (it.geo) map.set(it.key, it.geo);
  }
  return map;
}

// 遍历 spTree：sp/pic 渲染、grpSp 应用组合变换递归、graphicFrame 渲染表格/占位提示。
// 关键点：runs 必须按 XML 文档顺序收集（写回时按序替换 <a:t>，跨形状/表格不串位）。
async function walkContainer(
  ctx: ParseCtx,
  container: Element,
  g: SpaceMap,
  relById: Record<string, string>,
  layoutMap: Map<string, Geo>,
  shapes: Shape[],
  runs: string[]
) {
  for (const child of Array.from(container.children)) {
    const tag = child.localName;
    if (tag === "sp" || tag === "pic") {
      await walkShape(ctx, child, tag === "pic", g, relById, layoutMap, shapes, runs);
    } else if (tag === "grpSp") {
      const grpSpPr = getByLocal(child, "grpSpPr")[0];
      const xfrm = grpSpPr ? getByLocal(grpSpPr, "xfrm")[0] : null;
      let cg = g;
      if (xfrm) {
        const off = getByLocal(xfrm, "off")[0];
        const ext = getByLocal(xfrm, "ext")[0];
        const chOff = getByLocal(xfrm, "chOff")[0];
        const chExt = getByLocal(xfrm, "chExt")[0];
        const chCx = num(chOff ? chExt : null, "cx");
        const chCy = num(chOff ? chExt : null, "cy");
        if (off && ext && chOff && chExt && chCx !== 0 && chCy !== 0) {
          const sxg = num(ext, "cx") / chCx;
          const syg = num(ext, "cy") / chCy;
          // slide = g(off + (x - chOff) * ext/chExt)
          cg = {
            dx: g.dx + (num(off, "x") - num(chOff, "x") * sxg) * g.sx,
            dy: g.dy + (num(off, "y") - num(chOff, "y") * syg) * g.sy,
            sx: sxg * g.sx,
            sy: syg * g.sy,
          };
        }
      }
      await walkContainer(ctx, child, cg, relById, layoutMap, shapes, runs);
    } else if (tag === "graphicFrame") {
      await walkFrame(ctx, child, g, layoutMap, shapes, runs);
    }
  }
}

function xfrmToGeo(xfrm: Element, g: SpaceMap): Geo | null {
  const off = getByLocal(xfrm, "off")[0];
  const ext = getByLocal(xfrm, "ext")[0];
  if (!off || !ext) return null;
  return {
    x: g.dx + num(off, "x") * g.sx,
    y: g.dy + num(off, "y") * g.sy,
    cx: num(ext, "cx") * g.sx,
    cy: num(ext, "cy") * g.sy,
  };
}

async function walkShape(
  ctx: ParseCtx,
  sp: Element,
  isPic: boolean,
  g: SpaceMap,
  relById: Record<string, string>,
  layoutMap: Map<string, Geo>,
  shapes: Shape[],
  runs: string[]
) {
  const spPr = getByLocal(sp, "spPr")[0];
  const xfrm = spPr ? getByLocal(spPr, "xfrm")[0] : null;
  let geo: Geo | null = null;
  if (xfrm) {
    geo = xfrmToGeo(xfrm, g);
  } else {
    // 占位符无自身坐标 → 从 layout/master 继承（占位符几何在幻灯片坐标系，不再套组合映射）
    const ph = getByLocal(sp, "ph")[0];
    if (ph) geo = lookupPh(layoutMap, phKey(ph));
  }

  let lines: string[] = [];
  let color = "#1A2233";
  let fontSize = 1800;
  let bold = false;
  let align: "left" | "center" | "right" = "left";

  const txBody = getByLocal(sp, "txBody")[0];
  if (txBody) {
    const ps = getByLocal(txBody, "p");
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
    // runs 按文档顺序逐个 a:t 收集（与保存写回一一对应）
    for (const t of getByLocal(txBody, "t")) runs.push(t.textContent || "");
  }

  let img = "";
  if (isPic) {
    const blip = getByLocal(sp, "blip")[0];
    const rid = blip?.getAttribute("r:embed");
    if (rid && relById[rid]) {
      const mediaPath = relById[rid];
      const data = await ctx.zip.file(mediaPath)?.async("base64");
      const ext = mediaPath.split(".").pop() || "png";
      if (data) img = `data:image/${ext};base64,${data}`;
    }
  }

  if (!geo) return; // 无法定位的形状不渲染，但 runs 已收录（面板可改、保存可写回）
  shapes.push({ ...geo, isPic, lines, color, fontSize, bold, align, img });
}

async function walkFrame(
  ctx: ParseCtx,
  frame: Element,
  g: SpaceMap,
  layoutMap: Map<string, Geo>,
  shapes: Shape[],
  runs: string[]
) {
  const xfrm = getByLocal(frame, "xfrm")[0]; // graphicFrame 用 p:xfrm
  let geo: Geo | null = null;
  if (xfrm) {
    geo = xfrmToGeo(xfrm, g);
  } else {
    const ph = getByLocal(frame, "ph")[0];
    if (ph) geo = lookupPh(layoutMap, phKey(ph));
  }
  if (!geo) return;

  const tbl = getByLocal(frame, "tbl")[0];
  if (tbl) {
    const colW = getByLocal(tbl, "gridCol").map((c) => num(c, "w"));
    const rows: string[][] = [];
    for (const tr of getByLocal(tbl, "tr")) {
      const row: string[] = [];
      for (const tc of getByLocal(tr, "tc")) {
        const txBody = getByLocal(tc, "txBody")[0];
        const parts: string[] = [];
        if (txBody) {
          for (const p of getByLocal(txBody, "p")) {
            parts.push(
              getByLocal(p, "t")
                .map((t) => t.textContent || "")
                .join("")
            );
          }
        }
        for (const t of getByLocal(tc, "t")) runs.push(t.textContent || "");
        row.push(parts.join("\n"));
      }
      rows.push(row);
    }
    shapes.push({ ...geo, isPic: false, lines: [], color: "#1A2233", fontSize: 1400, bold: false, align: "left", img: "", table: { colW, rows } });
  } else {
    // 图表 / SmartArt 等复杂对象：位置正确、内容给出占位提示而非静默丢失
    shapes.push({ ...geo, isPic: false, lines: [], color: "#8a6d1a", fontSize: 1400, bold: false, align: "center", img: "", placeholder: "图表 / SmartArt 等复杂对象暂不支持预览" });
  }
}

async function parseSlide(ctx: ParseCtx, name: string): Promise<Slide> {
  const xml = await ctx.zip.file(name)!.async("string");
  const rels = await partRels(ctx.zip, name);

  // 版面尺寸来自 presentation.xml 的 sldSz（ctx 中已读取）；背景色仍在 slide 内
  const doc = new DOMParser().parseFromString(xml, "application/xml");

  // 占位符几何：layout → master 两级继承（按 slide rels 找到 layout）
  let layoutMap = new Map<string, Geo>();
  if (rels.byType["slideLayout"]) {
    layoutMap = await phMapFor(ctx, rels.byType["slideLayout"], true);
  }

  const shapes: Shape[] = [];
  const runs: string[] = [];
  const cSld = getByLocal(doc, "cSld")[0];
  const spTree = cSld ? getByLocal(cSld, "spTree")[0] : null;
  if (spTree) await walkContainer(ctx, spTree, IDENTITY, rels.byId, layoutMap, shapes, runs);

  let bgColor = "#FFFFFF";
  const bg = getByLocal(doc, "bg")[0];
  if (bg) {
    const sf = getByLocal(bg, "srgbClr")[0];
    if (sf) bgColor = "#" + sf.getAttribute("val")!;
  }
  return { W: ctx.W, H: ctx.H, bgColor, shapes, runs, xml };
}

export function PptxViewer({ file }: { file: OpenedFile }) {
  const [slides, setSlides] = useState<Slide[]>([]);
  const [page, setPage] = useState(0);
  const [notes, setNotes] = useState<{ page: number; x: number; y: number; text: string }[]>([]);
  const [saving, setSaving] = useState(false);
  const [loading, setLoading] = useState(true);
  const [notePrompt, setNotePrompt] = useState<null | { x: number; y: number; nx: number; ny: number }>(null);
  const zipRef = useRef<JSZip | null>(null);
  const stageRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    (async () => {
      setLoading(true);
      try {
        if (isLegacyBinaryFormat(file.buffer)) {
          throw new Error(
            "这是旧版 PowerPoint 97-2003 二进制格式（.ppt）或 WPS 以 pptx 后缀保存的旧格式，本版本暂不支持。请用 PowerPoint/WPS 打开后「另存为 .pptx」再打开。"
          );
        }
        const zip = await JSZip.loadAsync(new Uint8Array(file.buffer.slice(0)));
        zipRef.current = zip;

        // 版面尺寸从 presentation.xml 的 sldSz 读取（slide.xml 中没有该节点）
        let W = 12192000, H = 6858000;
        const presXml = await zip.file("ppt/presentation.xml")?.async("string");
        if (presXml) {
          const pd = new DOMParser().parseFromString(presXml, "application/xml");
          const s = getByLocal(pd, "sldSz")[0];
          if (s) {
            W = num(s, "cx", W);
            H = num(s, "cy", H);
          }
        }
        const ctx: ParseCtx = { zip, phCache: new Map(), W, H };

        const names = Object.keys(zip.files)
          .filter((n) => /^ppt\/slides\/slide\d+\.xml$/.test(n))
          .sort((a, b) => +a.match(/slide(\d+)/)![1] - +b.match(/slide(\d+)/)![1]);
        const parsed: Slide[] = [];
        for (const n of names) parsed.push(await parseSlide(ctx, n));
        setSlides(parsed);
      } catch (e) {
        setSlides([{ W: 12192000, H: 6858000, bgColor: "#fff", shapes: [], runs: ["解析失败：" + (e as Error).message], xml: "" }]);
      } finally {
        setLoading(false);
      }
    })();
  }, [file]);

  function addNote(e: React.MouseEvent) {
    const rect = (e.currentTarget as HTMLElement).getBoundingClientRect();
    const isStage = e.currentTarget === stageRef.current;
    setNotePrompt({
      x: e.clientX,
      y: e.clientY,
      nx: isStage ? (e.clientX - rect.left) / rect.width : 0.5,
      ny: isStage ? (e.clientY - rect.top) / rect.height : 0.15,
    });
  }

  // 受限文本替换：按文档顺序把编辑后的 runs 写回各 slide 的 <a:t>（DOM 序列化，自闭合/转义自动正确）
  async function handleSave() {
    if (!zipRef.current) return;
    setSaving(true);
    try {
      for (let i = 0; i < slides.length; i++) {
        const slideName = `ppt/slides/slide${i + 1}.xml`;
        const f = zipRef.current.file(slideName);
        if (!f) continue;
        const xml = await f.async("string");
        const doc = new DOMParser().parseFromString(xml, "application/xml");
        const ts = Array.from(doc.getElementsByTagName("*")).filter(
          (el) => el.localName === "t" && el.namespaceURI === A_NS
        );
        ts.forEach((el, k) => {
          const v = slides[i].runs[k];
          if (v !== undefined) el.textContent = v;
        });
        zipRef.current.file(slideName, new XMLSerializer().serializeToString(doc));
      }
      const out = await zipRef.current.generateAsync({ type: "uint8array" });
      await saveFileDialog(file.name, out);
    } finally {
      setSaving(false);
    }
  }

  const slide = slides[page];

  // SVG 用户单位用 pt（1pt=12700 EMU）。不能用原始 EMU：Blink 对 CSS font-size 有 10000px 上限，
  // EMU 级字号（如 44pt=558800）会被钳到 10000 用户单位，导致文字渲染缩小数十倍（表现为画布上的小短横）。
  const PT = 12700;
  const LINE_H = 1.25;

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
      <div style={{ flex: 1, minHeight: 0, display: "flex" }}>
        <div style={{ flex: 1, minWidth: 0, position: "relative", overflowY: "auto", padding: 24, background: "var(--muted)", display: "flex", justifyContent: "center" }}>
          <div
            ref={stageRef}
            onClick={addNote}
            style={{ width: "100%", maxWidth: 960, height: "fit-content", position: "relative", cursor: "crosshair", boxShadow: "var(--shadow)", background: slide?.bgColor }}
          >
            {loading && <div style={{ padding: 24, color: "var(--muted-fg)" }}>正在解析演示文稿…</div>}
            {slide && (
              <svg viewBox={`0 0 ${slide.W / PT} ${slide.H / PT}`} style={{ width: "100%", display: "block", background: slide.bgColor }}>
                {slide.shapes.map((s, i) => {
                  // 形状几何换算为 pt 单位；字号 sz（百分点数）→ pt
                  const gx = s.x / PT, gy = s.y / PT, gw = s.cx / PT, gh = s.cy / PT;
                  const size = s.fontSize / 100;
                  const lineH = size * LINE_H;
                  const anchorX = s.align === "center" ? gx + gw / 2 : s.align === "right" ? gx + gw * 0.98 : gx + gw * 0.02;
                  const anchor = s.align === "center" ? "middle" : s.align === "right" ? "end" : "start";
                  const firstBaseline = gy + gh / 2 - (s.lines.length * lineH) / 2 + size * 0.85;

                  if (s.isPic) {
                    return <image key={i} x={gx} y={gy} width={gw} height={gh} href={s.img} preserveAspectRatio="none" />;
                  }
                  if (s.table) {
                    const nCols = s.table.rows[0]?.length || 0;
                    const sumW = s.table.colW.reduce((a, b) => a + b, 0);
                    const rowH = gh / (s.table.rows.length || 1);
                    let colX = gx;
                    const rects: React.ReactNode[] = [];
                    const texts: React.ReactNode[] = [];
                    for (let c = 0; c < nCols; c++) {
                      const w = sumW > 0 ? (gw * (s.table.colW[c] || 0)) / sumW : gw / nCols;
                      rects.push(
                        <rect key={`r${c}`} x={colX} y={gy} width={w} height={gh} fill="none" stroke="#808080" strokeOpacity={0.55} strokeWidth={1} />
                      );
                      for (let r = 0; r < s.table.rows.length; r++) {
                        const cell = s.table.rows[r][c] || "";
                        const cellLines = cell.split("\n");
                        const cellBaseline = gy + r * rowH + rowH / 2 - (cellLines.length * lineH) / 2 + size * 0.85;
                        texts.push(
                          <text key={`t${r}-${c}`} x={colX + w * 0.03} y={cellBaseline} fill={s.color} fontSize={size} fontFamily="PingFang SC, Microsoft YaHei, sans-serif">
                            {cellLines.map((ln, j) => (
                              <tspan key={j} x={colX + w * 0.03} dy={j === 0 ? 0 : lineH}>{ln || " "}</tspan>
                            ))}
                          </text>
                        );
                      }
                      colX += w;
                    }
                    // 行分隔线
                    for (let r = 1; r < s.table.rows.length; r++) {
                      rects.push(
                        <line key={`h${r}`} x1={gx} y1={gy + r * rowH} x2={gx + gw} y2={gy + r * rowH} stroke="#808080" strokeOpacity={0.55} strokeWidth={1} />
                      );
                    }
                    return (
                      <g key={i}>
                        {rects}
                        {texts}
                      </g>
                    );
                  }
                  if (s.placeholder) {
                    return (
                      <g key={i}>
                        <rect x={gx} y={gy} width={gw} height={gh} fill="rgba(217,163,0,0.05)" stroke="#d9a300" strokeOpacity={0.55} strokeWidth={2} strokeDasharray="8 6" rx={6} />
                        <text x={gx + gw / 2} y={gy + gh / 2 + size * 0.35} fill={s.color} fontSize={size} textAnchor="middle" fontFamily="PingFang SC, Microsoft YaHei, sans-serif">
                          {s.placeholder}
                        </text>
                      </g>
                    );
                  }
                  return (
                    <text
                      key={i}
                      x={anchorX}
                      y={firstBaseline}
                      fill={s.color}
                      fontSize={size}
                      fontWeight={s.bold ? 700 : 400}
                      textAnchor={anchor}
                      fontFamily="PingFang SC, Microsoft YaHei, sans-serif"
                    >
                      {s.lines.map((ln, j) => (
                        <tspan key={j} x={anchorX} dy={j === 0 ? 0 : lineH}>{ln || " "}</tspan>
                      ))}
                    </text>
                  );
                })}
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
        {/* 文本替换面板：右侧固定宽度、独立滚动，不参与挤压画布 */}
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
      {notePrompt && (
        <InlinePrompt
          x={notePrompt.x}
          y={notePrompt.y}
          placeholder="批注内容（会话内显示）"
          onOk={(text) => {
            setNotes((n) => [...n, { page, x: notePrompt.nx, y: notePrompt.ny, text }]);
            setNotePrompt(null);
          }}
          onCancel={() => setNotePrompt(null)}
        />
      )}
    </div>
  );
}
