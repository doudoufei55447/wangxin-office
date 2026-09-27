// .pptx 图形对象（图表 c:chart / SmartArt dgm:diagram）数据提取。
// 目的：对无法在纯前端绘制原生图形的对象，提取其「缓存数据 / 文本」，以数据表格或文本框呈现，
//       避免只显示「暂不支持预览」而丢失内容。
export type ChartKind = "bar" | "column" | "line" | "pie" | "doughnut" | "area" | "other";
export interface ChartSeries { name: string; vals: number[]; }
export interface ChartData { kind: ChartKind; title: string; cats: string[]; series: ChartSeries[]; }

const A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";

function byLocal(root: Element | Document, name: string): Element[] {
  const all = root.getElementsByTagName("*");
  const out: Element[] = [];
  for (let i = 0; i < all.length; i++) if (all[i].localName === name) out.push(all[i]);
  return out;
}

function parseXml(xml: string): Document | null {
  try {
    const d = new DOMParser().parseFromString(xml, "application/xml");
    return d.getElementsByTagName("parsererror").length ? null : d;
  } catch {
    return null;
  }
}

// 解析 chartN.xml：类型 / 标题 / 类别（c:cat 缓存）/ 各系列（c:tx 名称 + c:val 数值缓存）
export function parseChartXml(xml: string): ChartData | null {
  const doc = parseXml(xml);
  if (!doc) return null;
  const has = (n: string) => byLocal(doc, n).length > 0;
  let kind: ChartKind = "other";
  if (has("pieChart")) kind = "pie";
  else if (has("doughnutChart")) kind = "doughnut";
  else if (has("barChart")) kind = byLocal(doc, "barDir")[0]?.getAttribute("val") === "bar" ? "bar" : "column";
  else if (has("lineChart")) kind = "line";
  else if (has("areaChart")) kind = "area";

  let title = "";
  const titleEl = byLocal(doc, "title")[0];
  if (titleEl) {
    title = byLocal(titleEl, "t")
      .map((e) => e.textContent || "")
      .join("")
      .trim();
    if (!title) title = (byLocal(titleEl, "v")[0]?.textContent || "").trim();
  }

  const catEl = byLocal(doc, "cat")[0];
  let cats: string[] = [];
  if (catEl) {
    cats = byLocal(catEl, "pt")
      .sort((a, b) => +(a.getAttribute("idx") || 0) - +(b.getAttribute("idx") || 0))
      .map((pt) => byLocal(pt, "v")[0]?.textContent || "");
  }

  const series: ChartSeries[] = [];
  for (const ser of byLocal(doc, "ser")) {
    const tx = byLocal(ser, "tx")[0];
    const name = tx
      ? (byLocal(tx, "v")[0]?.textContent || byLocal(tx, "t").map((e) => e.textContent || "").join("")).trim()
      : "";
    const valEl = byLocal(ser, "val")[0];
    const vals = valEl
      ? byLocal(valEl, "pt")
          .sort((a, b) => +(a.getAttribute("idx") || 0) - +(b.getAttribute("idx") || 0))
          .map((pt) => {
            const raw = byLocal(pt, "v")[0]?.textContent || "";
            return raw === "" ? NaN : +raw;
          })
      : [];
    if (!vals.length) continue; // 空系列（占位/无缓存数据）跳过，避免污染数据表格
    series.push({ name, vals });
  }
  if (!series.length || series.every((s) => s.vals.length === 0)) return null;
  const maxLen = Math.max(...series.map((s) => s.vals.length));
  if (!cats.length) cats = Array.from({ length: maxLen }, (_, i) => String(i + 1));
  return { kind, title, cats, series };
}

// 解析 SmartArt 的 data 部件（ppt/diagrams/dataN.xml）：按文档顺序取 a:t 文本，去连续重复
export function parseDiagramText(xml: string): string[] {
  const doc = parseXml(xml);
  if (!doc) return [];
  const out: string[] = [];
  for (const t of byLocal(doc, "t")) {
    if (t.namespaceURI !== A_NS) continue;
    const v = (t.textContent || "").replace(/\s+/g, " ").trim();
    if (v && out[out.length - 1] !== v) out.push(v);
  }
  return out;
}
