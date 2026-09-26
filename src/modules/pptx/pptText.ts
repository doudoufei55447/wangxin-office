// 旧版 PowerPoint 97-2003 二进制格式（.ppt / WPS 以 ppt 后缀保存的旧格式）文本级提取。
// 原理：OLE 容器(CFB) → PowerPoint Document 流 → 顶层记录扫描：
//   1) 合并全部 PersistPtr 块(6001 全量/6002 增量)得到 persistId→偏移 映射（流序后者覆盖前者，天然处理增量保存）
//   2) DocumentContainer(1000) 内 SlideListWithText(4026) 的 SlidePersistAtom(4057) 给出幻灯片顺序
//   3) 只取 persistMap 指向的当前版本 SlideContainer(1006)，递归收集 TextCharsAtom(4000,UTF-16)/TextBytesAtom(4008,cp1252)
// 输出每页段落数组（保文字、不保排版）。
import * as XLSX from "xlsx";

const RT_SLIDE = 1006;
const RT_DOCUMENT = 1000;
const RT_DOCUMENT_ATOM = 1001; // body 前 8 字节 = slideSizeX/slideSizeY (EMU)
const RT_SLIDE_LIST = 4026;
const RT_SLIDE_PERSIST = 4057;
const RT_TEXTCHARS = 4000; // UTF-16LE
const RT_TEXTBYTES = 4008; // 单字节 cp1252
const RT_PERSIST_FULL = 6001;
const RT_PERSIST_INCR = 6002;

const u16 = (a: Uint8Array, p: number) => a[p] | (a[p + 1] << 8);
const u32 = (a: Uint8Array, p: number) => (a[p] | (a[p + 1] << 8) | (a[p + 2] << 16) | (a[p + 3] << 24)) >>> 0;

interface Rec { off: number; ver: number; type: number; len: number; body: number; }

function scanRecords(buf: Uint8Array, start: number, end: number, out: Rec[]): boolean {
  let pos = start;
  while (pos + 8 <= end) {
    const ver = buf[pos] & 0x0f;
    const type = u16(buf, pos + 2);
    const len = u32(buf, pos + 4);
    const body = pos + 8;
    if (body + len > end) return false; // 越界：流损坏，到此为止
    out.push({ off: pos, ver, type, len, body });
    pos = body + len;
  }
  return true;
}

function collectTexts(buf: Uint8Array, start: number, end: number, out: string[], depth = 0) {
  if (depth > 16) return;
  let pos = start;
  while (pos + 8 <= end) {
    const ver = buf[pos] & 0x0f;
    const type = u16(buf, pos + 2);
    const len = u32(buf, pos + 4);
    const body = pos + 8;
    if (body + len > end) return;
    if (ver === 0xf) {
      collectTexts(buf, body, body + len, out, depth + 1);
    } else if (type === RT_TEXTCHARS) {
      out.push(...decodePptText(new TextDecoder("utf-16le").decode(buf.slice(body, body + len))));
    } else if (type === RT_TEXTBYTES) {
      out.push(...decodePptText(new TextDecoder("windows-1252").decode(buf.slice(body, body + len))));
    }
    pos = body + len;
  }
}

function decodePptText(t: string): string[] {
  // \r 段落尾 / \v(0x0B) 软换行 / \x07 表格单元格尾 → 分段；其余控制符丢弃（保留 \t 制表）
  return t
    .split(/[\r\x0b\x07]/)
    .map((s) => s.replace(/[\x00-\x08\x0e-\x1f]/g, "").replace(/\u00a0/g, " ").trim())
    .filter(Boolean);
}

export function extractPptSlides(buffer: ArrayBuffer): { W: number; H: number; slides: string[][] } {
  const cfb = XLSX.CFB.read(new Uint8Array(buffer.slice(0)), { type: "array" });
  const entry = XLSX.CFB.find(cfb, "PowerPoint Document");
  if (!entry) throw new Error("缺少 PowerPoint Document 流");
  const buf = new Uint8Array(entry.content);

  // 顶层记录扫描 + persistId→偏移 映射（全部块按流序合并，后者覆盖前者）
  const top: Rec[] = [];
  scanRecords(buf, 0, buf.length, top);
  const persistMap = new Map<number, number>();
  for (const r of top) {
    if (r.type === RT_PERSIST_INCR) {
      const first = u32(buf, r.body);
      const count = (r.len - 4) / 4;
      if (Number.isInteger(count)) for (let i = 0; i < count; i++) persistMap.set(first + i, u32(buf, r.body + 4 + i * 4));
    } else if (r.type === RT_PERSIST_FULL) {
      for (let o = 0; o + 8 <= r.len; o += 8) persistMap.set(u32(buf, r.body + o), u32(buf, r.body + o + 4));
    }
  }
  const byOffset = new Map<number, string[]>(); // 偏移 → 该页文本
  const slideRecs = top.filter((r) => r.type === RT_SLIDE);
  for (const s of slideRecs) {
    const texts: string[] = [];
    collectTexts(buf, s.body, s.body + s.len, texts);
    byOffset.set(s.off, texts);
  }

  // 页序：DocumentContainer → SlideListWithText → SlidePersistAtom.persistIdRef → 当前版本 Slide 偏移
  const ordered: string[][] = [];
  const doc = top.find((r) => r.type === RT_DOCUMENT);
  let W = 0, H = 0;
  if (doc) {
    const inners: Rec[] = [];
    scanRecords(buf, doc.body, doc.body + doc.len, inners);
    // 幻灯片版面尺寸来自 DocumentAtom(1001)（EMU）
    const docAtom = inners.find((r) => r.type === RT_DOCUMENT_ATOM);
    if (docAtom && docAtom.len >= 8) {
      W = u32(buf, docAtom.body);
      H = u32(buf, docAtom.body + 4);
    }
    for (const list of inners) {
      if (list.type !== RT_SLIDE_LIST) continue;
      const atoms: Rec[] = [];
      scanRecords(buf, list.body, list.body + list.len, atoms);
      for (const a of atoms) {
        if (a.type !== RT_SLIDE_PERSIST || a.len < 4) continue;
        const off = persistMap.get(u32(buf, a.body));
        if (off === undefined || u16(buf, off + 2) !== RT_SLIDE) continue;
        const t = byOffset.get(off);
        if (t) { ordered.push(t); byOffset.delete(off); }
      }
    }
  }
  // 回退/兜底：流序剩余的当前版本 Slide（SlideListWithText 缺失或解析失败时）
  for (const s of slideRecs) {
    if (persistMap.size && ![...persistMap.values()].includes(s.off)) continue;
    const t = byOffset.get(s.off);
    if (t) { ordered.push(t); byOffset.delete(s.off); }
  }
  if (W <= 0 || H <= 0) { W = 9144000; H = 6858000; } // .ppt 常见 4:3 兜底
  return { W, H, slides: ordered };
}
