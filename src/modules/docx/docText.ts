// 旧版 Word 97-2003 二进制格式（.doc / WPS 以 doc 后缀保存的旧格式）提取。
// 原理：OLE Compound File（CFB）→ WordDocument 流 FIB → 表流 CLX 分片表（PlcPcd）解码正文
//      + plcfbtePapx（段落属性 FKP）解析表格结构（单元格 / 行尾标记）。
// 输出为文字级块（段落 + 表格），保文字、不保排版，供 DocxViewer 渲染并复用「另存为 docx」链路。
import * as XLSX from "xlsx";

export type DocBlock =
  | { type: "p"; text: string }
  | { type: "table"; rows: string[][] };

const u16 = (a: Uint8Array, p: number) => a[p] | (a[p + 1] << 8);
const u32 = (a: Uint8Array, p: number) =>
  (a[p] | (a[p + 1] << 8) | (a[p + 2] << 16) | (a[p + 3] << 24)) >>> 0;

function decodeText(bytes: Uint8Array, compressed: boolean): string {
  // 压缩分片是 Windows-1252 单字节；未压缩分片是 UTF-16LE
  return new TextDecoder(compressed ? "windows-1252" : "utf-16le").decode(bytes);
}

// 清理控制字符与 Word 域：保留可见文字；0x07 单元格标记 / 0x0D 段尾 / 0x0C 分节 / 0x0B 软换行丢弃；
// 0x09 制表保留；0x1e→'-'；Word 域状态机（0x13 域始→丢弃指令文本，0x14 分隔→保留结果，0x15 域终，可嵌套）。
function cleanText(raw: string): string {
  let fieldDepth = 0;
  const inResult: boolean[] = [];
  let out = "";
  for (const ch of raw) {
    const c = ch.codePointAt(0)!;
    if (c === 0x13) { fieldDepth++; inResult[fieldDepth] = false; continue; }
    if (c === 0x14) { if (fieldDepth) inResult[fieldDepth] = true; continue; }
    if (c === 0x15) { if (fieldDepth) fieldDepth--; continue; }
    if (fieldDepth && !inResult[fieldDepth]) continue;
    if (c === 0x09 || c >= 0x20) out += ch;
    else if (c === 0x1e) out += "-";
    // 其余控制字符（0x07/0x0D/0x0C/0x0B/0x01 图片/0x02 脚注 等）丢弃
  }
  return out;
}

// CLX：一串 Prc（clxt=1，跳过）+ 一个 Pcdt（clxt=2，含 PlcPcd）→ 返回分片边界 cps 与分片 pieces。
function parseClxPieces(
  tbl: Uint8Array,
  fcClx: number,
  lcbClx: number
): { cps: number[]; pieces: { compressed: boolean; abs: number }[] } | null {
  let pos = fcClx;
  const end = fcClx + lcbClx;
  while (pos < end) {
    const t = tbl[pos];
    if (t === 1) {
      pos += 3 + u16(tbl, pos + 1);
    } else if (t === 2) {
      const lcb = u32(tbl, pos + 1);
      const body = pos + 5;
      const n = (lcb - 4) / 12;
      if (!Number.isInteger(n) || n < 1) return null;
      const cps: number[] = [];
      for (let i = 0; i <= n; i++) cps.push(u32(tbl, body + i * 4));
      const pieces: { compressed: boolean; abs: number }[] = [];
      for (let i = 0; i < n; i++) {
        const fc = u32(tbl, body + (n + 1) * 4 + i * 8 + 2);
        pieces.push({ compressed: !!(fc & 0x40000000), abs: fc & 0x3fffffff });
      }
      return { cps, pieces };
    } else break;
  }
  return null;
}

export function extractDocParagraphs(buffer: ArrayBuffer): DocBlock[] {
  const cfb = XLSX.CFB.read(new Uint8Array(buffer.slice(0)), { type: "array" });
  const wdEntry = XLSX.CFB.find(cfb, "WordDocument");
  if (!wdEntry) throw new Error("缺少 WordDocument 流");
  const wd = new Uint8Array(wdEntry.content);
  if (u16(wd, 0) !== 0xa5ec) throw new Error("wIdent 校验失败，非 Word 97-2003 格式");
  const flags = u16(wd, 0x0a);
  const tblEntry = XLSX.CFB.find(cfb, flags & 0x0200 ? "1Table" : "0Table");
  const tbl = tblEntry ? new Uint8Array(tblEntry.content) : null;

  // ---- 正文分片 → 字符流 + fc→str 映射 ----
  const fcClx = u32(wd, 0x01a2);
  const lcbClx = u32(wd, 0x01a6);
  const parsed = tbl && lcbClx > 0 ? parseClxPieces(tbl, fcClx, lcbClx) : null;
  let full = "";
  const starts: { fc: number; per: number; strStart: number }[] = [];

  if (parsed) {
    const { cps, pieces } = parsed;
    for (let i = 0; i < pieces.length; i++) {
      const p = pieces[i];
      const nChars = cps[i + 1] - cps[i];
      const per = p.compressed ? 1 : 2;
      starts.push({ fc: p.abs, per, strStart: full.length });
      if (p.compressed) {
        const off = p.abs / 2;
        const b = wd.slice(off, off + nChars);
        let s = "";
        for (let k = 0; k < nChars; k++) s += String.fromCharCode(b[k]);
        full += s;
      } else {
        full += new TextDecoder("utf-16le").decode(wd.slice(p.abs, p.abs + nChars * 2));
      }
    }
    starts.push({ fc: 0x7fffffff, per: 1, strStart: full.length }); // 哨兵，便于二分上界
  } else {
    // 无分片表回退：FIB fcMin/fcMac 直读正文
    const ext = !!(flags & 0x1000);
    full = decodeText(wd.slice(u32(wd, 0x18), u32(wd, 0x1c)), !ext);
  }

  const at = (fc: number): number => {
    let lo = 0,
      hi = starts.length - 1,
      best = starts.length - 1;
    while (lo <= hi) {
      const mid = (lo + hi) >> 1;
      if (starts[mid].fc <= fc) {
        best = mid;
        lo = mid + 1;
      } else hi = mid - 1;
    }
    const s = starts[best];
    return s.strStart + Math.floor((fc - s.fc) / s.per);
  };
  const fcText = (a: number, b: number) => full.slice(at(a), at(b));

  // ---- plcfbtePapx → FKP 页 → 段落（含表格标记 fInTable / fTtp）----
  const fcPapx = u32(wd, 0x0102);
  const lcbPapx = u32(wd, 0x0106);
  interface PPara {
    fcA: number;
    fcB: number;
    fInTable: boolean;
    fTtp: boolean;
  }
  const paras: PPara[] = [];

  if (tbl && lcbPapx > 4) {
    const n = (lcbPapx - 4) / 8;
    if (Number.isInteger(n) && n >= 1) {
      const pn: number[] = [];
      for (let i = 0; i < n; i++) pn.push(u32(tbl, fcPapx + (n + 1) * 4 + i * 4));
      const sprmLen = (op: number): { varlen: boolean; len: number } => {
        const spra = (op >> 13) & 7; // sprm 操作数类型取高 3 位
        if (op === 0xc615 || op === 0xd608 || op === 0xd609 || op === 0xd60a) return { varlen: true, len: 0 };
        if (spra <= 2) return { varlen: false, len: 1 };
        if (spra === 3 || spra === 5) return { varlen: false, len: 2 };
        if (spra === 4) return { varlen: false, len: 4 };
        if (spra === 7) return { varlen: false, len: 3 };
        return { varlen: true, len: 0 };
      };
      const parsePAPX = (page: Uint8Array, g0: number, gl: number) => {
        const r = { fInTable: false, fTtp: false };
        let p = g0 + 2; // 跳过 2 字节 istd
        const e = g0 + gl;
        while (p + 2 <= e) {
          const op = u16(page, p);
          const L = sprmLen(op);
          if (L.varlen) {
            const size = u16(page, p + 2);
            p += 4 + size;
          } else {
            if (op === 0x2416) r.fInTable = page[p + 2] !== 0; // sprmPFInTable
            if (op === 0x2417) r.fTtp = page[p + 2] !== 0; // sprmPFTtp（行尾标记）
            p += 2 + L.len;
          }
        }
        return r;
      };
      for (const x of pn) {
        const page = wd.slice(x * 512, x * 512 + 512);
        const crun = page[511];
        if (crun === 0) continue;
        const fcs: number[] = [];
        for (let i = 0; i <= crun; i++) fcs.push(u32(page, i * 4));
        const rgbBase = (crun + 1) * 4;
        for (let i = 0; i < crun; i++) {
          const bx = page[rgbBase + i * 13]; // PAPX FKP：1 字节 BX + 12 字节 PHE，步长 13
          const o = bx * 2;
          const cb = page[o];
          const g0 = cb === 0 ? o + 2 : o + 1;
          const gl = cb === 0 ? page[o + 1] * 2 - 1 : cb * 2 - 1;
          if (gl < 0 || g0 + gl > 512) continue;
          const pr = parsePAPX(page, g0, gl);
          paras.push({ fcA: fcs[i], fcB: fcs[i + 1], fInTable: pr.fInTable, fTtp: pr.fTtp });
        }
      }
    }
  }

  const blocks: DocBlock[] = [];

  // 无段落属性表 → 回退：整篇按行切分
  if (paras.length === 0) {
    full
      .split("\n")
      .map((t) => cleanText(t).trim())
      .filter((t) => t.length > 0)
      .forEach((t) => blocks.push({ type: "p", text: t }));
    return blocks;
  }

  // ---- 块装配：段落与表格混合 ----
  // 结构（WPS/Word 97-2003 常见）：单元格段落以 0x07 单元格标记结尾（自身不带 sprmPFInTable），
  // 每个行尾是一个 fTtp=1 的标记段落（文本仅为 0x07），用于关闭当前行。
  let table: { rows: string[][] } | null = null;
  let currentRow: string[] = [];

  const closeTable = () => {
    if (table && currentRow.length > 0) table.rows.push(currentRow);
    if (table && table.rows.length > 0) {
      const cols = Math.max(...table.rows.map((r) => r.length));
      for (const r of table.rows) while (r.length < cols) r.push(""); // 短行补空，保持网格对齐
      blocks.push({ type: "table", rows: table.rows });
    }
    table = null;
    currentRow = [];
  };

  for (const para of paras) {
    const raw = fcText(para.fcA, para.fcB);
    if (para.fInTable) {
      // 行尾标记段落：关闭当前行（该段自身不贡献单元格）
      if (table && currentRow.length > 0) {
        table.rows.push(currentRow);
        currentRow = [];
      }
      continue;
    }
    if (/[\u0007]$/.test(raw)) {
      // 单元格段落：文本以 0x07 单元格标记结尾
      if (!table) table = { rows: [] };
      const body = raw.replace(/[\u0007\u000d\u000c\u000b]*$/, "");
      currentRow.push(cleanText(body));
      continue;
    }
    // 普通段落
    const t = cleanText(raw).trim();
    if (table && t === "") continue; // 表格区域内的空段（分隔符）忽略，保持表格连续
    closeTable();
    if (t) blocks.push({ type: "p", text: t });
  }
  closeTable();
  return blocks;
}
