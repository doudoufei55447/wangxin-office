// 旧版 Word 97-2003 二进制格式（.doc / WPS 以 doc 后缀保存的旧格式）纯文本提取。
// 原理：OLE Compound File（CFB）→ WordDocument 流 FIB → 表流 CLX 分片表（PlcPcd）→ 按分片解码正文。
// 输出为文字级段落（保文字、不保排版），供 DocxViewer 渲染并复用「另存为 docx」链路。
import * as XLSX from "xlsx";

const u16 = (a: Uint8Array, p: number) => a[p] | (a[p + 1] << 8);
const u32 = (a: Uint8Array, p: number) => (a[p] | (a[p + 1] << 8) | (a[p + 2] << 16) | (a[p + 3] << 24)) >>> 0;

function decodeText(bytes: Uint8Array, compressed: boolean): string {
  // 压缩分片是 Windows-1252 单字节；未压缩分片是 UTF-16LE
  return new TextDecoder(compressed ? "windows-1252" : "utf-16le").decode(bytes);
}

// CLX：一串 Prc（clxt=1，跳过）+ 一个 Pcdt（clxt=2，含 PlcPcd）
function parseClx(tbl: Uint8Array, fcClx: number, lcbClx: number): { off: number; lcb: number } | null {
  let pos = fcClx;
  const end = fcClx + lcbClx;
  while (pos < end) {
    const t = tbl[pos];
    if (t === 1) {
      pos += 3 + u16(tbl, pos + 1);
    } else if (t === 2) {
      return { off: pos + 5, lcb: u32(tbl, pos + 1) };
    } else break;
  }
  return null;
}

export function extractDocParagraphs(buffer: ArrayBuffer): string[] {
  const cfb = XLSX.CFB.read(new Uint8Array(buffer.slice(0)), { type: "array" });
  const wdEntry = XLSX.CFB.find(cfb, "WordDocument");
  if (!wdEntry) throw new Error("缺少 WordDocument 流");
  const wd = new Uint8Array(wdEntry.content);
  if (u16(wd, 0) !== 0xa5ec) throw new Error("wIdent 校验失败，非 Word 97-2003 格式");
  const flags = u16(wd, 0x0a);
  const tblEntry = XLSX.CFB.find(cfb, flags & 0x0200 ? "1Table" : "0Table");
  const tbl = tblEntry ? new Uint8Array(tblEntry.content) : null;

  // 正文分片表：fcClx/lcbClx 位于 FIB 的 FibRgFcLcb97（约定偏移 0x01A2/0x01A6）
  const fcClx = u32(wd, 0x01a2);
  const lcbClx = u32(wd, 0x01a6);
  let text: string;
  const clx = tbl && lcbClx > 0 ? parseClx(tbl, fcClx, lcbClx) : null;
  if (clx) {
    const n = (clx.lcb - 4) / 12;
    if (!Number.isInteger(n) || n < 1) throw new Error("分片表尺寸异常");
    const cpsOff = clx.off;
    const pcdsOff = clx.off + (n + 1) * 4;
    const parts: string[] = [];
    for (let i = 0; i < n; i++) {
      const len = u32(tbl, cpsOff + (i + 1) * 4) - u32(tbl, cpsOff + i * 4);
      const fc = u32(tbl, pcdsOff + i * 8 + 2);
      const compressed = !!(fc & 0x40000000);
      const byteOff = compressed ? (fc & 0x3fffffff) >> 1 : fc & 0x3fffffff;
      parts.push(decodeText(wd.slice(byteOff, byteOff + (compressed ? len : len * 2)), compressed));
    }
    text = parts.join("");
  } else {
    // 无分片表回退：FIB fcMin/fcMac 直读正文
    const ext = !!(flags & 0x1000);
    text = decodeText(wd.slice(u32(wd, 0x18), u32(wd, 0x1c)), !ext);
  }

  // 清理：0x0D 段落尾 / 0x07 表格单元格尾 / 0x0C 分节 / 0x0B 软换行 → 分段；0x09 制表保留；其余 C0 丢弃。
  // Word 域状态机：0x13 域开始→丢弃指令文本（如 TOC/HYPERLINK/PAGEREF），0x14 分隔→保留结果文本，0x15 域结束（可嵌套）。
  const cleaned: string[] = [];
  let fieldDepth = 0;
  const inResult: boolean[] = [];
  for (const ch of text) {
    const c = ch.codePointAt(0)!;
    if (c === 0x13) { fieldDepth++; inResult[fieldDepth] = false; continue; }
    if (c === 0x14) { if (fieldDepth) inResult[fieldDepth] = true; continue; }
    if (c === 0x15) { if (fieldDepth) fieldDepth--; continue; }
    if (fieldDepth && !inResult[fieldDepth]) continue;
    if (c === 0x0d || c === 0x07 || c === 0x0c || c === 0x0b) cleaned.push("\n");
    else if (c === 0x09 || c >= 0x20) cleaned.push(ch);
    else if (c === 0x1e) cleaned.push("-");
    // 0x01 图片/OLE、0x02 脚注引用、0x05 批注引用等标记直接丢弃
  }
  return cleaned
    .join("")
    .split("\n")
    .map((p) => p.replace(/\u00a0/g, " ").trim());
}
