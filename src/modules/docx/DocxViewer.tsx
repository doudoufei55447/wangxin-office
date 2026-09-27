import { useEffect, useRef, useState } from "react";
import mammoth from "mammoth";
import { Document, Paragraph, TextRun, Packer, HeadingLevel, Table, TableRow, TableCell } from "docx";
import { Save, FileWarning } from "lucide-react";
import { saveFileDialog, isLegacyBinaryFormat } from "../../platform";
import { extractDocParagraphs } from "./docText";
import type { DocBlock } from "./docText";
import type { OpenedFile } from "../../platform";

export function DocxViewer({ file }: { file: OpenedFile }) {
  const [loading, setLoading] = useState(true);
  const [saving, setSaving] = useState(false);
  const [note, setNote] = useState<string>("");
  const editorRef = useRef<HTMLDivElement>(null);
  const originalBuffer = useRef<ArrayBuffer>(file.buffer);

  useEffect(() => {
    (async () => {
      setLoading(true);
      try {
        if (isLegacyBinaryFormat(file.buffer)) {
          // 旧版 Word 97-2003 二进制（.doc / WPS 旧格式）：文字级提取（FIB → CLX 分片表），
          // 保文字不保排版，渲染为段落进 contentEditable，可编辑并复用「另存为 docx」。
          try {
            const blocks: DocBlock[] = extractDocParagraphs(file.buffer);
            if (editorRef.current) {
              editorRef.current.innerHTML = "";
              for (const blk of blocks) {
                if (blk.type === "p") {
                  const el = document.createElement("p");
                  el.textContent = blk.text || "\u00a0"; // 空段用 nbsp 占位，避免编辑时被合并
                  editorRef.current.appendChild(el);
                } else {
                  editorRef.current.appendChild(buildTableEl(blk.rows));
                }
              }
            }
            setNote(
              ".doc 旧版格式：已按纯文本 + 表格读取（保文字、不保排版），可直接编辑后「另存为 docx」。"
            );
          } catch (e) {
            setNote(
              "解析失败：旧版 Word 97-2003 二进制格式（.doc）读取失败（" +
                (e as Error).message +
                "）。请用 Word/WPS 打开后「另存为 .docx」再打开。"
            );
          }
          return;
        }
        const { value, messages } = await mammoth.convertToHtml({ arrayBuffer: file.buffer.slice(0) });
        if (editorRef.current) {
          editorRef.current.innerHTML = value;
          // mammoth 输出的 <img> 无宽高属性，浏览器按原始像素渲染会溢出页面卡片；
          // 宽度约束由下方 .wx-docx-editor img 样式统一处理。
          // 另对浏览器无法解码的图片（EMF/WMF 矢量图等）降级为友好占位提示。
          editorRef.current.querySelectorAll("img").forEach((img) => {
            img.onerror = () => {
              const ph = document.createElement("div");
              ph.textContent = "该图片格式暂不支持预览（如 EMF/WMF 矢量图），正文内容不受影响";
              ph.setAttribute(
                "style",
                "padding:12px;border:1px dashed #d9a300;border-radius:6px;color:#8a6d1a;font-size:13px;margin:8px 0;background:rgba(217,163,0,0.06)"
              );
              img.replaceWith(ph);
            };
          });
        }
        if (messages.length) setNote(`部分样式/元素未保留（如复杂表格、页眉页脚），属 MVP 已知边界。`);
      } catch (e) {
        setNote("解析失败：" + (e as Error).message);
      } finally {        setLoading(false);
      }
    })();
  }, [file]);

  async function handleSave() {
    if (!editorRef.current) return;
    setSaving(true);
    try {
      const doc = htmlToDocx(editorRef.current);
      const blob = await Packer.toBlob(doc);
      // 输出恒为 docx 格式：.doc 旧文件另存时自动换成 .docx 后缀，避免产生「doc 后缀实为 zip」的坏文件
      const outName = file.name.replace(/\.(doc|docx)$/i, "") + ".docx";
      await saveFileDialog(outName, blob);
    } finally {
      setSaving(false);
    }
  }

  return (
    <div style={{ height: "100%", display: "flex", flexDirection: "column" }}>
      {/* 约束文档内图片宽度：不超出页面卡片，等比缩放（修复大图横向溢出截断） */}
      <style>{".wx-docx-editor img{max-width:100%;height:auto;border-radius:4px;}"}</style>
      <div style={{ display: "flex", gap: 8, alignItems: "center", padding: "8px 12px", background: "var(--card)", borderBottom: "1px solid var(--border)" }}>
        <button className="btn primary" disabled={saving} onClick={handleSave}>
          <Save size={16} /> {saving ? "保存中…" : "保存 / 另存为 docx"}
        </button>
        {note && <span style={{ display: "flex", gap: 4, alignItems: "center", fontSize: 12, color: "var(--warning)" }}><FileWarning size={14} />{note}</span>}
      </div>
      <div style={{ flex: 1, position: "relative", overflowY: "auto", padding: 24, background: "var(--muted)" }}>
        {/* 编辑器容器：React 不渲染任何子节点（避免 innerHTML 与 React 协调器冲突导致整窗白屏） */}
        <div
          ref={editorRef}
          className="wx-docx-editor"
          contentEditable
          suppressContentEditableWarning
          style={{
            background: "var(--card)",
            maxWidth: 820,
            margin: "0 auto",
            padding: "48px 56px",
            borderRadius: 8,
            boxShadow: "var(--shadow)",
            minHeight: "100%",
            outline: "none",
            lineHeight: 1.8,
          }}
        />
        {loading && (
          <div style={{ position: "absolute", top: 32, left: 0, right: 0, textAlign: "center", color: "var(--muted-fg)" }}>
            正在解析文档…
          </div>
        )}
        {!loading && !!note && note.startsWith("解析失败") && (
          <div style={{ position: "absolute", top: 32, left: 0, right: 0, padding: "0 24px", textAlign: "center", color: "var(--warning)", fontSize: 13 }}>
            {note}
          </div>
        )}
      </div>
    </div>
  );
}

// 在 contentEditable 容器内渲染表格（保网格，便于预览与编辑）
function buildTableEl(rows: string[][]): HTMLTableElement {
  const table = document.createElement("table");
  table.setAttribute(
    "style",
    "border-collapse:collapse;width:100%;margin:10px 0;font-size:14px;table-layout:fixed;"
  );
  for (const r of rows) {
    const tr = document.createElement("tr");
    for (const c of r) {
      const td = document.createElement("td");
      td.setAttribute(
        "style",
        "border:1px solid #d0d5dd;padding:6px 10px;vertical-align:top;word-break:break-word;"
      );
      td.textContent = c || "\u00a0";
      tr.appendChild(td);
    }
    table.appendChild(tr);
  }
  return table;
}

// 轻量 HTML→docx：提取段落/标题/表格文本（MVP 范围，复杂排版不保真）
function htmlToDocx(root: HTMLElement): Document {
  const blocks: (Paragraph | Table)[] = [];
  const pushPara = (text: string, heading?: (typeof HeadingLevel)[keyof typeof HeadingLevel]) => {
    const t = text.trim();
    if (!t) return;
    blocks.push(new Paragraph({ text: t, heading, spacing: { after: 120 } }));
  };
  const handleTable = (tbl: HTMLTableElement) => {
    const rows: TableRow[] = [];
    tbl.querySelectorAll("tr").forEach((tr) => {
      const cells: TableCell[] = [];
      tr.querySelectorAll("td,th").forEach((td) => {
        const t = (td.textContent || "").replace(/ /g, " ").trim();
        cells.push(new TableCell({ children: [new Paragraph(t || " ")] }));
      });
      if (cells.length) rows.push(new TableRow({ children: cells }));
    });
    if (rows.length) blocks.push(new Table({ rows }));
  };
  const walk = (node: Node) => {
    node.childNodes.forEach((child) => {
      if (child.nodeType !== Node.ELEMENT_NODE) return;
      const el = child as HTMLElement;
      const tag = el.tagName.toLowerCase();
      if (tag === "table") {
        handleTable(el as HTMLTableElement);
        return;
      }
      if (["p", "h1", "h2", "h3", "h4", "li"].includes(tag)) {
        let heading: (typeof HeadingLevel)[keyof typeof HeadingLevel] | undefined;
        if (tag === "h1") heading = HeadingLevel.HEADING_1;
        else if (tag === "h2") heading = HeadingLevel.HEADING_2;
        else if (tag === "h3") heading = HeadingLevel.HEADING_3;
        pushPara(el.textContent || "", heading);
        return;
      }
      if (tag === "div") {
        // 含块级子元素（表格/段落）则递归，否则按段落处理
        const hasBlock = Array.from(el.children).some((c) =>
          ["p", "table", "div", "h1", "h2", "h3", "h4", "li"].includes(
            (c as HTMLElement).tagName.toLowerCase()
          )
        );
        if (hasBlock) {
          walk(el);
          return;
        }
        pushPara(el.textContent || "");
        return;
      }
      walk(el);
    });
  };
  walk(root);
  if (blocks.length === 0) blocks.push(new Paragraph({ children: [new TextRun("（空文档）")] }));
  return new Document({ sections: [{ children: blocks }] });
}
