import { useEffect, useRef, useState } from "react";
import mammoth from "mammoth";
import { Document, Paragraph, TextRun, Packer, HeadingLevel } from "docx";
import { Save, FileWarning } from "lucide-react";
import { saveFileDialog, isLegacyBinaryFormat } from "../../platform";
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
          setNote(
            "这是旧版 Word 97-2003 二进制格式（.doc）或 WPS 以 docx 后缀保存的旧格式，本版本暂不支持。请用 Word/WPS 打开后「另存为 .docx」再打开（旧版 .doc 引擎计划二期支持）。"
          );
          return;
        }
        const { value, messages } = await mammoth.convertToHtml({ arrayBuffer: file.buffer.slice(0) });
        if (editorRef.current) {
          editorRef.current.innerHTML = value;
        }
        if (messages.length) setNote(`部分样式/元素未保留（如复杂表格、页眉页脚），属 MVP 已知边界。`);
      } catch (e) {
        setNote("解析失败：" + (e as Error).message);
      } finally {
        setLoading(false);
      }
    })();
  }, [file]);

  async function handleSave() {
    if (!editorRef.current) return;
    setSaving(true);
    try {
      const doc = htmlToDocx(editorRef.current);
      const blob = await Packer.toBlob(doc);
      await saveFileDialog(file.name, blob);
    } finally {
      setSaving(false);
    }
  }

  return (
    <div style={{ height: "100%", display: "flex", flexDirection: "column" }}>
      <div style={{ display: "flex", gap: 8, alignItems: "center", padding: "8px 12px", background: "var(--card)", borderBottom: "1px solid var(--border)" }}>
        <button className="btn primary" disabled={saving} onClick={handleSave}>
          <Save size={16} /> {saving ? "保存中…" : "保存 / 另存为 docx"}
        </button>
        {note && <span style={{ display: "flex", gap: 4, alignItems: "center", fontSize: 12, color: "var(--warning)" }}><FileWarning size={14} />{note}</span>}
      </div>
      <div style={{ flex: 1, overflowY: "auto", padding: 24, background: "var(--muted)" }}>
        <div
          ref={editorRef}
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
        >
          {loading && <span style={{ color: "var(--muted-fg)" }}>正在解析文档…</span>}
        </div>
      </div>
    </div>
  );
}

// 轻量 HTML→docx：仅提取段落/标题/列表文本（MVP 范围，复杂排版不保真）
function htmlToDocx(root: HTMLElement): Document {
  const paras: Paragraph[] = [];
  const walk = (node: Node) => {
    node.childNodes.forEach((child) => {
      if (child.nodeType === Node.ELEMENT_NODE) {
        const el = child as HTMLElement;
        const tag = el.tagName.toLowerCase();
        if (["p", "h1", "h2", "h3", "h4", "li", "div"].includes(tag)) {
          const text = el.textContent?.trim() || "";
          if (!text) return;
          let heading: HeadingLevel | undefined;
          if (tag === "h1") heading = HeadingLevel.HEADING_1;
          else if (tag === "h2") heading = HeadingLevel.HEADING_2;
          else if (tag === "h3") heading = HeadingLevel.HEADING_3;
          paras.push(new Paragraph({ text, heading, spacing: { after: 120 } }));
        } else {
          walk(child);
        }
      }
    });
  };
  walk(root);
  if (paras.length === 0) paras.push(new Paragraph({ children: [new TextRun("（空文档）")] }));
  return new Document({ sections: [{ children: paras }] });
}
