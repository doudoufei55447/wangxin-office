import { X, FileText, FileSpreadsheet, Presentation, FileType2 } from "lucide-react";
import type { TabItem } from "../App";

function extIcon(ext: string) {
  if (ext === "pdf") return <FileType2 size={14} />;
  if (["doc", "docx"].includes(ext)) return <FileText size={14} />;
  if (["xls", "xlsx"].includes(ext)) return <FileSpreadsheet size={14} />;
  if (["ppt", "pptx"].includes(ext)) return <Presentation size={14} />;
  return <FileText size={14} />;
}

export function Tabs({
  tabs,
  activeId,
  onSelect,
  onClose,
}: {
  tabs: TabItem[];
  activeId: string | null;
  onSelect: (id: string) => void;
  onClose: (id: string) => void;
}) {
  return (
    <div
      style={{
        display: "flex",
        gap: 2,
        background: "var(--card)",
        borderBottom: "1px solid var(--border)",
        padding: "6px 8px 0",
        overflowX: "auto",
        minHeight: 40,
        alignItems: "flex-end",
      }}
    >
      {tabs.map((t) => {
        const active = t.id === activeId;
        return (
          <div
            key={t.id}
            onClick={() => onSelect(t.id)}
            style={{
              display: "flex",
              alignItems: "center",
              gap: 6,
              padding: "6px 8px 6px 12px",
              borderRadius: "8px 8px 0 0",
              background: active ? "var(--muted)" : "transparent",
              color: active ? "var(--fg)" : "var(--muted-fg)",
              fontSize: 13,
              maxWidth: 200,
              cursor: "pointer",
              borderTop: active ? "2px solid var(--primary)" : "2px solid transparent",
            }}
          >
            {extIcon(t.file.ext)}
            <span style={{ overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{t.file.name}</span>
            <span className="icon-btn" style={{ width: 20, height: 20 }} onClick={(e) => { e.stopPropagation(); onClose(t.id); }}>
              <X size={13} />
            </span>
          </div>
        );
      })}
    </div>
  );
}
