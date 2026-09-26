import { FileText, FileSpreadsheet, Presentation, FileType2, Clock, Plus } from "lucide-react";

function extIcon(ext: string) {
  if (ext === "pdf") return <FileType2 size={16} color="var(--destructive)" />;
  if (["doc", "docx"].includes(ext)) return <FileText size={16} color="var(--primary)" />;
  if (["xls", "xlsx"].includes(ext)) return <FileSpreadsheet size={16} color="var(--success)" />;
  if (["ppt", "pptx"].includes(ext)) return <Presentation size={16} color="var(--warning)" />;
  return <FileText size={16} />;
}

export function Sidebar({ recents, onOpen }: { recents: any[]; onOpen: () => void }) {
  return (
    <aside
      style={{
        width: 240,
        background: "var(--card)",
        borderRight: "1px solid var(--border)",
        display: "flex",
        flexDirection: "column",
        flexShrink: 0,
      }}
    >
      <div style={{ padding: 14 }}>
        <button className="btn primary" style={{ width: "100%", justifyContent: "center" }} onClick={onOpen}>
          <Plus size={16} /> 打开文件
        </button>
      </div>
      <div style={{ padding: "8px 14px", color: "var(--muted-fg)", fontSize: 12, fontWeight: 600, display: "flex", gap: 6, alignItems: "center" }}>
        <Clock size={14} /> 最近文件
      </div>
      <div style={{ flex: 1, overflowY: "auto", padding: "0 8px 12px" }}>
        {recents.length === 0 && <div style={{ padding: 12, fontSize: 12, color: "var(--muted-fg)" }}>暂无记录</div>}
        {recents.map((r, i) => (
          <div
            key={i}
            style={{
              display: "flex",
              gap: 10,
              alignItems: "center",
              padding: "8px 10px",
              borderRadius: "var(--radius-sm)",
              cursor: "default",
            }}
            title={r.path || r.name}
          >
            {extIcon(r.ext)}
            <div style={{ overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", fontSize: 13 }}>{r.name}</div>
          </div>
        ))}
      </div>
    </aside>
  );
}
