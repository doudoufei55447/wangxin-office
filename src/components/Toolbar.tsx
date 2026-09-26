import { FolderOpen, Sun, Moon, Save } from "lucide-react";

export function Toolbar({
  onOpen,
  theme,
  onToggleTheme,
}: {
  onOpen: () => void;
  theme: "light" | "dark";
  onToggleTheme: () => void;
}) {
  return (
    <header
      style={{
        height: 48,
        display: "flex",
        alignItems: "center",
        gap: 8,
        padding: "0 12px",
        background: "var(--card)",
        borderBottom: "1px solid var(--border)",
        flexShrink: 0,
      }}
    >
      <div style={{ flex: 1 }} />

      <button className="btn" onClick={onOpen}>
        <FolderOpen size={16} /> 打开
      </button>
      <button className="icon-btn" title="保存" onClick={() => alert("请在文档视图中使用「保存 / 另存为」")}>
        <Save size={18} />
      </button>
      <span className="separator" />
      <button className="icon-btn" title={theme === "light" ? "切换深色" : "切换浅色"} onClick={onToggleTheme}>
        {theme === "light" ? <Moon size={18} /> : <Sun size={18} />}
      </button>
    </header>
  );
}
