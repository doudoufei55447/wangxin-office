import { useEffect, useRef, useState } from "react";

// 内联输入气泡：替代 window.prompt（Tauri WebView 中 prompt/alert/confirm 为空操作）
export function InlinePrompt({
  x, y, placeholder, onOk, onCancel,
}: {
  x: number;
  y: number;
  placeholder: string;
  onOk: (text: string) => void;
  onCancel: () => void;
}) {
  const [val, setVal] = useState("");
  const ref = useRef<HTMLInputElement>(null);
  useEffect(() => {
    ref.current?.focus();
  }, []);
  return (
    <div
      style={{
        position: "fixed", left: Math.min(x, window.innerWidth - 340), top: Math.min(y, window.innerHeight - 60),
        zIndex: 1000, background: "var(--card)", border: "1px solid var(--border)", borderRadius: 8,
        boxShadow: "var(--shadow)", padding: 8, display: "flex", gap: 6, alignItems: "center",
      }}
    >
      <input
        ref={ref}
        value={val}
        placeholder={placeholder}
        onChange={(e) => setVal(e.target.value)}
        onKeyDown={(e) => {
          if (e.key === "Enter" && val.trim()) onOk(val.trim());
          else if (e.key === "Escape") onCancel();
        }}
        style={{ border: "1px solid var(--border)", borderRadius: 6, padding: "4px 8px", background: "var(--muted)", color: "var(--fg)", font: "inherit", width: 200, outline: "none" }}
      />
      <button className="btn primary" onClick={() => val.trim() && onOk(val.trim())}>确定</button>
      <button className="btn" onClick={onCancel}>取消</button>
    </div>
  );
}
