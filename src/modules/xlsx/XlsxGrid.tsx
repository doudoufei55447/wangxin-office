import { useEffect, useRef, useState } from "react";
import * as XLSX from "xlsx";
import { Save } from "lucide-react";
import { saveFileDialog } from "../../platform";
import type { OpenedFile } from "../../platform";

interface Sheet { name: string; rows: (string | number)[][]; }

export function XlsxGrid({ file }: { file: OpenedFile }) {
  const [sheets, setSheets] = useState<Sheet[]>([]);
  const [active, setActive] = useState(0);
  const [saving, setSaving] = useState(false);
  const [loading, setLoading] = useState(true);
  const wsRef = useRef<XLSX.WorkSheet[]>([]);

  useEffect(() => {
    (async () => {
      setLoading(true);
      try {
        const wb = XLSX.read(new Uint8Array(file.buffer.slice(0)), { type: "array" });
        wsRef.current = wb.SheetNames.map((n) => wb.Sheets[n]);
        const parsed: Sheet[] = wb.SheetNames.map((name) => {
          const ws = wb.Sheets[name];
          const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" }) as (string | number)[][];
          return { name, rows: aoa.length ? aoa : [[""]] };
        });
        setSheets(parsed);
      } catch (e) {
        setSheets([{ name: "Sheet1", rows: [["解析失败：" + (e as Error).message]] }]);
      } finally {
        setLoading(false);
      }
    })();
  }, [file]);

  function setCell(si: number, ri: number, ci: number, val: string) {
    setSheets((prev) => {
      const next = prev.slice();
      const rows = next[si].rows.map((r) => r.slice());
      while (rows.length <= ri) rows.push([]);
      while (rows[ri].length <= ci) rows[ri].push("");
      const num = Number(val);
      rows[ri][ci] = val !== "" && !isNaN(num) && val.trim() !== "" ? num : val;
      next[si] = { ...next[si], rows };
      return next;
    });
  }

  async function handleSave() {
    setSaving(true);
    try {
      const wb = XLSX.utils.book_new();
      sheets.forEach((s) => {
        const ws = XLSX.utils.aoa_to_sheet(s.rows);
        XLSX.utils.book_append_sheet(wb, ws, s.name.slice(0, 31));
      });
      const out = XLSX.write(wb, { bookType: "xlsx", type: "array" }) as ArrayBuffer;
      await saveFileDialog(file.name, new Uint8Array(out));
    } finally {
      setSaving(false);
    }
  }

  const current = sheets[active];

  return (
    <div style={{ height: "100%", display: "flex", flexDirection: "column" }}>
      <div style={{ display: "flex", gap: 8, alignItems: "center", padding: "8px 12px", background: "var(--card)", borderBottom: "1px solid var(--border)" }}>
        <button className="btn primary" disabled={saving} onClick={handleSave}>
          <Save size={16} /> {saving ? "保存中…" : "保存 / 另存为 xlsx"}
        </button>
        {loading && <span style={{ fontSize: 12, color: "var(--muted-fg)" }}>解析中…</span>}
      </div>
      {sheets.length > 1 && (
        <div style={{ display: "flex", gap: 2, padding: "4px 12px", background: "var(--card)", borderBottom: "1px solid var(--border)" }}>
          {sheets.map((s, i) => (
            <button key={i} className={`btn ${i === active ? "primary" : ""}`} style={{ height: 26 }} onClick={() => setActive(i)}>
              {s.name}
            </button>
          ))}
        </div>
      )}
      <div style={{ flex: 1, overflow: "auto", padding: 16, background: "var(--muted)" }}>
        {current && (
          <table style={{ borderCollapse: "collapse", background: "var(--card)", fontSize: 13, width: "100%" }}>
            <tbody>
              {current.rows.map((row, ri) => (
                <tr key={ri}>
                  <td style={cellStyle(true)}>{ri + 1}</td>
                  {row.map((val, ci) => (
                    <td key={ci} style={cellStyle(false)}>
                      <input
                        value={val as any}
                        onChange={(e) => setCell(active, ri, ci, e.target.value)}
                        style={{ width: "100%", border: "none", background: "transparent", color: "inherit", font: "inherit", outline: "none", padding: "4px 6px" }}
                      />
                    </td>
                  ))}
                </tr>
              ))}
            </tbody>
          </table>
        )}
      </div>
    </div>
  );
}

function cellStyle(header: boolean): React.CSSProperties {
  return {
    border: "1px solid var(--border)",
    padding: header ? "2px 8px" : "0",
    background: header ? "var(--muted)" : "var(--card)",
    color: header ? "var(--muted-fg)" : "var(--fg)",
    textAlign: header ? "center" : "left",
    minWidth: 80,
    maxWidth: 320,
  };
}
