import { useEffect, useMemo, useState, useCallback } from "react";
import { Sidebar } from "./components/Sidebar";
import { Tabs } from "./components/Tabs";
import { Toolbar } from "./components/Toolbar";
import { PdfViewer } from "./modules/pdf/PdfViewer";
import { DocxViewer } from "./modules/docx/DocxViewer";
import { XlsxGrid } from "./modules/xlsx/XlsxGrid";
import { PptxViewer } from "./modules/pptx/PptxViewer";
import { openFileDialog, type OpenedFile } from "./platform";

export interface TabItem {
  id: string;
  file: OpenedFile;
}

interface RecentItem {
  name: string;
  path?: string;
  ext: string;
  lastOpen: number;
}

const RECENTS_KEY = "wx_office_recents";

function loadRecents(): RecentItem[] {
  try {
    return JSON.parse(localStorage.getItem(RECENTS_KEY) || "[]");
  } catch {
    return [];
  }
}

function pushRecent(recents: RecentItem[], item: RecentItem): RecentItem[] {
  const next = [item, ...recents.filter((r) => r.name !== item.name || r.path !== item.path)].slice(0, 20);
  localStorage.setItem(RECENTS_KEY, JSON.stringify(next));
  return next;
}

export default function App() {
  const [theme, setTheme] = useState<"light" | "dark">(
    () => (localStorage.getItem("wx_theme") as "light" | "dark") || "light"
  );
  const [recents, setRecents] = useState<RecentItem[]>(() => loadRecents());
  const [tabs, setTabs] = useState<TabItem[]>([]);
  const [activeId, setActiveId] = useState<string | null>(null);

  useEffect(() => {
    document.documentElement.setAttribute("data-theme", theme);
    localStorage.setItem("wx_theme", theme);
  }, [theme]);

  const openFile = useCallback(async () => {
    const f = await openFileDialog();
    if (!f) return;
    const id = `${f.name}-${Date.now()}`;
    setTabs((t) => [...t, { id, file: f }]);
    setActiveId(id);
    setRecents((r) => pushRecent(r, { name: f.name, path: f.path, ext: f.ext, lastOpen: Date.now() }));
  }, []);

  const closeTab = useCallback(
    (id: string) => {
      setTabs((t) => {
        const next = t.filter((x) => x.id !== id);
        if (activeId === id) setActiveId(next[0]?.id ?? null);
        return next;
      });
    },
    [activeId]
  );

  const activeTab = useMemo(() => tabs.find((t) => t.id === activeId) || null, [tabs, activeId]);

  return (
    <div style={{ display: "flex", flexDirection: "column", height: "100%" }}>
      <Toolbar onOpen={openFile} theme={theme} onToggleTheme={() => setTheme((t) => (t === "light" ? "dark" : "light"))} />
      <div style={{ display: "flex", flex: 1, minHeight: 0 }}>
        <Sidebar recents={recents} onOpen={openFile} />
        <div style={{ display: "flex", flexDirection: "column", flex: 1, minWidth: 0 }}>
          <Tabs tabs={tabs} activeId={activeId} onSelect={setActiveId} onClose={closeTab} />
          <main style={{ flex: 1, minHeight: 0, overflow: "hidden", background: "var(--muted)" }}>
            {activeTab ? <EditorRouter tab={activeTab} key={activeTab.id} /> : <EmptyState onOpen={openFile} />}
          </main>
        </div>
      </div>
    </div>
  );
}

function EditorRouter({ tab }: { tab: TabItem }) {
  const ext = tab.file.ext;
  if (ext === "pdf") return <PdfViewer file={tab.file} />;
  if (ext === "docx" || ext === "doc") return <DocxViewer file={tab.file} />;
  if (ext === "xlsx" || ext === "xls") return <XlsxGrid file={tab.file} />;
  if (ext === "pptx" || ext === "ppt") return <PptxViewer file={tab.file} />;
  return <div style={{ padding: 24 }}>不支持的文件格式：.{ext}</div>;
}

function EmptyState({ onOpen }: { onOpen: () => void }) {
  return (
    <div style={{ height: "100%", display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", gap: 16, color: "var(--muted-fg)" }}>
      <div style={{ fontSize: 18, fontWeight: 600, color: "var(--fg)" }}>网信办公 · 轻量全格式读改</div>
      <div style={{ fontSize: 13 }}>打开一个 Word / Excel / PPT / PDF，做轻量修改与批注</div>
      <button className="btn primary" onClick={onOpen}>打开文件</button>
      <div style={{ fontSize: 12, maxWidth: 420, textAlign: "center", lineHeight: 1.8 }}>
        隐私优先：所有解析与编辑均在本地完成，无需登录、不上传云端。
      </div>
    </div>
  );
}
