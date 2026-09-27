// 平台抽象层。
// 采用「Web 层文件选择 + Blob 下载」方案：在浏览器和 Tauri WebView 中行为一致，
// 无需配置 Tauri 插件权限作用域，保证首版安装包开箱即用。
// 后续若需原生系统文件对话框，可在 src-tauri 增加自定义命令(pick_open/pick_save)并切换此处实现。

export function isTauri(): boolean {
  return typeof window !== "undefined" && "__TAURI_INTERNALS__" in window;
}

export interface OpenedFile {
  path?: string;
  name: string;
  ext: string;
  blob: Blob;
  buffer: ArrayBuffer;
}

function extOf(name: string): string {
  const i = name.lastIndexOf(".");
  return i >= 0 ? name.slice(i + 1).toLowerCase() : "";
}

const ACCEPT = ".doc,.docx,.xls,.xlsx,.ppt,.pptx,.pdf";

export async function openFileDialog(): Promise<OpenedFile | null> {
  return new Promise((resolve) => {
    const input = document.createElement("input");
    input.type = "file";
    input.accept = ACCEPT;
    input.onchange = async () => {
      const f = input.files?.[0];
      if (!f) return resolve(null);
      const buffer = await f.arrayBuffer();
      resolve({ name: f.name, ext: extOf(f.name), blob: f, buffer });
    };
    input.click();
  });
}

export async function saveFileDialog(defaultName: string, data: Uint8Array | Blob): Promise<void> {
  const blob = data instanceof Blob ? data : new Blob([new Uint8Array(data)]);
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = defaultName;
  document.body.appendChild(a);
  a.click();
  a.remove();
  // 延迟回收，避免部分浏览器下载未触发就被 revoke
  setTimeout(() => URL.revokeObjectURL(url), 4000);
}

// 旧版 Office 二进制容器（OLE Compound File：.doc/.xls/.ppt 及 WPS 以 docx 后缀保存的旧格式）魔数。
// OOXML（docx/xlsx/pptx）本质是 zip，开头为 PK\x03\x04，不会被此处命中。
const OLE_MAGIC = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];

export function isLegacyBinaryFormat(buffer: ArrayBuffer): boolean {
  if (buffer.byteLength < 8) return false;
  const head = new Uint8Array(buffer, 0, 8);
  return OLE_MAGIC.every((b, i) => head[i] === b);
}
