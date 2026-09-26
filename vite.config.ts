import { defineConfig } from "vite";

// 相对路径，便于 Tauri 以 file:// 或自定义协议加载
// 使用 Vite 内置 esbuild 的 automatic JSX（无需 @vitejs/plugin-react 依赖）
export default defineConfig({
  plugins: [],
  esbuild: {
    jsx: "automatic",
  },
  base: "./",
  build: {
    outDir: "dist",
    chunkSizeWarningLimit: 4000,
  },
  // Tauri 期望固定的本地端口
  server: {
    port: 1420,
    strictPort: true,
  },
});
