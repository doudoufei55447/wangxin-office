# 网信办公 · 轻量全格式读改（WangXin Office Light）

品牌：网信科技（WX 蓝色渐变字标）。一款不到 150MB、秒级启动、无需登录的桌面端轻量工具，
能打开任意 Word / Excel / PPT / PDF 做轻量修改与批注。

- 技术栈：**Tauri 2 + React 18 + Vite + TypeScript**
- 全格式引擎（纯前端）：pdf.js + pdf-lib（PDF）、mammoth + docx（Word）、SheetJS（Excel）、jszip（PPT 降级）
- 设计：Lucide 图标、纯色蓝 #2E5BF0、浅/深双主题，对标 PDF Expert「轻快净」

## 目录结构
```
src/
  App.tsx                 主工作台（最近文件 / 标签页 / 工具栏 / 格式路由）
  platform.ts             Tauri ↔ 浏览器 文件读写抽象（自动回退）
  theme/tokens.css        设计 Token（浅/深双主题）
  components/             Sidebar / Tabs / Toolbar
  modules/pdf|docx|xlsx|pptx/   四个格式模块
src-tauri/                Rust 原生壳（打包配置：msi / nsis / dmg / app）
docs/Spec-wangxin-office-light.md   规格契约（Spec）
```

## 本地开发（浏览器验证 UI 与格式逻辑）
```bash
npm install
npm run dev          # http://localhost:1420
npm run build        # 产出 dist/（Web 层，即 Tauri 嵌入层）
```
> 浏览器模式下「打开文件」会回退为系统文件选择框，编辑逻辑与 Tauri 内一致。

## 产出安装包（各自平台）
> ⚠️ 跨平台安装包必须在对应平台本机 `tauri build` 产出：Windows 出 .msi/.exe(nsis)，macOS 出 .dmg（需 Mac 本机 + Apple 开发者证书做公证）。

### 前置（每台构建机）
1. 安装 Rust：`https://rustup.rs`
2. 安装 Tauri CLI：`npm i -D @tauri-apps/cli`
3. **放入应用图标**：在 `src-tauri/icons/` 下放 `icon.png`（以及 macOS 的 `icon.icns`、Windows 的 `icon.ico`）。
   可用 Tauri 官方 `npx @tauri-apps/cli icon <源图>` 由一张图自动生成全套。

### Windows（本机，需 WiX / NSIS）
```bash
npm run tauri build     # 产出 src-tauri/target/release/bundle/{msi,nsis}
```

### macOS（Mac 本机）
```bash
npm run tauri build     # 产出 src-tauri/target/release/bundle/dmg
# 公证：codesign + notarize（参考 Tauri 官方文档）
```

## 能力与边界（MVP 范围）
| 格式 | 能力 | 边界 |
|------|------|------|
| PDF  | 高保真查看 + 高亮/评论/自由绘制批注 + 另存带批注 PDF | 批注烘焙进图片，重叠区文字不可选（可接受）|
| Word | mammoth 高保真查看 + contenteditable 改字 + 另存 docx | 仅提取段落/标题文本重建，复杂表格/页眉不保真 |
| Excel| 多 Sheet 查看 + 单元格数据编辑 + 公式保留 + 另存 xlsx | 大数据量未做虚拟滚动（二期）|
| PPT  | 幻灯片文本查看 + 批注（会话内）+ 文本替换写回 | **降级路线**：查看+批注+受限文本替换，排版不保真 |

## 二期规划
- 引轻后端（python-pptx / LibreOffice）补齐 PPT 改字保排版
- 标签页增强、最近文件缩略图、暗色默认、Word 转 PDF、批注汇总
- iOS / 安卓原生安装包（Tauri 2 移动端或独立方案）
