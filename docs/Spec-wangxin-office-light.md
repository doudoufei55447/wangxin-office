# Spec - 网信办公（轻量版）v1.0.0

> 生成日期：2026-09-26
> 基于：PRD v1 + 架构文档 v1 + UIUX 文档 v1
> 状态：已确认
> 品牌：网信科技（WX 蓝色渐变字标 #2E4FE8→#45C8F0；界面主色纯色蓝 #2E5BF0）

---

## 1. 产品定义
- **一句话描述**：一款不到 150MB、秒级启动、无需登录的桌面端轻量工具，能打开任意 Word/Excel/PPT/PDF 做轻量修改与批注。
- **目标用户**：偶尔改文件的职场人/学生、低配老旧设备用户、注重隐私不愿上云的用户。
- **核心问题**：用户只想改几个字/改个数字/加个批注，却被迫装几 GB 重型 Office 或用满是广告的 WPS。

## 2. MVP 范围（锁定）
| 优先级 | 功能 | 验收标准摘要 |
|--------|------|-------------|
| P0 | 全格式本地高保真打开：docx/doc、xlsx/xls、pptx/ppt、pdf | 本地解析，格式尽量还原，不依赖云 |
| P0 | PDF 批注：高亮/下划线/评论/自由绘制 | 批注可保存进 PDF 文件 |
| P0 | Word 轻量文字编辑（改字、字体、基础段落） | 内容可改并另存为 docx |
| P0 | Excel 单元格数据编辑、基础公式、单元格格式 | 单元格改动可保存回 xlsx |
| P0 | 本地保存 / 另存为（保留原格式） | 输出文件可被原软件正常打开 |
| P0 | 极小安装包（目标 <150MB）+ 秒级启动 | 无账号、离线优先、无广告 |
| P1 | 标签页多文档、最近文件、暗色模式、快速打开 | 单窗口多标签 |
| P1 | Word/PPT 转 PDF 导出 | 导出可见 |
| P2 | 进阶函数库、批注汇总、轻量模板、可选云同步、OCR | 二期 |

## 3. 明确不做（Out-of-Scope）
| 不做的功能 | 原因 | 何时考虑 |
|------------|------|----------|
| 复杂排版/样式深度编辑 | MVP 定位轻量，避免臃肿 | 二期专业版 |
| 多人实时协作 | 脱离"轻净"定位，需后端 | 二期 |
| 修订追踪/批注协作流 | 同上 | 二期 |
| 宏/VBA、邮件合并、数据库 | 重型 Office 能力，非刚需 | 不计划 |
| 本期 AI 生成式写作 | 用户未要求，保持轻 | 视需求 |
| PPT 改字保排版完整编辑 | 纯前端 OOXML 不可靠，降级处理 | 二期引轻后端 |
| iOS/安卓原生安装包 | 本期仅 Win+macOS | 二期 |

## 4. 技术架构（锁定 — 版本锚定）
| 层 | 技术 | 版本 | 锁定原因 |
|----|------|------|----------|
| 框架 | Tauri 2 | 2.x | Rust 内核，整包数 MB 级，冷启 <1s，满足 <150MB+秒开 |
| 前端框架 | React | 18.x | 生态成熟，与 Tauri 集成好 |
| 构建 | Vite | 5.x | 快，Tauri 官方推荐 |
| 语言 | TypeScript | 5.x | 类型安全 |
| 状态/UI | Zustand + 原生 CSS 变量 | - | 轻量，无重型 UI 库（保持小体积）|
| 图标 | Lucide React | 最新 | SVG 描边，16/20/24px，统一一套 |
| 本地存储 | localStorage（recents）+ Tauri FS | - | MVP 轻量，免后端 |
| PDF 渲染 | pdfjs-dist | 4.x | Mozilla 官方，高保真 |
| PDF 写回 | pdf-lib | 1.x | 批注/文本写入 |
| Word 解析 | mammoth | 最新 | docx→HTML 高保真查看 |
| Word 回写 | docx | 最新 | 另存为 docx |
| Excel | xlsx (SheetJS) | 最新 | 解析/写回 xlsx |
| PPT 解析 | jszip + pptx 文本提取 | 最新 | 降级查看+批注 |
| 打包 | tauri-build (WiX msi / nsis / dmg) | - | 双平台安装包 |

## 5. 命令/接口清单（Tauri invoke + 前端模块）
| 类型 | 标识 | 功能 | 说明 |
|------|------|------|------|
| Tauri cmd | open_file_dialog | 打开系统文件选择 | dialog plugin |
| Tauri cmd | read_file(path) | 读文件为字节 | fs plugin |
| Tauri cmd | write_file(path, bytes) | 写文件 | fs plugin |
| 前端路由 | routeByExt(ext) | 按扩展名路由到对应 Viewer | 纯前端 |
| 前端模块 | pdf/annotate | PDF 批注 | 纯前端 |
| 前端模块 | docx/edit | Word 轻改 | 纯前端 |
| 前端模块 | xlsx/edit | Excel 单元格编辑 | 纯前端 |
| 前端模块 | pptx/view | PPT 降级查看+批注 | 纯前端 |

## 6. 数据库表清单
无服务端数据库。本地 recents 用 localStorage（key: `wx_office_recents`），结构：
```
{ path, name, ext, lastOpen, thumb? }
```

## 7. 页面清单
| 页面/视图 | 路由 | 核心组件 | 设计 Token 主题 |
|------|------|----------|-----------------|
| 主工作台 | / | Sidebar(Recents) + Tabs + Toolbar + EditorArea | 浅/深双主题 |
| PDF 视图 | tab | PdfViewer + AnnotationToolbar | - |
| Word 视图 | tab | DocxViewer(contenteditable) | - |
| Excel 视图 | tab | XlsxGrid | - |
| PPT 视图 | tab | PptxViewer + AnnotationOverlay | - |

## 8. 设计 Token（锁定）
- 主色 Primary: #2E5BF0（品牌蓝青渐变仅留 LOGO）
- 浅色：Background #F7F9FC / Card #FFFFFF / Foreground #1A2233 / Muted #EEF2F8 / Border #E2E8F1
- 深色：Background #0E1320 / Card #161D2E / Foreground #E8EDF5 / Border #283146
- 语义色：Destructive #E5484D / Success #16A34A / Warning #E8A33D
- 字体：拉丁 Plus Jakarta Sans；中文 系统原生（PingFang SC/微软雅黑/Noto Sans SC）；等宽 JetBrains Mono
- 图标：Lucide，16/20/24px，stroke 1.75 round
- 对标：PDF Expert「轻快净」；单窗口标签页；最近文件即首页

## 9. 验收标准（EARS）
| 编号 | 功能 | 验收标准 |
|------|------|----------|
| AC-01 | 打开 | When 用户选择本地文件，系统 必须 在 2s 内渲染首屏 |
| AC-02 | PDF 批注 | When 用户添加高亮/评论，系统 必须 将其写入并能在重开后显示 |
| AC-03 | Word 编辑 | If 用户修改文字并保存，系统 必须 输出可被 Word 打开的 docx |
| AC-04 | Excel 编辑 | When 用户改单元格并保存，系统 必须 保留数据到 xlsx |
| AC-05 | 体积 | The system 应该 安装包 <150MB |
| AC-06 | 隐私 | The system 必须 不强制登录、不联网上传文档 |

## 10. 边界与约束
- 不支持 IE；最低 Win10 / macOS 11
- PPT 为降级能力：查看+批注+受限文本替换，明说保真上限
- 性能目标：冷启 <1s，打开 10MB 文档 <2s

## 11. 内嵌已知坑
| 坑 | 技术栈 | 根因 | 修法 |
|----|--------|------|------|
| pdf.js worker 路径 | pdfjs-dist | 需显式配置 workerSrc | 用 vite ?worker 导入 |
| mammoth 样式丢失 | mammoth | 仅转结构 | 自定义 styleMap 映射 |
| xlsx 大文件卡顿 | SheetJS | 全量解析 | 分片/虚拟滚动（二期）|

## 12. 端到端验证步骤
```bash
npm install
npm run tauri dev        # 开发模式（需 Rust）
# 或纯前端验证：
npm run dev              # Vite 浏览器验证 UI 与格式逻辑
npm run build            # vite build 验证编译
# 出安装包（各自平台）：
npm run tauri build      # Win 出 msi/nsis；Mac 出 dmg（需 Mac 本机）
```

## 13. 变更记录
| 日期 | 变更 | 原因 | 影响 |
|------|------|------|------|
| 2026-09-26 | 初版 | 需求确认 | 全范围 |
