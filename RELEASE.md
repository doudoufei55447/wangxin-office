# 发布安装包（Step 1：产出首个 Win / macOS 安装包）

> 网信办公轻量版 MVP 已具备完整发布链路。安装包因跨平台打包约束，**必须在对应平台本机或 CI 产出**。
> 本沙箱缺 Rust / MSVC / NSIS / WiX / WebView2，无法在此直接出包——以下两条路径任选其一。

## 路径 A：您本机一键出包（推荐先验证）
前置（Windows）：安装 [Rust](https://rustup.rs) + Visual Studio Build Tools（勾选「C++ 桌面开发」含 MSVC）+ [NSIS](https://nsis.sourceforge.io/)（msi 用 WiX 可省，nsis 出 .exe 更轻）。
前置（macOS）：安装 Rust + Xcode Command Line Tools。

```bash
cd wangxin-office
npm install
npx tauri icon src-tauri/icons/icon.png     # 由占位 icon.png 生成全套图标
npm run tauri build                          # Win → src-tauri/target/release/bundle/{msi,nsis}
                                            # Mac → src-tauri/target/release/bundle/{dmg,app}
```

产物：
- Windows：`bundle/nsis/*.exe`（轻量安装包）、`bundle/msi/*.msi`
- macOS：`bundle/dmg/*.dmg`、`bundle/macos/*.app`

> macOS 对外分发需 Apple 开发者证书做 `codesign` + `notarize`（公证），否则用户首次打开需右键「打开」绕过 Gatekeeper。

## 路径 B：GitHub Actions 自动出包（无需本机工具链）
已内置 `.github/workflows/release.yml`，推送 tag 或手动触发即分别在 Windows / macOS runner 上构建并上传产物。

```bash
git tag v1.0.0 && git push --tags        # 或到 GitHub → Actions → Release Installers → Run workflow
# 构建完成后到 Actions 页下载 artifacts：windows-installers / macos-installers
```

## 本期已知边界（已写入 Spec）
- PPT 为降级能力（查看 + 批注 + 文本替换写回），排版不保真；二期引轻后端补齐。
- Word 保存为按段落/标题重建 docx，复杂表格/页眉不保真。
- 文件选择当前用 Web 层对话框（浏览器与 Tauri WebView 一致，开箱即用）；后续可换原生系统对话框（加自定义 Tauri 命令 pick_open/pick_save）。

## 下一步
- 出包实测通过后，推进 PPT 完整保真（python-pptx / LibreOffice 轻后端）与移动端（Tauri2 同源 iOS/安卓）。
