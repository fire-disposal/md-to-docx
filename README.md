# md-to-docx

在 Obsidian 里把笔记导成 .docx，给不想为一次导出装 Pandoc 或 LibreOffice 的人；插件可用，还没进社区插件市场。

**状态**：可用（v1.0.0，桌面端，Obsidian ≥ 0.15.0）；**入口**：命令面板「Export current note to .docx」/ 侧栏图标；**仓库**：[fire-disposal/md-to-docx](https://github.com/fire-disposal/md-to-docx)

## 本质

在 Obsidian 里通过 XML 快速拼出无格式 DOCX 的临时 MD→DOCX 插件。

## 选型

基于 TypeScript + esbuild 的 Obsidian 桌面插件，零运行时依赖，压缩用 Node zlib，产物单个 CJS main.js。

- 运行时
  - Obsidian API
    - vault
    - metadataCache
    - electron.shell
- 语言与构建
  - TypeScript 5
  - esbuild
- DOCX
  - ZIP 容器
  - Node zlib
  - WordprocessingML
- 质量
  - GitHub Actions

## 功能

- 块级：H1–H6、段落、无序/有序/任务列表、引用、围栏代码、表格、分隔线、图片
- 行内：粗体、斜体、粗斜体、删除线、`==高亮==`、行内代码、链接文字、wikilink 显示文字
- 图片：嵌入 vault 内 png / jpg / jpeg / gif / bmp / webp；PNG 与 GIF 读头取真实尺寸，超宽按比例缩小；可选把 alt 写成 Caption 题注；找不到文件时输出 `[Image: alt]` 占位
- 页面：A4 或 Letter、1 英寸四边距；可配置输出子目录、静默覆盖、导出后自动打开
- 命令：普通导出、强制覆盖导出；中英界面跟随 Obsidian 语言设置
- 简化版：列表是「前缀 + 悬挂缩进」，缩进 4 空格以上的子项会掉成普通段落；链接只留显示文字，URL 不写进 relationship；代码逐行 Courier New，无底纹与语法高亮；表格按列数均分，表头不跨页重复
- 未完成：没有单测，`REVIEW.md` 只是人工自检清单；WebP 只按扩展名与内容类型打包，Word 的显示支持并不一致

## 启动流程

```bash
npm install
npm run dev      # esbuild watch，产物写到 md-to-docx/main.js
npm run build    # tsc --noEmit + esbuild production
npm run dist     # build 后按 OBSIDIAN_VAULT 复制进 vault
```

手动安装：把 `md-to-docx/`（`main.js`、`manifest.json`、`styles.css`）放进 `<vault>/.obsidian/plugins/md-to-docx/`，在设置的第三方插件里启用。

## 边界与已知问题

- 桌面端限定：`isDesktopOnly: true`，打开文件依赖 `electron.shell.openPath`。
- 未上市场：README 里还写着 “once approved”，现在只能 BRAT 或手动装。
- 构建产物实际 38281 字节（约 38 KB），旧 README 与 manifest 里的 31 KB 已过期。
- 不支持脚注、数学公式、callout、Mermaid、`![[笔记]]` 这类嵌入、PDF、页眉页脚、目录。
- 表格单元格内再套行内代码或多层强调，样式会叠成一个 run，效果不一定对。

## 文档索引

- [REVIEW.md](REVIEW.md) — 社区插件审核自检
