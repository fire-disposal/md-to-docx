# MD to DOCX

> Zero-dependency Markdown → DOCX exporter for Obsidian.  
> No Pandoc. No JSZip. No docx library. Just 31 KB of handcrafted ZIP + XML.

**Verified on:** Arch Linux, Windows 11 · Obsidian ≥ 0.15.0

[中文说明](#中文说明) | [Installation](#installation) | [Features](#features) | [Settings](#settings)

## Why This Exists

Most Markdown → DOCX solutions require Pandoc, LibreOffice, or heavy npm packages (`jszip` + `docx`). This plugin does everything from scratch:

- **Handcrafted ZIP packer** with DEFLATE compression (RFC 1951)
- **CRC-32** checksums computed inline
- **WordprocessingML XML** generated directly — no DOM builder
- **Zero runtime dependencies** — uses only Node.js built-in `zlib`

The entire plugin is **31 KB** (built). It runs instantly, nowhere to break.

## Installation

### From Community Plugins (once approved)
Search "MD to DOCX" in Obsidian → Community Plugins → Install.

### Manual
```bash
git clone https://github.com/fire-disposal/md-to-docx.git
cd md-to-docx
npm install && npm run build
# Copy md-to-docx/ folder to .obsidian/plugins/
```

### BRAT (beta testing)
Add `fire-disposal/md-to-docx` in BRAT plugin settings.

## Features

### Markdown → DOCX mapping

| Markdown | DOCX |
|----------|------|
| `#` ~ `######` | `Heading1` ~ `Heading6` |
| `**bold**` / `*italic*` / `***both***` | `<w:b/>` / `<w:i/>` / both |
| `~~strikethrough~~` | `<w:strike/>` |
| `==highlight==` | `<w:highlight/>` |
| `` `code` `` / ` ```fenced``` ` | Courier New monospace |
| `- list` / `1. ordered` / `- [ ] task` | Bullet / Numbered / ☐☑ |
| `> blockquote` | Indented paragraph |
| `---` horizontal rule | Bottom border |
| `[link](url)` / `[[wikilink]]` | Hyperlink style |
| `![image](path)` / `![[image]]` | Embedded image |
| `| table |` | DOCX table with borders |
| Caption | `Caption` style below image |

### Images
- Embed local PNG/JPEG/GIF/BMP/WebP images into the DOCX
- Auto-detects PNG/GIF dimensions; scales down to configurable max width
- Wiki links (`![[image.png]]`) resolved via Obsidian's metadata cache

### UX
- Ribbon icon + Command palette
- **English / 简体中文** — auto-detected via Obsidian language setting
- Auto-open exported file with system default app
- Configurable output subfolder
- Silent overwrite (optional)

## Settings

| Setting | Default | Description |
|---------|---------|-------------|
| Output folder | *(empty)* | Save to a subfolder (e.g. `exports`) |
| Overwrite existing | ON | Silently overwrite |
| Open after export | ON | Open with system default app |
| Page size | A4 | A4 or Letter |
| Embed images | ON | Embed local images |
| Add image captions | ON | Alt-text as caption |
| Max image width | 6 in | Scale down if wider |

## Development

```bash
npm install
npm run dev      # Watch mode
npm run build    # Type-check + production build
npm run dist     # Build + deploy to vault
```

Set `OBSIDIAN_VAULT` environment variable for `npm run dist` auto-deployment.

## License

MIT

---

## 中文说明

零依赖的 Obsidian Markdown → DOCX 导出插件。不依赖 Pandoc、JSZip、docx 等第三方库，仅使用 Node.js 内置 `zlib`，纯手写 ZIP 打包 + XML 生成。

**已验证平台：** Arch Linux、Windows 11 · Obsidian ≥ 0.15.0

### 与竞品差异

| | obsidian-docx-exporter | 本插件 |
|---|---|---|
| 运行时依赖 | JSZip + docx 库 | **零** |
| 体积 | 较大 | **31 KB** |
| 表格 | ? | ✅ |
| 题注 | ❌ | ✅ |
| 页面尺寸 | ? | ✅ A4/Letter |
| 导出后打开 | ❌ | ✅ |
| 图片宽度约束 | ? | ✅ |
| 许可 | LGPL-2.1 | MIT |
