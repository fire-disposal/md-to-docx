# Self-Review Checklist

Self-review against [Obsidian plugin review guidelines](https://docs.obsidian.md/Plugins/Releasing/Plugin+guidelines).

## Manifest

- [x] `id` matches repo name (`md-to-docx`)
- [x] `name` is descriptive (`MD to DOCX`)
- [x] `version` follows semver (`1.0.0`)
- [x] `minAppVersion` is correct (`0.15.0` — required for `createBinary`)
- [x] `description` is accurate
- [x] `author` is set (`firedisposal`)
- [x] `isDesktopOnly` is `true` (uses `electron.shell.openPath`)
- [x] No `fundingUrl` set (optional)

## Code

- [x] Plugin class extends `Plugin`, implements `onload` / `onunload`
- [x] Uses only public Obsidian API (`Notice`, `Plugin`, `TFile`, `Setting`, etc.)
- [x] No `innerHTML`, `eval`, or dynamic code execution
- [x] No access to file system outside vault (uses only `vault.read` / `vault.createBinary` / etc.)
- [x] No network access (all processing is local)
- [x] No secret/API key leakage
- [x] `styles.css` present (empty — no custom CSS injected into Obsidian UI)

## Dependencies

- [x] Zero runtime dependencies (only dev: `esbuild`, `typescript`, `obsidian`)
- [x] No `node_modules` shipped in release
- [x] No external binaries called

## Releasing

- [x] `main.js` bundled via esbuild to single CJS file
- [x] `versions.json` present for backward compatibility
- [ ] GitHub Release with `main.js`, `manifest.json`, `styles.css`
- [ ] No `.env` or secrets in release assets

## Behavior

- [x] Plugin does not modify core Obsidian behavior
- [x] Ribbon icon + commands are non-intrusive
- [x] No automatic execution — only runs on user action
- [x] File creation respects user's vault structure
- [x] Error handling via `Notice` — no crashes, no silent failures
- [x] Verified on Arch Linux and Windows 11
