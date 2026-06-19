import type { App, TAbstractFile, TFile } from "obsidian";
import type { ImageBinary } from "./docx/types";

const SUPPORTED_EXTENSIONS = new Set([
  "png",
  "jpg",
  "jpeg",
  "gif",
  "bmp",
  "webp",
]);

export async function resolveImages(
  markdown: string,
  sourceFile: TFile,
  app: App
): Promise<Map<string, ImageBinary>> {
  const map = new Map<string, ImageBinary>();
  const refs = extractImageRefs(markdown);

  for (const ref of refs) {
    const resolved = await resolveImageRef(ref, sourceFile, app);
    if (resolved) {
      map.set(ref, resolved);
    }
  }

  return map;
}

// ============ Internal ============

function extractImageRefs(md: string): string[] {
  const refs: string[] = [];
  const seen = new Set<string>();

  // ![alt](path)
  const mdImgRe = /!\[([^\]]*)\]\(([^)]+)\)/g;
  let m: RegExpExecArray | null;
  while ((m = mdImgRe.exec(md)) !== null) {
    const path = m[2];
    if (!seen.has(path)) {
      seen.add(path);
      refs.push(path);
    }
  }

  // ![[image.png]] or ![[image.png|alias]]
  const wikiImgRe = /!\[\[([^\]]+)\]\]/g;
  while ((m = wikiImgRe.exec(md)) !== null) {
    const path = m[1].split("|")[0];
    if (!seen.has(path)) {
      seen.add(path);
      refs.push(path);
    }
  }

  return refs;
}

async function resolveImageRef(
  ref: string,
  sourceFile: TFile,
  app: App
): Promise<ImageBinary | null> {
  // Try as wiki link first: resolve via Obsidian's link resolution
  const resolvedFile = app.metadataCache.getFirstLinkpathDest(
    ref,
    sourceFile.path
  );
  if (resolvedFile && isImageFile(resolvedFile)) {
    return readImageFile(resolvedFile, app);
  }

  // Try as relative path
  const parentPath = sourceFile.parent?.path ?? "";
  const fullPath = parentPath ? `${parentPath}/${ref}` : ref;
  const normalizedPath = fullPath.replace(/\/\//g, "/");

  // Direct file lookup
  const abstractFile = app.vault.getAbstractFileByPath(normalizedPath);
  if (isImageFile(abstractFile)) {
    return readImageFile(abstractFile, app);
  }

  // Try just the filename (lookup in vault root)
  const abstractFile2 = app.vault.getAbstractFileByPath(ref);
  if (isImageFile(abstractFile2)) {
    return readImageFile(abstractFile2, app);
  }

  return null;
}

function isTFile(file: TAbstractFile | null): file is TFile {
  return file !== null && "extension" in file;
}

function isImageFile(file: TAbstractFile | null): file is TFile {
  if (!isTFile(file)) return false;
  return SUPPORTED_EXTENSIONS.has(file.extension.toLowerCase());
}

async function readImageFile(
  file: TFile,
  app: App
): Promise<ImageBinary | null> {
  try {
    const buffer = await app.vault.readBinary(file);
    const ext = file.extension.toLowerCase();
    const dims = tryReadDimensions(buffer, ext);

    return {
      buffer,
      extension: ext,
      contentType: extToMime(ext),
      width: dims?.width,
      height: dims?.height,
    };
  } catch {
    return null;
  }
}

function tryReadDimensions(
  buffer: ArrayBuffer,
  ext: string
): { width: number; height: number } | null {
  try {
    const bytes = new Uint8Array(buffer);
    if (ext === "png" && bytes.length > 24) {
      const view = new DataView(buffer);
      return {
        width: view.getUint32(16),
        height: view.getUint32(20),
      };
    }
    if (ext === "gif" && bytes.length > 10) {
      const view = new DataView(buffer);
      return {
        width: view.getUint16(6, true),
        height: view.getUint16(8, true),
      };
    }
  } catch {
    // fall through
  }
  return null;
}

function extToMime(ext: string): string {
  const m: Record<string, string> = {
    png: "image/png",
    jpg: "image/jpeg",
    jpeg: "image/jpeg",
    gif: "image/gif",
    bmp: "image/bmp",
    webp: "image/webp",
  };
  return m[ext] || "application/octet-stream";
}
