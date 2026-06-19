import { parseBlocks } from "./parser";
import {
  generateDocumentXml,
  generateContentTypesXml,
  generateRelsXml,
  generateDocRelsXml,
  generateStylesXml,
  PAGE_A4,
  type PageSettings,
} from "./xml";
import { zipEntriesToBuffer } from "./zip";
import type { ImageBinary, ZipEntry, Block } from "./types";

export interface DocxOptions {
  images?: Map<string, ImageBinary>;
  addCaptions?: boolean;
  page?: PageSettings;
  imageMaxWidth?: number;
}

export function markdownToDocx(
  markdown: string,
  options: DocxOptions = {}
): ArrayBuffer {
  const blocks = parseBlocks(markdown);
  const page = options.page ?? PAGE_A4;
  const entries = buildZipEntries(blocks, { ...options, page });
  const zipBuffer = zipEntriesToBuffer(entries);
  return zipBuffer.buffer.slice(
    zipBuffer.byteOffset,
    zipBuffer.byteOffset + zipBuffer.byteLength
  ) as ArrayBuffer;
}

interface BuildOptions extends DocxOptions {
  page: PageSettings;
}

function buildZipEntries(blocks: Block[], options: BuildOptions): ZipEntry[] {
  const entries: ZipEntry[] = [];
  const { images, addCaptions, page, imageMaxWidth } = options;

  // Map image paths to rId (rId2+ since rId1 = styles)
  const imageRels = new Map<string, { rid: string; ext: string; buf: ArrayBuffer }>();
  let imgIdx = 2;
  if (images) {
    for (const block of blocks) {
      if (block.type === "image") {
        const img = images.get(block.path);
        if (img) {
          imageRels.set(block.path, {
            rid: `rId${imgIdx}`,
            ext: img.extension,
            buf: img.buffer,
          });
          imgIdx++;
        }
      }
    }
  }

  const imageExts = Array.from(imageRels.values()).map((v) => v.ext);

  // [Content_Types].xml — MUST be first entry per OPC spec
  entries.push({
    name: "[Content_Types].xml",
    data: Buffer.from(generateContentTypesXml(imageExts), "utf-8"),
  });

  // _rels/.rels
  entries.push({
    name: "_rels/.rels",
    data: Buffer.from(generateRelsXml(), "utf-8"),
  });

  // word/_rels/document.xml.rels — includes styles + images
  entries.push({
    name: "word/_rels/document.xml.rels",
    data: Buffer.from(generateDocRelsXml(imageExts), "utf-8"),
  });

  // word/document.xml
  entries.push({
    name: "word/document.xml",
    data: Buffer.from(
      generateDocumentXml(blocks, page, imageMaxWidth, imageRels, images, addCaptions),
      "utf-8"
    ),
  });

  // word/styles.xml
  entries.push({
    name: "word/styles.xml",
    data: Buffer.from(generateStylesXml(), "utf-8"),
  });

  // Image binaries
  let imgNum = 1;
  for (const [, v] of imageRels) {
    entries.push({
      name: `word/media/image${imgNum}.${v.ext}`,
      data: Buffer.from(v.buf),
    });
    imgNum++;
  }

  return entries;
}

export type { Block, ImageBinary, ZipEntry } from "./types";
export { PAGE_A4, PAGE_LETTER, type PageSettings } from "./xml";
