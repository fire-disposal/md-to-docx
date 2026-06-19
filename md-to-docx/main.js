"use strict";
var __defProp = Object.defineProperty;
var __getOwnPropDesc = Object.getOwnPropertyDescriptor;
var __getOwnPropNames = Object.getOwnPropertyNames;
var __hasOwnProp = Object.prototype.hasOwnProperty;
var __export = (target, all) => {
  for (var name in all)
    __defProp(target, name, { get: all[name], enumerable: true });
};
var __copyProps = (to, from, except, desc) => {
  if (from && typeof from === "object" || typeof from === "function") {
    for (let key of __getOwnPropNames(from))
      if (!__hasOwnProp.call(to, key) && key !== except)
        __defProp(to, key, { get: () => from[key], enumerable: !(desc = __getOwnPropDesc(from, key)) || desc.enumerable });
  }
  return to;
};
var __toCommonJS = (mod) => __copyProps(__defProp({}, "__esModule", { value: true }), mod);

// main.ts
var main_exports = {};
__export(main_exports, {
  default: () => MdToDocxPlugin
});
module.exports = __toCommonJS(main_exports);
var import_obsidian2 = require("obsidian");

// src/docx/parser.ts
function parseBlocks(md) {
  const lines = md.replace(/\r\n/g, "\n").split("\n");
  const blocks = [];
  let i = 0;
  while (i < lines.length) {
    const line = lines[i];
    if (line.trim() === "") {
      blocks.push({ type: "empty" });
      i++;
      continue;
    }
    if (/^ {0,3}```/.test(line)) {
      const lang = line.replace(/^ {0,3}```\s*/, "") || null;
      const codeLines = [];
      i++;
      while (i < lines.length) {
        if (/^ {0,3}```/.test(lines[i])) {
          break;
        }
        codeLines.push(lines[i]);
        i++;
      }
      blocks.push({ type: "code", lang, code: codeLines.join("\n") });
      i++;
      continue;
    }
    if (/^\|.+\|$/.test(line)) {
      const headerCells = parseTableRow(line);
      if (i + 1 < lines.length && /^\|[\s\-:|]+\|$/.test(lines[i + 1])) {
        i += 2;
        const rows = [
          { cells: headerCells.map((text) => ({ text, header: true })) }
        ];
        while (i < lines.length && /^\|.+\|$/.test(lines[i])) {
          const dataCells = parseTableRow(lines[i]);
          rows.push({
            cells: dataCells.map((text) => ({ text, header: false }))
          });
          i++;
        }
        const columnCount = headerCells.length;
        blocks.push({ type: "table", rows, columnCount });
        continue;
      }
      blocks.push({ type: "paragraph", text: line });
      i++;
      continue;
    }
    if (/^ {0,3}(-{3,}|\*{3,}|_{3,})\s*$/.test(line)) {
      blocks.push({ type: "hr" });
      i++;
      continue;
    }
    const headingMatch = line.match(/^ {0,3}(#{1,6})\s+(.*)/);
    if (headingMatch) {
      blocks.push({
        type: "heading",
        level: headingMatch[1].length,
        text: headingMatch[2]
      });
      i++;
      continue;
    }
    const imgMatch = line.match(/^ {0,3}!\[([^\]]*)\]\(([^)]+)\)\s*$/);
    if (imgMatch) {
      blocks.push({
        type: "image",
        alt: imgMatch[1] || "",
        path: imgMatch[2]
      });
      i++;
      continue;
    }
    const wikiImgMatch = line.match(/^ {0,3}!\[\[([^\]]+)\]\]\s*$/);
    if (wikiImgMatch) {
      const parts = wikiImgMatch[1].split("|");
      const alt = parts.length > 1 ? parts[1] : parts[0];
      blocks.push({ type: "image", alt, path: parts[0] });
      i++;
      continue;
    }
    const ulMatch = line.match(/^ {0,3}[\-\*\+]\s+(.*)/);
    if (ulMatch) {
      const items = [];
      while (i < lines.length) {
        const m = lines[i].match(/^ {0,3}[\-\*\+]\s+(.*)/);
        if (!m)
          break;
        const taskMatch = m[1].match(/^\[([ xX])\]\s+(.*)/);
        if (taskMatch) {
          items.push({
            text: taskMatch[2],
            checked: taskMatch[1].toLowerCase() === "x"
          });
        } else {
          items.push({ text: m[1], checked: null });
        }
        i++;
      }
      blocks.push({ type: "list", items, ordered: false, start: 1 });
      continue;
    }
    const olMatch = line.match(/^ {0,3}(\d+)\.\s+(.*)/);
    if (olMatch) {
      const items = [];
      const start = parseInt(olMatch[1], 10);
      while (i < lines.length) {
        const m = lines[i].match(/^ {0,3}(\d+)\.\s+(.*)/);
        if (!m)
          break;
        items.push({ text: m[2], checked: null });
        i++;
      }
      blocks.push({ type: "list", items, ordered: true, start });
      continue;
    }
    if (line.match(/^ {0,3}>/)) {
      const quotedLines = [];
      while (i < lines.length && /^ {0,3}>/.test(lines[i])) {
        quotedLines.push(lines[i].replace(/^ {0,3}> ?/, ""));
        i++;
      }
      blocks.push({ type: "blockquote", lines: quotedLines });
      continue;
    }
    blocks.push({ type: "paragraph", text: line });
    i++;
  }
  return blocks;
}
function parseTableRow(line) {
  return line.replace(/^\|/, "").replace(/\|$/, "").split("|").map((cell) => cell.trim());
}
function parseInline(text) {
  if (text.length === 0)
    return "";
  const regex = /(\*\*\*(.+?)\*\*\*)|(\*\*(.+?)\*\*)|(\*(.+?)\*)|(~~(.+?)~~)|(==(.+?)==)|(`([^`]+)`)|(\[([^\]]+)\]\(([^)]+)\))|(!\[([^\]]*)\]\(([^)]+)\))|(\[\[([^\]]+?)\]\])/g;
  let result = "";
  let lastIdx = 0;
  let match;
  while ((match = regex.exec(text)) !== null) {
    const before = text.slice(lastIdx, match.index);
    if (before) {
      result += `<w:r><w:t xml:space="preserve">${escapeXml(before)}</w:t></w:r>`;
    }
    if (match[1]) {
      result += `<w:r><w:rPr><w:b/><w:i/></w:rPr><w:t xml:space="preserve">${escapeXml(match[2])}</w:t></w:r>`;
    } else if (match[3]) {
      result += `<w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">${escapeXml(match[4])}</w:t></w:r>`;
    } else if (match[5]) {
      result += `<w:r><w:rPr><w:i/></w:rPr><w:t xml:space="preserve">${escapeXml(match[6])}</w:t></w:r>`;
    } else if (match[7]) {
      result += `<w:r><w:rPr><w:strike w:val="true"/></w:rPr><w:t xml:space="preserve">${escapeXml(match[8])}</w:t></w:r>`;
    } else if (match[9]) {
      result += `<w:r><w:rPr><w:highlight w:val="yellow"/></w:rPr><w:t xml:space="preserve">${escapeXml(match[10])}</w:t></w:r>`;
    } else if (match[11]) {
      result += `<w:r><w:rPr><w:rFonts w:ascii="Courier New" w:hAnsi="Courier New"/></w:rPr><w:t xml:space="preserve">${escapeXml(match[12])}</w:t></w:r>`;
    } else if (match[13]) {
      result += `<w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr><w:t xml:space="preserve">${escapeXml(match[14])}</w:t></w:r>`;
    } else if (match[16]) {
      const alt = match[17] || match[18];
      result += `<w:r><w:t xml:space="preserve">[Image: ${escapeXml(alt)}]</w:t></w:r>`;
    } else if (match[19]) {
      const parts = match[20].split("|");
      const display = parts.length > 1 ? parts[1] : parts[0];
      result += `<w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr><w:t xml:space="preserve">${escapeXml(display)}</w:t></w:r>`;
    }
    lastIdx = match.index + match[0].length;
  }
  const after = text.slice(lastIdx);
  if (after) {
    result += `<w:r><w:t xml:space="preserve">${escapeXml(after)}</w:t></w:r>`;
  }
  if (result.length === 0) {
    result = `<w:r><w:t xml:space="preserve"></w:t></w:r>`;
  }
  return result;
}
function escapeXml(str) {
  return str.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

// src/docx/xml.ts
var PAGE_A4 = {
  width: 11906,
  // 210 mm
  height: 16838,
  // 297 mm
  top: 1440,
  // 1 inch
  bottom: 1440,
  left: 1440,
  right: 1440
};
var PAGE_LETTER = {
  width: 12240,
  // 8.5 inch
  height: 15840,
  // 11 inch
  top: 1440,
  bottom: 1440,
  left: 1440,
  right: 1440
};
var DOC_HEADER = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"
            xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
            xmlns:o="urn:schemas-microsoft-com:office:office"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
            xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"
            xmlns:v="urn:schemas-microsoft-com:vml"
            xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
            xmlns:w10="urn:schemas-microsoft-com:office:word"
            xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml"
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
            xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"
            mc:Ignorable="wpc wne w10 wp wp14 wps wpg w15 w16"
            w:conformance="transitional">
  <w:body>`;
var DOC_FOOTER = "</w:body>\n</w:document>";
function buildSectPr(page) {
  return `<w:sectPr>
    <w:pgSz w:w="${page.width}" w:h="${page.height}"/>
    <w:pgMar w:top="${page.top}" w:bottom="${page.bottom}" w:left="${page.left}" w:right="${page.right}"/>
  </w:sectPr>`;
}
function generateDocumentXml(blocks, page, imageMaxWidth, imageRels, images, addCaptions) {
  const bodyXml = blocks.map((block) => blockToXml(block, imageRels, images, addCaptions, imageMaxWidth)).join("");
  return DOC_HEADER + bodyXml + buildSectPr(page) + DOC_FOOTER;
}
function blockToXml(block, imageRels, images, addCaptions, imageMaxWidth) {
  switch (block.type) {
    case "empty":
      return "<w:p></w:p>";
    case "heading": {
      const runs = parseInline(block.text);
      return `<w:p><w:pPr><w:pStyle w:val="Heading${block.level}"/></w:pPr>${runs}</w:p>`;
    }
    case "paragraph": {
      const runs = parseInline(block.text);
      return `<w:p>${runs}</w:p>`;
    }
    case "hr":
      return `<w:p><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="auto"/></w:pBdr></w:pPr></w:p>`;
    case "blockquote": {
      return block.lines.map((line) => {
        if (!line.trim()) {
          return `<w:p><w:pPr><w:ind w:left="720" w:right="720"/></w:pPr></w:p>`;
        }
        const runs = parseInline(line);
        return `<w:p><w:pPr><w:ind w:left="720" w:right="720"/></w:pPr>${runs}</w:p>`;
      }).join("");
    }
    case "code": {
      const codeLines = block.code.split("\n");
      return codeLines.map((codeLine) => {
        const escaped = escapeXml2(codeLine || " ");
        return `<w:p><w:pPr><w:rPr><w:rFonts w:ascii="Courier New" w:hAnsi="Courier New"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Courier New" w:hAnsi="Courier New"/></w:rPr><w:t xml:space="preserve">${escaped}</w:t></w:r></w:p>`;
      }).join("");
    }
    case "list": {
      return block.items.map((item, idx) => {
        let prefix;
        if (item.checked === true) {
          prefix = "\u2611 ";
        } else if (item.checked === false) {
          prefix = "\u2610 ";
        } else if (block.ordered) {
          prefix = `${block.start + idx}. `;
        } else {
          prefix = "\u2022 ";
        }
        const runs = parseInline(item.text);
        return `<w:p><w:pPr><w:ind w:left="360" w:hanging="360"/></w:pPr><w:r><w:t xml:space="preserve">${escapeXml2(prefix)}</w:t></w:r>${runs}</w:p>`;
      }).join("");
    }
    case "image": {
      const img = images?.get(block.path);
      if (!img) {
        const alt = block.alt || block.path;
        return `<w:p><w:r><w:t xml:space="preserve">[Image: ${escapeXml2(alt)}]</w:t></w:r></w:p>`;
      }
      const rel = imageRels?.get(block.path);
      const rid = rel?.rid ?? "rId1";
      const { width, height } = imgDimensionsToEmu(img, imageMaxWidth);
      const drawingXml = buildImageDrawing(rid, width, height);
      let result = `<w:p>${drawingXml}</w:p>`;
      if (addCaptions && block.alt) {
        result += `<w:p><w:pPr><w:pStyle w:val="Caption"/><w:jc w:val="center"/></w:pPr><w:r><w:t xml:space="preserve">${escapeXml2(block.alt)}</w:t></w:r></w:p>`;
      }
      return result;
    }
    case "table": {
      return buildTableXml(block);
    }
    default:
      return "<w:p></w:p>";
  }
}
function buildTableXml(table) {
  if (table.rows.length === 0)
    return "";
  const colWidth = Math.floor(9e3 / Math.max(table.columnCount, 1));
  const gridCols = Array.from(
    { length: table.columnCount },
    () => `<w:gridCol w:w="${colWidth}"/>`
  ).join("");
  const rowsXml = table.rows.map((row) => buildTableRowXml(row)).join("");
  return `<w:tbl>
  <w:tblPr>
    <w:tblW w:w="0" w:type="auto"/>
    <w:tblBorders>
      <w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>
      <w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>
      <w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>
      <w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>
      <w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/>
      <w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/>
    </w:tblBorders>
  </w:tblPr>
  <w:tblGrid>${gridCols}</w:tblGrid>
${rowsXml}
</w:tbl>`;
}
function buildTableRowXml(row) {
  const cellsXml = row.cells.map((cell) => buildTableCellXml(cell)).join("");
  return `  <w:tr>${cellsXml}</w:tr>`;
}
function buildTableCellXml(cell) {
  const runs = parseInline(cell.text);
  const boldProp = cell.header ? `<w:b/>` : "";
  return `
    <w:tc>
      <w:p><w:r><w:rPr>${boldProp}</w:rPr>${runs}</w:r></w:p>
    </w:tc>`;
}
function buildImageDrawing(relId, cx, cy) {
  return `<w:drawing>
  <wp:inline distT="0" distB="0" distL="0" distR="0">
    <wp:extent cx="${cx}" cy="${cy}"/>
    <wp:docPr id="1" name="Picture 1"/>
    <wp:cNvGraphicFramePr>
      <a:graphicFrameLocks noChangeAspect="1"/>
    </wp:cNvGraphicFramePr>
    <a:graphic>
      <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
        <pic:pic>
          <pic:nvPicPr>
            <pic:cNvPr id="0" name="image"/>
            <pic:cNvPicPr/>
          </pic:nvPicPr>
          <pic:blipFill>
            <a:blip r:embed="${relId}"/>
            <a:stretch><a:fillRect/></a:stretch>
          </pic:blipFill>
          <pic:spPr>
            <a:xfrm>
              <a:off x="0" y="0"/>
              <a:ext cx="${cx}" cy="${cy}"/>
            </a:xfrm>
          </pic:spPr>
        </pic:pic>
      </a:graphicData>
    </a:graphic>
  </wp:inline>
</w:drawing>`;
}
function generateContentTypesXml(imageExtensions) {
  const defaults = imageExtensions.map((ext) => {
    const mime = extToMime(ext);
    return `  <Default Extension="${ext}" ContentType="${mime}"/>`;
  }).join("\n");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
${defaults}
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>`;
}
function generateRelsXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;
}
function generateDocRelsXml(imageExtensions) {
  let rels = `  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
`;
  let imgIdx = 2;
  for (let i = 0; i < imageExtensions.length; i++) {
    rels += `  <Relationship Id="rId${imgIdx}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image${i + 1}.${imageExtensions[i]}"/>
`;
    imgIdx++;
  }
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
${rels}</Relationships>`;
}
function generateStylesXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
          xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
          mc:Ignorable="w14 w15">
  <w:docDefaults>
    <w:rPrDefault/>
    <w:pPrDefault/>
  </w:docDefaults>
  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="Heading 1"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:qFormat/>
    <w:rPr>
      <w:b/>
      <w:bCs/>
      <w:sz w:val="32"/>
      <w:szCs w:val="32"/>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading2">
    <w:name w:val="Heading 2"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:qFormat/>
    <w:rPr>
      <w:b/>
      <w:bCs/>
      <w:sz w:val="26"/>
      <w:szCs w:val="26"/>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading3">
    <w:name w:val="Heading 3"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:qFormat/>
    <w:rPr>
      <w:b/>
      <w:bCs/>
      <w:sz w:val="24"/>
      <w:szCs w:val="24"/>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading4">
    <w:name w:val="Heading 4"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:qFormat/>
    <w:rPr>
      <w:b/>
      <w:bCs/>
      <w:i/>
      <w:iCs/>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading5">
    <w:name w:val="Heading 5"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:qFormat/>
    <w:rPr>
      <w:b/>
      <w:bCs/>
      <w:i/>
      <w:iCs/>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading6">
    <w:name w:val="Heading 6"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:qFormat/>
    <w:rPr>
      <w:b/>
      <w:bCs/>
      <w:i/>
      <w:iCs/>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Caption">
    <w:name w:val="caption"/>
    <w:basedOn w:val="Normal"/>
    <w:rPr>
      <w:sz w:val="18"/>
      <w:szCs w:val="18"/>
    </w:rPr>
  </w:style>
  <w:style w:type="character" w:styleId="Hyperlink">
    <w:name w:val="Hyperlink"/>
    <w:rPr>
      <w:color w:val="0563C1"/>
      <w:u w:val="single"/>
    </w:rPr>
  </w:style>
</w:styles>`;
}
function escapeXml2(str) {
  return str.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}
function extToMime(ext) {
  const map = {
    png: "image/png",
    jpg: "image/jpeg",
    jpeg: "image/jpeg",
    gif: "image/gif",
    bmp: "image/bmp",
    webp: "image/webp",
    svg: "image/svg+xml"
  };
  return map[ext.toLowerCase()] || "application/octet-stream";
}
function imgDimensionsToEmu(img, maxWidthInch) {
  const EMU_PER_INCH = 914400;
  const PX_PER_INCH = 96;
  const EMU_PER_PX = Math.round(EMU_PER_INCH / PX_PER_INCH);
  let w;
  let h;
  if (img.width && img.height) {
    w = img.width * EMU_PER_PX;
    h = img.height * EMU_PER_PX;
  } else {
    w = 6 * EMU_PER_INCH;
    h = 4 * EMU_PER_INCH;
  }
  if (maxWidthInch && maxWidthInch > 0) {
    const maxEMU = maxWidthInch * EMU_PER_INCH;
    if (w > maxEMU) {
      const ratio = maxEMU / w;
      w = maxEMU;
      h = Math.round(h * ratio);
    }
  }
  return { width: w, height: h };
}

// src/docx/zip.ts
var import_zlib = require("zlib");
function crc32(data) {
  let crc = 4294967295;
  for (let i = 0; i < data.length; i++) {
    crc ^= data[i];
    for (let j = 0; j < 8; j++) {
      if (crc & 1) {
        crc = crc >>> 1 ^ 3988292384;
      } else {
        crc = crc >>> 1;
      }
    }
  }
  return (crc ^ 4294967295) >>> 0;
}
function zipEntriesToBuffer(entries) {
  const centralDir = [];
  const parts = [];
  let centralOffset = 0;
  for (const entry of entries) {
    const nameBuffer = Buffer.from(entry.name, "utf-8");
    const compressed = (0, import_zlib.deflateRawSync)(entry.data);
    const crcVal = crc32(entry.data);
    const compSize = compressed.length;
    const uncompSize = entry.data.length;
    const localHeader = Buffer.alloc(30);
    localHeader.writeUInt32LE(67324752, 0);
    localHeader.writeUInt16LE(20, 4);
    localHeader.writeUInt16LE(2048, 6);
    localHeader.writeUInt16LE(8, 8);
    localHeader.writeUInt16LE(0, 10);
    localHeader.writeUInt16LE(0, 12);
    localHeader.writeUInt32LE(crcVal, 14);
    localHeader.writeUInt32LE(compSize, 18);
    localHeader.writeUInt32LE(uncompSize, 22);
    const nameLen = nameBuffer.length;
    localHeader.writeUInt16LE(nameLen, 26);
    localHeader.writeUInt16LE(0, 28);
    parts.push(localHeader, nameBuffer, compressed);
    const cdHeader = Buffer.alloc(46);
    cdHeader.writeUInt32LE(33639248, 0);
    cdHeader.writeUInt16LE(20, 4);
    cdHeader.writeUInt16LE(20, 6);
    cdHeader.writeUInt16LE(2048, 8);
    cdHeader.writeUInt16LE(8, 10);
    cdHeader.writeUInt16LE(0, 12);
    cdHeader.writeUInt16LE(0, 14);
    cdHeader.writeUInt32LE(crcVal, 16);
    cdHeader.writeUInt32LE(compSize, 20);
    cdHeader.writeUInt32LE(uncompSize, 24);
    cdHeader.writeUInt16LE(nameLen, 28);
    cdHeader.writeUInt16LE(0, 30);
    cdHeader.writeUInt16LE(0, 32);
    cdHeader.writeUInt16LE(0, 34);
    cdHeader.writeUInt16LE(0, 36);
    cdHeader.writeUInt32LE(0, 38);
    cdHeader.writeUInt32LE(centralOffset, 42);
    centralDir.push(cdHeader, nameBuffer);
    centralOffset += localHeader.length + nameLen + compSize;
  }
  const centralDirBuffer = Buffer.concat(centralDir);
  const centralOffsetStart = centralOffset;
  const eocd = Buffer.alloc(22);
  eocd.writeUInt32LE(101010256, 0);
  eocd.writeUInt16LE(0, 4);
  eocd.writeUInt16LE(0, 6);
  eocd.writeUInt16LE(entries.length, 8);
  eocd.writeUInt16LE(entries.length, 10);
  eocd.writeUInt32LE(centralDirBuffer.length, 12);
  eocd.writeUInt32LE(centralOffsetStart, 16);
  eocd.writeUInt16LE(0, 20);
  return Buffer.concat([...parts, centralDirBuffer, eocd]);
}

// src/docx/index.ts
function markdownToDocx(markdown, options = {}) {
  const blocks = parseBlocks(markdown);
  const page = options.page ?? PAGE_A4;
  const entries = buildZipEntries(blocks, { ...options, page });
  const zipBuffer = zipEntriesToBuffer(entries);
  return zipBuffer.buffer.slice(
    zipBuffer.byteOffset,
    zipBuffer.byteOffset + zipBuffer.byteLength
  );
}
function buildZipEntries(blocks, options) {
  const entries = [];
  const { images, addCaptions, page, imageMaxWidth } = options;
  const imageRels = /* @__PURE__ */ new Map();
  let imgIdx = 2;
  if (images) {
    for (const block of blocks) {
      if (block.type === "image") {
        const img = images.get(block.path);
        if (img) {
          imageRels.set(block.path, {
            rid: `rId${imgIdx}`,
            ext: img.extension,
            buf: img.buffer
          });
          imgIdx++;
        }
      }
    }
  }
  const imageExts = Array.from(imageRels.values()).map((v) => v.ext);
  entries.push({
    name: "[Content_Types].xml",
    data: Buffer.from(generateContentTypesXml(imageExts), "utf-8")
  });
  entries.push({
    name: "_rels/.rels",
    data: Buffer.from(generateRelsXml(), "utf-8")
  });
  entries.push({
    name: "word/_rels/document.xml.rels",
    data: Buffer.from(generateDocRelsXml(imageExts), "utf-8")
  });
  entries.push({
    name: "word/document.xml",
    data: Buffer.from(
      generateDocumentXml(blocks, page, imageMaxWidth, imageRels, images, addCaptions),
      "utf-8"
    )
  });
  entries.push({
    name: "word/styles.xml",
    data: Buffer.from(generateStylesXml(), "utf-8")
  });
  let imgNum = 1;
  for (const [, v] of imageRels) {
    entries.push({
      name: `word/media/image${imgNum}.${v.ext}`,
      data: Buffer.from(v.buf)
    });
    imgNum++;
  }
  return entries;
}

// src/images.ts
var SUPPORTED_EXTENSIONS = /* @__PURE__ */ new Set([
  "png",
  "jpg",
  "jpeg",
  "gif",
  "bmp",
  "webp"
]);
async function resolveImages(markdown, sourceFile, app) {
  const map = /* @__PURE__ */ new Map();
  const refs = extractImageRefs(markdown);
  for (const ref of refs) {
    const resolved = await resolveImageRef(ref, sourceFile, app);
    if (resolved) {
      map.set(ref, resolved);
    }
  }
  return map;
}
function extractImageRefs(md) {
  const refs = [];
  const seen = /* @__PURE__ */ new Set();
  const mdImgRe = /!\[([^\]]*)\]\(([^)]+)\)/g;
  let m;
  while ((m = mdImgRe.exec(md)) !== null) {
    const path = m[2];
    if (!seen.has(path)) {
      seen.add(path);
      refs.push(path);
    }
  }
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
async function resolveImageRef(ref, sourceFile, app) {
  const resolvedFile = app.metadataCache.getFirstLinkpathDest(
    ref,
    sourceFile.path
  );
  if (resolvedFile && isImageFile(resolvedFile)) {
    return readImageFile(resolvedFile, app);
  }
  const parentPath = sourceFile.parent?.path ?? "";
  const fullPath = parentPath ? `${parentPath}/${ref}` : ref;
  const normalizedPath = fullPath.replace(/\/\//g, "/");
  const abstractFile = app.vault.getAbstractFileByPath(normalizedPath);
  if (isImageFile(abstractFile)) {
    return readImageFile(abstractFile, app);
  }
  const abstractFile2 = app.vault.getAbstractFileByPath(ref);
  if (isImageFile(abstractFile2)) {
    return readImageFile(abstractFile2, app);
  }
  return null;
}
function isTFile(file) {
  return file !== null && "extension" in file;
}
function isImageFile(file) {
  if (!isTFile(file))
    return false;
  return SUPPORTED_EXTENSIONS.has(file.extension.toLowerCase());
}
async function readImageFile(file, app) {
  try {
    const buffer = await app.vault.readBinary(file);
    const ext = file.extension.toLowerCase();
    const dims = tryReadDimensions(buffer, ext);
    return {
      buffer,
      extension: ext,
      contentType: extToMime2(ext),
      width: dims?.width,
      height: dims?.height
    };
  } catch {
    return null;
  }
}
function tryReadDimensions(buffer, ext) {
  try {
    const bytes = new Uint8Array(buffer);
    if (ext === "png" && bytes.length > 24) {
      const view = new DataView(buffer);
      return {
        width: view.getUint32(16),
        height: view.getUint32(20)
      };
    }
    if (ext === "gif" && bytes.length > 10) {
      const view = new DataView(buffer);
      return {
        width: view.getUint16(6, true),
        height: view.getUint16(8, true)
      };
    }
  } catch {
  }
  return null;
}
function extToMime2(ext) {
  const m = {
    png: "image/png",
    jpg: "image/jpeg",
    jpeg: "image/jpeg",
    gif: "image/gif",
    bmp: "image/bmp",
    webp: "image/webp"
  };
  return m[ext] || "application/octet-stream";
}

// src/settings.ts
var import_obsidian = require("obsidian");
var DEFAULT_SETTINGS = {
  overwriteExisting: true,
  embedImages: true,
  addCaptions: true,
  pageSize: "A4",
  outputFolder: "",
  openAfterExport: true,
  imageMaxWidth: 6
};
var MdToDocxSettingTab = class extends import_obsidian.PluginSettingTab {
  constructor(app, plugin) {
    super(app, plugin);
    this.plugin = plugin;
  }
  display() {
    const { containerEl } = this;
    containerEl.empty();
    containerEl.createEl("h2", { text: "Markdown to DOCX" });
    containerEl.createEl("h3", { text: "Export" });
    new import_obsidian.Setting(containerEl).setName("Output folder").setDesc(
      "Save .docx files to a subfolder (e.g. 'exports'). Leave empty to save next to the note."
    ).addText(
      (text) => text.setPlaceholder("e.g. exports").setValue(this.plugin.settings.outputFolder).onChange(async (value) => {
        this.plugin.settings.outputFolder = value.replace(/^\/|\/$/g, "");
        await this.plugin.saveSettings();
      })
    );
    new import_obsidian.Setting(containerEl).setName("Overwrite existing files").setDesc("Silently overwrite an existing .docx with the same name.").addToggle(
      (toggle) => toggle.setValue(this.plugin.settings.overwriteExisting).onChange(async (value) => {
        this.plugin.settings.overwriteExisting = value;
        await this.plugin.saveSettings();
      })
    );
    new import_obsidian.Setting(containerEl).setName("Open after export").setDesc("Open the .docx file with the system default application after export.").addToggle(
      (toggle) => toggle.setValue(this.plugin.settings.openAfterExport).onChange(async (value) => {
        this.plugin.settings.openAfterExport = value;
        await this.plugin.saveSettings();
      })
    );
    new import_obsidian.Setting(containerEl).setName("Page size").setDesc("Document page dimensions.").addDropdown(
      (dropdown) => dropdown.addOption("A4", "A4").addOption("Letter", "Letter").setValue(this.plugin.settings.pageSize).onChange(async (value) => {
        this.plugin.settings.pageSize = value;
        await this.plugin.saveSettings();
      })
    );
    containerEl.createEl("h3", { text: "Images" });
    new import_obsidian.Setting(containerEl).setName("Embed images").setDesc(
      "Embed referenced local images into the .docx. PNG, JPEG, GIF, BMP, WebP are supported."
    ).addToggle(
      (toggle) => toggle.setValue(this.plugin.settings.embedImages).onChange(async (value) => {
        this.plugin.settings.embedImages = value;
        await this.plugin.saveSettings();
      })
    );
    new import_obsidian.Setting(containerEl).setName("Add image captions").setDesc(
      "Insert the image alt-text as a caption below each embedded image."
    ).addToggle(
      (toggle) => toggle.setValue(this.plugin.settings.addCaptions).onChange(async (value) => {
        this.plugin.settings.addCaptions = value;
        await this.plugin.saveSettings();
      })
    );
    new import_obsidian.Setting(containerEl).setName("Max image width (inches)").setDesc(
      "Scale images down if wider than this. Set to 0 to disable. Default: 6 inches."
    ).addSlider(
      (slider) => slider.setLimits(0, 12, 0.5).setValue(this.plugin.settings.imageMaxWidth).setDynamicTooltip().onChange(async (value) => {
        this.plugin.settings.imageMaxWidth = value;
        await this.plugin.saveSettings();
      })
    );
  }
};

// main.ts
var MdToDocxPlugin = class extends import_obsidian2.Plugin {
  constructor() {
    super(...arguments);
    this.settings = DEFAULT_SETTINGS;
  }
  async onload() {
    await this.loadSettings();
    this.addRibbonIcon("file-down", "Export to DOCX", () => {
      this.exportCurrentNote();
    });
    this.addCommand({
      id: "export-current-note-to-docx",
      name: "Export current note to .docx",
      callback: () => this.exportCurrentNote()
    });
    this.addCommand({
      id: "export-current-note-to-docx-overwrite",
      name: "Export current note to .docx (overwrite)",
      callback: () => this.exportCurrentNote(true)
    });
    this.addSettingTab(new MdToDocxSettingTab(this.app, this));
  }
  onunload() {
  }
  async loadSettings() {
    this.settings = Object.assign({}, DEFAULT_SETTINGS, await this.loadData());
  }
  async saveSettings() {
    await this.saveData(this.settings);
  }
  async exportCurrentNote(forceOverwrite = false) {
    const file = this.app.workspace.getActiveFile();
    if (!file) {
      new import_obsidian2.Notice("No active note");
      return;
    }
    if (file.extension !== "md") {
      new import_obsidian2.Notice("The active file is not a Markdown note.");
      return;
    }
    const outPath = this.getOutputPath(file);
    const folderPath = outPath.substring(0, outPath.lastIndexOf("/"));
    if (folderPath && !await this.app.vault.adapter.exists(folderPath)) {
      try {
        await this.app.vault.createFolder(folderPath);
      } catch {
        new import_obsidian2.Notice(`Failed to create folder: ${folderPath}`);
        return;
      }
    }
    if (!forceOverwrite && !this.settings.overwriteExisting) {
      if (await this.app.vault.adapter.exists(outPath)) {
        new import_obsidian2.Notice(`File already exists: ${outPath}`);
        return;
      }
    }
    try {
      const markdown = await this.app.vault.read(file);
      let images;
      if (this.settings.embedImages) {
        images = await resolveImages(markdown, file, this.app);
      }
      const page = this.settings.pageSize === "Letter" ? PAGE_LETTER : PAGE_A4;
      const docxBuffer = markdownToDocx(markdown, {
        images,
        addCaptions: this.settings.addCaptions && this.settings.embedImages,
        page,
        imageMaxWidth: this.settings.imageMaxWidth
      });
      if (forceOverwrite || this.settings.overwriteExisting) {
        if (await this.app.vault.adapter.exists(outPath)) {
          const existing = this.app.vault.getAbstractFileByPath(outPath);
          if (existing instanceof import_obsidian2.TFile) {
            await this.app.vault.modifyBinary(existing, docxBuffer);
            new import_obsidian2.Notice(`Updated ${outPath}`);
            if (this.settings.openAfterExport) {
              this.openWithDefaultApp(outPath);
            }
            return;
          }
        }
      }
      await this.app.vault.createBinary(outPath, docxBuffer);
      new import_obsidian2.Notice(`Exported to ${outPath}`);
      if (this.settings.openAfterExport) {
        this.openWithDefaultApp(outPath);
      }
    } catch (e) {
      new import_obsidian2.Notice(`Export failed \u2014 ${e}`);
    }
  }
  getOutputPath(file) {
    const base = file.path.replace(/\.md$/, ".docx");
    const name = base.substring(base.lastIndexOf("/") + 1);
    if (this.settings.outputFolder) {
      const folder = this.settings.outputFolder;
      return `${folder}/${name}`;
    }
    const parent = file.parent?.path;
    if (parent) {
      return `${parent}/${name}`;
    }
    return name;
  }
  openWithDefaultApp(vaultPath) {
    try {
      const adapter = this.app.vault.adapter;
      const fullPath = adapter.getFullPath ? adapter.getFullPath(vaultPath) : `${adapter.basePath}/${vaultPath}`;
      const { shell } = require("electron");
      shell.openPath(fullPath);
    } catch {
    }
  }
};
