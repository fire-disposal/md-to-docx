import type { Block, ImageBinary, TableBlock, TableCell } from "./types";
import { parseInline } from "./parser";

// ============ Page Settings ============

export interface PageSettings {
  width: number;   // twips
  height: number;  // twips
  top: number;
  bottom: number;
  left: number;
  right: number;
}

export const PAGE_A4: PageSettings = {
  width: 11906,   // 210 mm
  height: 16838,  // 297 mm
  top: 1440,      // 1 inch
  bottom: 1440,
  left: 1440,
  right: 1440,
};

export const PAGE_LETTER: PageSettings = {
  width: 12240,   // 8.5 inch
  height: 15840,  // 11 inch
  top: 1440,
  bottom: 1440,
  left: 1440,
  right: 1440,
};

// ============ Document XML ============

const DOC_HEADER = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
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

const DOC_FOOTER = "</w:body>\n</w:document>";

function buildSectPr(page: PageSettings): string {
  return `<w:sectPr>
    <w:pgSz w:w="${page.width}" w:h="${page.height}"/>
    <w:pgMar w:top="${page.top}" w:bottom="${page.bottom}" w:left="${page.left}" w:right="${page.right}"/>
  </w:sectPr>`;
}

export function generateDocumentXml(
  blocks: Block[],
  page: PageSettings,
  imageMaxWidth?: number,
  imageRels?: Map<string, { rid: string; ext: string }>,
  images?: Map<string, ImageBinary>,
  addCaptions?: boolean
): string {
  const bodyXml = blocks
    .map((block) => blockToXml(block, imageRels, images, addCaptions, imageMaxWidth))
    .join("");
  return DOC_HEADER + bodyXml + buildSectPr(page) + DOC_FOOTER;
}

// ============ Block to XML ============

function blockToXml(
  block: Block,
  imageRels?: Map<string, { rid: string }>,
  images?: Map<string, ImageBinary>,
  addCaptions?: boolean,
  imageMaxWidth?: number
): string {
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
      return block.lines
        .map((line) => {
          if (!line.trim()) {
            return `<w:p><w:pPr><w:ind w:left="720" w:right="720"/></w:pPr></w:p>`;
          }
          const runs = parseInline(line);
          return `<w:p><w:pPr><w:ind w:left="720" w:right="720"/></w:pPr>${runs}</w:p>`;
        })
        .join("");
    }

    case "code": {
      const codeLines = block.code.split("\n");
      return codeLines
        .map((codeLine) => {
          const escaped = escapeXml(codeLine || " ");
          return `<w:p><w:pPr><w:rPr><w:rFonts w:ascii="Courier New" w:hAnsi="Courier New"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Courier New" w:hAnsi="Courier New"/></w:rPr><w:t xml:space="preserve">${escaped}</w:t></w:r></w:p>`;
        })
        .join("");
    }

    case "list": {
      return block.items
        .map((item, idx) => {
          let prefix: string;
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
          return `<w:p><w:pPr><w:ind w:left="360" w:hanging="360"/></w:pPr><w:r><w:t xml:space="preserve">${escapeXml(prefix)}</w:t></w:r>${runs}</w:p>`;
        })
        .join("");
    }

    case "image": {
      const img = images?.get(block.path);
      if (!img) {
        const alt = block.alt || block.path;
        return `<w:p><w:r><w:t xml:space="preserve">[Image: ${escapeXml(alt)}]</w:t></w:r></w:p>`;
      }

      const rel = imageRels?.get(block.path);
      const rid = rel?.rid ?? "rId1";
      const { width, height } = imgDimensionsToEmu(img, imageMaxWidth);
      const drawingXml = buildImageDrawing(rid, width, height);

      let result = `<w:p>${drawingXml}</w:p>`;

      if (addCaptions && block.alt) {
        result += `<w:p><w:pPr><w:pStyle w:val="Caption"/><w:jc w:val="center"/></w:pPr><w:r><w:t xml:space="preserve">${escapeXml(block.alt)}</w:t></w:r></w:p>`;
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

// ============ Table XML ============

function buildTableXml(table: TableBlock): string {
  if (table.rows.length === 0) return "";

  const colWidth = Math.floor(9000 / Math.max(table.columnCount, 1)); // ~total width / cols
  const gridCols = Array.from({ length: table.columnCount }, () =>
    `<w:gridCol w:w="${colWidth}"/>`
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

function buildTableRowXml(row: import("./types").TableRow): string {
  const cellsXml = row.cells.map((cell) => buildTableCellXml(cell)).join("");
  return `  <w:tr>${cellsXml}</w:tr>`;
}

function buildTableCellXml(cell: TableCell): string {
  const runs = parseInline(cell.text);
  const boldProp = cell.header ? `<w:b/>` : "";
  return `
    <w:tc>
      <w:p><w:r><w:rPr>${boldProp}</w:rPr>${runs}</w:r></w:p>
    </w:tc>`;
}

// ============ Image Drawing XML ============

function buildImageDrawing(relId: string, cx: number, cy: number): string {
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

// ============ Content Types XML ============

export function generateContentTypesXml(
  imageExtensions: string[]
): string {
  const defaults = imageExtensions
    .map((ext) => {
      const mime = extToMime(ext);
      return `  <Default Extension="${ext}" ContentType="${mime}"/>`;
    })
    .join("\n");

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
${defaults}
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>`;
}

// ============ Relationships XML ============

export function generateRelsXml(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;
}

export function generateDocRelsXml(
  imageExtensions: string[]
): string {
  let rels = `  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>\n`;
  let imgIdx = 2;
  for (let i = 0; i < imageExtensions.length; i++) {
    rels += `  <Relationship Id="rId${imgIdx}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image${i + 1}.${imageExtensions[i]}"/>\n`;
    imgIdx++;
  }
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
${rels}</Relationships>`;
}

// ============ Styles XML ============

export function generateStylesXml(): string {
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

// ============ Helpers ============

function escapeXml(str: string): string {
  return str
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function extToMime(ext: string): string {
  const map: Record<string, string> = {
    png: "image/png",
    jpg: "image/jpeg",
    jpeg: "image/jpeg",
    gif: "image/gif",
    bmp: "image/bmp",
    webp: "image/webp",
    svg: "image/svg+xml",
  };
  return map[ext.toLowerCase()] || "application/octet-stream";
}

function imgDimensionsToEmu(img: ImageBinary, maxWidthInch?: number): {
  width: number;
  height: number;
} {
  const EMU_PER_INCH = 914400;
  const PX_PER_INCH = 96;
  const EMU_PER_PX = Math.round(EMU_PER_INCH / PX_PER_INCH);

  let w: number;
  let h: number;

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
