import type { Block, ListItem, TableCell, TableRow } from "./types";

// ============ Block-level Parser ============

export function parseBlocks(md: string): Block[] {
  const lines = md.replace(/\r\n/g, "\n").split("\n");
  const blocks: Block[] = [];
  let i = 0;

  while (i < lines.length) {
    const line = lines[i];

    // Empty line
    if (line.trim() === "") {
      blocks.push({ type: "empty" });
      i++;
      continue;
    }

    // Fenced code block
    if (/^ {0,3}```/.test(line)) {
      const lang = line.replace(/^ {0,3}```\s*/, "") || null;
      const codeLines: string[] = [];
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

    // Table — starts with |, has separator row
    if (/^\|.+\|$/.test(line)) {
      const headerCells = parseTableRow(line);

      // Look for separator row
      if (
        i + 1 < lines.length &&
        /^\|[\s\-:|]+\|$/.test(lines[i + 1])
      ) {
        i += 2; // skip header and separator
        const rows: TableRow[] = [
          { cells: headerCells.map((text) => ({ text, header: true })) },
        ];

        // Parse alignments from separator (not used for now)
        while (i < lines.length && /^\|.+\|$/.test(lines[i])) {
          const dataCells = parseTableRow(lines[i]);
          rows.push({
            cells: dataCells.map((text) => ({ text, header: false })),
          });
          i++;
        }

        const columnCount = headerCells.length;
        blocks.push({ type: "table", rows, columnCount });
        continue;
      }

      // No separator — treat as regular paragraph
      blocks.push({ type: "paragraph", text: line });
      i++;
      continue;
    }

    // Horizontal rule
    if (/^ {0,3}(-{3,}|\*{3,}|_{3,})\s*$/.test(line)) {
      blocks.push({ type: "hr" });
      i++;
      continue;
    }

    // Heading
    const headingMatch = line.match(/^ {0,3}(#{1,6})\s+(.*)/);
    if (headingMatch) {
      blocks.push({
        type: "heading",
        level: headingMatch[1].length,
        text: headingMatch[2],
      });
      i++;
      continue;
    }

    // Image (standalone)
    const imgMatch = line.match(/^ {0,3}!\[([^\]]*)\]\(([^)]+)\)\s*$/);
    if (imgMatch) {
      blocks.push({
        type: "image",
        alt: imgMatch[1] || "",
        path: imgMatch[2],
      });
      i++;
      continue;
    }

    // Wiki image (standalone)
    const wikiImgMatch = line.match(/^ {0,3}!\[\[([^\]]+)\]\]\s*$/);
    if (wikiImgMatch) {
      const parts = wikiImgMatch[1].split("|");
      const alt = parts.length > 1 ? parts[1] : parts[0];
      blocks.push({ type: "image", alt, path: parts[0] });
      i++;
      continue;
    }

    // Unordered list / task list
    const ulMatch = line.match(/^ {0,3}[\-\*\+]\s+(.*)/);
    if (ulMatch) {
      const items: ListItem[] = [];
      while (i < lines.length) {
        const m = lines[i].match(/^ {0,3}[\-\*\+]\s+(.*)/);
        if (!m) break;
        const taskMatch = m[1].match(/^\[([ xX])\]\s+(.*)/);
        if (taskMatch) {
          items.push({
            text: taskMatch[2],
            checked: taskMatch[1].toLowerCase() === "x",
          });
        } else {
          items.push({ text: m[1], checked: null });
        }
        i++;
      }
      blocks.push({ type: "list", items, ordered: false, start: 1 });
      continue;
    }

    // Ordered list
    const olMatch = line.match(/^ {0,3}(\d+)\.\s+(.*)/);
    if (olMatch) {
      const items: ListItem[] = [];
      const start = parseInt(olMatch[1], 10);
      while (i < lines.length) {
        const m = lines[i].match(/^ {0,3}(\d+)\.\s+(.*)/);
        if (!m) break;
        items.push({ text: m[2], checked: null });
        i++;
      }
      blocks.push({ type: "list", items, ordered: true, start });
      continue;
    }

    // Blockquote
    if (line.match(/^ {0,3}>/)) {
      const quotedLines: string[] = [];
      while (i < lines.length && /^ {0,3}>/.test(lines[i])) {
        quotedLines.push(lines[i].replace(/^ {0,3}> ?/, ""));
        i++;
      }
      blocks.push({ type: "blockquote", lines: quotedLines });
      continue;
    }

    // Regular paragraph
    blocks.push({ type: "paragraph", text: line });
    i++;
  }

  return blocks;
}

function parseTableRow(line: string): string[] {
  return line
    .replace(/^\|/, "")
    .replace(/\|$/, "")
    .split("|")
    .map((cell) => cell.trim());
}

// ============ Inline Parser ============

export function parseInline(text: string): string {
  if (text.length === 0) return "";

  const regex =
    /(\*\*\*(.+?)\*\*\*)|(\*\*(.+?)\*\*)|(\*(.+?)\*)|(~~(.+?)~~)|(==(.+?)==)|(`([^`]+)`)|(\[([^\]]+)\]\(([^)]+)\))|(!\[([^\]]*)\]\(([^)]+)\))|(\[\[([^\]]+?)\]\])/g;

  let result = "";
  let lastIdx = 0;
  let match: RegExpExecArray | null;

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

function escapeXml(str: string): string {
  return str
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}
