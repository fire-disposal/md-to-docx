// ============ Parsed Document Types ============

export interface ListItem {
  text: string;
  checked: boolean | null;
}

export interface TableCell {
  text: string;
  header: boolean;
}

export interface TableRow {
  cells: TableCell[];
}

export interface TableBlock {
  type: "table";
  rows: TableRow[];
  columnCount: number;
}

export type Block =
  | { type: "heading"; level: number; text: string }
  | { type: "paragraph"; text: string }
  | { type: "list"; items: ListItem[]; ordered: boolean; start: number }
  | { type: "blockquote"; lines: string[] }
  | { type: "code"; lang: string | null; code: string }
  | { type: "image"; alt: string; path: string; binary?: ImageBinary }
  | { type: "table"; rows: TableRow[]; columnCount: number }
  | { type: "hr" }
  | { type: "empty" };

// ============ Image Types ============

export interface ImageBinary {
  buffer: ArrayBuffer;
  extension: string;
  contentType: string;
  width?: number;
  height?: number;
}

// ============ ZIP Types ============

export interface ZipEntry {
  name: string;
  data: Buffer;
}
