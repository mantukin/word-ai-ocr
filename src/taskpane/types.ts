export type Alignment = 'left' | 'center' | 'right' | 'justify';

export interface TextStyle {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  fontSize?: number;
  fontFamily?: string;
  color?: string;
  highlight?: string;
}

export interface ParagraphStyle extends TextStyle {
  alignment?: Alignment;
  indentFirstLine?: number; // Points
  indentLeft?: number;      // Points
  indentRight?: number;     // Points
  spacingBefore?: number;   // Points
  spacingAfter?: number;    // Points
  lineSpacing?: number;     // Multiplier (e.g., 1.5)
}

export interface TextRun {
  text: string;
  style?: TextStyle;
}

export type ParagraphContent = string | TextRun[];

export interface ParagraphBlock {
  type: 'paragraph';
  content: ParagraphContent;
  style?: ParagraphStyle;
}

export interface HeadingBlock {
  type: 'heading';
  level: 1 | 2 | 3 | 4 | 5 | 6;
  content: ParagraphContent;
  style?: ParagraphStyle;
}

export interface ListItem {
  content: ParagraphContent;
  sublist?: ListBlock; // Recursive for nested lists
}

export interface ListBlock {
  type: 'list';
  listType: 'bullet' | 'number';
  items: ListItem[];
}

export interface TableCell {
  content: DocumentElement[]; // Cells can contain paragraphs, lists, etc.
  colSpan?: number;
  rowSpan?: number;
  width?: number; // Percentage or fixed
  shadingColor?: string; // Hex color
}

export interface TableRow {
  cells: TableCell[];
  height?: number;
}

export interface TableBlock {
  type: 'table';
  rows: TableRow[];
  style?: {
    width?: number; // Percentage (0-100)
    borders?: boolean; // Default true
  };
}

export type DocumentElement = ParagraphBlock | HeadingBlock | ListBlock | TableBlock;

export interface DocumentStructure {
  elements: DocumentElement[];
}
