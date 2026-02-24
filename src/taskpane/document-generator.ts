/* global Word console */

import { DocumentStructure, HeadingBlock, ListBlock, TableBlock, ParagraphContent, ParagraphStyle, TextRun, ParagraphBlock } from "./types";

export async function generateDocument(structure: DocumentStructure) {
  await Word.run(async (context) => {
    const body = context.document.body;

    for (let i = 0; i < structure.elements.length; i++) {
      const element = structure.elements[i];
      try {
        if (element.type === "paragraph") {
          insertContentAcrossParagraphs(body, element.content, element.style);
        } else if (element.type === "heading") {
          insertContentAcrossParagraphs(body, element.content, element.style, {
            isHeading: true,
            headingLevel: (element as HeadingBlock).level
          });
        } else if (element.type === "list") {
          insertList(body, element as ListBlock);
        } else if (element.type === "table") {
          await insertTableExplicit(context, body, element as TableBlock);
        }
        await context.sync();
      } catch (elemError) {
        console.error(`Error at element ${i}:`, elemError);
      }
    }
  });
}

function insertContentAcrossParagraphs(
  parent: Word.Body,
  content: ParagraphContent,
  blockStyle: ParagraphStyle | undefined,
  options: {
    isHeading?: boolean,
    headingLevel?: number,
    listLevel?: number,
    listPrefix?: string
  } = {}
) {
  let runs: TextRun[] = [];
  if (typeof content === "string") {
    runs = [{ text: content }];
  } else if (Array.isArray(content)) {
    runs = (content as any[]).map(item => {
      if (typeof item === "string") return { text: item };
      return item as TextRun;
    });
  }

  let p: Word.Paragraph | null = null;
  let paragraphIndexAcrossRuns = 0;

  for (const run of runs) {
    const lines = (run.text || "").split("\n");
    for (let i = 0; i < lines.length; i++) {
      const lineText = lines[i];
      if (i > 0 || p === null) {
        p = parent.insertParagraph("", "End");
        if (options.isHeading) {
          p.font.bold = true;
          p.font.size = 18 - ((options.headingLevel || 1) * 2);
        }
        if (options.listLevel !== undefined) {
          p.leftIndent = (options.listLevel + 1) * 20;
          if (paragraphIndexAcrossRuns === 0 && options.listPrefix) {
            p.insertText(options.listPrefix, "Start");
            p.firstLineIndent = -15;
          }
        }
        applyParagraphStyle(p, blockStyle);
        paragraphIndexAcrossRuns++;
      }
      if (lineText.length > 0) {
        const range = p.insertText(lineText, "End");
        if (run.style) {
          if (run.style.bold !== undefined) range.font.bold = run.style.bold;
          if (run.style.italic !== undefined) range.font.italic = run.style.italic;
          if (run.style.fontSize) range.font.size = run.style.fontSize;
          if (run.style.color) range.font.color = run.style.color;
          if (run.style.fontFamily) range.font.name = run.style.fontFamily;
        }
      }
    }
  }
}

function insertList(parent: Word.Body, block: ListBlock, level: number = 0) {
  if (!block.items || block.items.length === 0) return;
  for (let i = 0; i < block.items.length; i++) {
    const item = block.items[i];
    const prefix = block.listType === "number" ? `${i + 1}. ` : "• ";
    insertContentAcrossParagraphs(parent, item.content, undefined, {
      listLevel: level,
      listPrefix: prefix
    });
    if (item.sublist) insertList(parent, item.sublist, level + 1);
  }
}

async function insertTableExplicit(context: Word.RequestContext, parent: Word.Body, block: TableBlock) {
  const rowCount = block.rows.length;
  if (rowCount === 0) return;

  let colCount = 0;
  block.rows.forEach(row => {
    let rowCols = 0;
    row.cells.forEach(cell => rowCols += (cell.colSpan || 1));
    if (rowCols > colCount) colCount = rowCols;
  });

  // Prepare data
  const data: string[][] = [];
  for (let r = 0; r < rowCount; r++) {
    const rowData: string[] = new Array(colCount).fill("");
    let currentCol = 0;
    for (const cellBlock of block.rows[r].cells) {
      let cellText = "";
      if (typeof cellBlock.content === "string") {
        cellText = cellBlock.content;
      } else if (Array.isArray(cellBlock.content)) {
        cellText = cellBlock.content.map(item => {
          if (typeof item === "string") return item;
          const anyItem = item as any;
          if (typeof anyItem.content === "string") return anyItem.content;
          if (Array.isArray(anyItem.content)) return anyItem.content.map((run: any) => run.text || "").join("");
          return "";
        }).join("\n");
      }
      if (currentCol < colCount) rowData[currentCol] = cellText;
      currentCol += (cellBlock.colSpan || 1);
    }
    data.push(rowData);
  }

  const table = parent.insertTable(rowCount, colCount, "End", data);

  // Style borders
  if (block.style && block.style.borders === false) {
    table.getBorder("Top").type = "None";
    table.getBorder("Bottom").type = "None";
    table.getBorder("Left").type = "None";
    table.getBorder("Right").type = "None";
    table.getBorder("InsideHorizontal").type = "None";
    table.getBorder("InsideVertical").type = "None";
  } else {
    table.style = "Table Grid";
  }

  // Now apply individual formatting to cells from JSON
  for (let r = 0; r < rowCount; r++) {
    let currentCol = 0;
    for (let c = 0; c < block.rows[r].cells.length; c++) {
      const cellBlock = block.rows[r].cells[c];
      const cell = table.getCell(r, currentCol);

      // Default cell font if not specified
      cell.body.font.name = "Times New Roman";
      cell.body.font.size = 10;

      // Find style to apply (from first paragraph inside cell if complex)
      let styleToApply: ParagraphStyle | undefined;
      if (Array.isArray(cellBlock.content) && cellBlock.content.length > 0) {
        const firstElem = cellBlock.content[0] as any;
        if (firstElem.style) styleToApply = firstElem.style;
      }

      if (styleToApply) {
        // Apply to all paragraphs in cell
        const paragraphs = cell.body.paragraphs;
        paragraphs.load("items");
        await context.sync();

        for (let p of paragraphs.items) {
          applyParagraphStyle(p, styleToApply);
        }
      }

      currentCol += (cellBlock.colSpan || 1);
    }
  }

  await context.sync();
}

function applyParagraphStyle(paragraph: Word.Paragraph, style: ParagraphStyle | undefined) {
  paragraph.font.bold = false;
  paragraph.font.italic = false;
  if (!style) return;
  if (style.bold !== undefined) paragraph.font.bold = style.bold;
  if (style.italic !== undefined) paragraph.font.italic = style.italic;
  if (style.fontSize) paragraph.font.size = style.fontSize;
  if (style.fontFamily) paragraph.font.name = style.fontFamily;
  if (style.alignment) {
    const align = style.alignment.toLowerCase();
    if (align === "center" || align === "centered") paragraph.alignment = "Centered";
    else if (align === "right") paragraph.alignment = "Right";
    else if (align === "justify" || align === "justified") paragraph.alignment = "Justified";
    else paragraph.alignment = "Left";
  }
  // Line spacing
  if (style.lineSpacing !== undefined) {
    const rule = style.lineRule || "Multiple";
    if (rule === "Multiple") {
      // Convert multiplier to points: Word lineSpacing is in points (12pt = single)
      paragraph.lineSpacing = style.lineSpacing * 12;
    } else {
      // "AtLeast" or "Exactly" — value is already in points
      paragraph.lineSpacing = style.lineSpacing;
    }
  }
  // Paragraph spacing
  if (style.spacingBefore !== undefined) paragraph.spaceBefore = style.spacingBefore;
  if (style.spacingAfter !== undefined) paragraph.spaceAfter = style.spacingAfter;
  // Indentation
  if (style.indentFirstLine !== undefined) paragraph.firstLineIndent = style.indentFirstLine;
  if (style.indentLeft !== undefined) paragraph.leftIndent = style.indentLeft;
  if (style.indentRight !== undefined) paragraph.rightIndent = style.indentRight;
}
