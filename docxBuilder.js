import { Document, Packer, Paragraph, TextRun, HeadingLevel, Table, TableRow, TableCell, WidthType, PageBreak, ShadingType } from 'docx';

// Splits text on {{...}} markers and returns an array of TextRun,
// highlighting the marked (low-confidence) fragments.
function runsForText(text) {
  const parts = String(text).split(/(\{\{[^{}]*\}\})/g).filter(p => p.length > 0);
  return parts.map(part => {
    const m = part.match(/^\{\{([^{}]*)\}\}$/);
    if (m) {
      return new TextRun({ text: m[1], shading: { type: ShadingType.CLEAR, fill: 'FCE4C8' }, color: '8A4B0F' });
    }
    return new TextRun(part);
  });
}

function blockToParagraphs(block) {
  if (block.type === 'heading') {
    return [new Paragraph({ heading: HeadingLevel.HEADING_2, children: runsForText(block.text || '') })];
  }
  if (block.type === 'paragraph') {
    return [new Paragraph({ spacing: { after: 160 }, children: runsForText(block.text || '') })];
  }
  if (block.type === 'table' && Array.isArray(block.rows) && block.rows.length) {
    const colCount = Math.max(...block.rows.map(r => r.length));
    const rows = block.rows.map((row, ri) =>
      new TableRow({
        children: Array.from({ length: colCount }, (_, ci) =>
          new TableCell({
            width: { size: Math.floor(9000 / colCount), type: WidthType.DXA },
            shading: ri === 0 ? { type: ShadingType.CLEAR, fill: 'EEF0F6' } : undefined,
            margins: { top: 80, bottom: 80, left: 100, right: 100 },
            children: [new Paragraph({ children: runsForText(row[ci] ?? '') })]
          })
        )
      })
    );
    return [new Table({ width: { size: 9000, type: WidthType.DXA }, rows }), new Paragraph({ text: '', spacing: { after: 160 } })];
  }
  return [];
}

export async function buildDocx(pageResults) {
  const children = [];
  pageResults.forEach((page, idx) => {
    for (const block of page.blocks) {
      children.push(...blockToParagraphs(block));
    }
    if (idx < pageResults.length - 1) {
      children.push(new Paragraph({ children: [new PageBreak()] }));
    }
  });

  if (children.length === 0) {
    children.push(new Paragraph({ text: 'Не удалось извлечь содержимое из выбранных страниц.' }));
  }

  const doc = new Document({ sections: [{ children }] });
  return Packer.toBuffer(doc);
}
