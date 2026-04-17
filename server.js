const express = require('express');
const { Document, Packer, Paragraph, TextRun, AlignmentType, LevelFormat } = require('docx');

const app = express();
app.use(express.json({ limit: '2mb' }));

const WEBHOOK_SECRET = process.env.WEBHOOK_SECRET || 'changeme';

app.get('/health', (req, res) => res.json({ status: 'ok' }));

app.post('/generate-docx', async (req, res) => {
  const secret = req.headers['x-webhook-secret'];
  if (secret !== WEBHOOK_SECRET) {
    return res.status(401).json({ error: 'Unauthorized' });
  }

  const { letterText, filename } = req.body;

  if (!letterText || letterText.trim().length < 50) {
    return res.status(400).json({ error: 'letterText is required and must contain a complete letter' });
  }

  try {
    const buffer = await buildDocx(letterText);
    const safeFilename = (filename || 'LPO_Letter').replace(/[^a-zA-Z0-9_\- ]/g, '_');
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', `attachment; filename="${safeFilename}.docx"`);
    res.setHeader('Content-Length', buffer.length);
    res.send(buffer);
  } catch (err) {
    console.error('Error generating docx:', err);
    res.status(500).json({ error: 'Failed to generate document', detail: err.message });
  }
});

function buildDocx(letterText) {
  const letterSections = letterText.split(/\n---\n/).filter(s => s.trim().length > 0);
  const allChildren = [];

  letterSections.forEach((section, idx) => {
    if (idx > 0) {
      allChildren.push(new Paragraph({ pageBreakBefore: true, children: [] }));
    }
    allChildren.push(...parseLetter(section.trim()));
  });

  const doc = new Document({
    numbering: {
      config: [{
        reference: 'bullets',
        levels: [{
          level: 0,
          format: LevelFormat.BULLET,
          text: '\u2022',
          alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } }
        }]
      }]
    },
    styles: {
      default: { document: { run: { font: 'Arial', size: 22 } } }
    },
    sections: [{
      properties: {
        page: {
          size: { width: 11906, height: 16838 },
          margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 }
        }
      },
      children: allChildren
    }]
  });

  return Packer.toBuffer(doc);
}

function parseLetter(text) {
  const lines = text.split('\n');
  const paragraphs = [];
  let lineIndex = 0;
  let letterheadDone = false;
  let letterheadCount = 0;
  let recipientDone = false;
  let inBody = false;

  while (lineIndex < lines.length) {
    const line = lines[lineIndex].trim();
    lineIndex++;

    if (line === '') {
      if (inBody) paragraphs.push(new Paragraph({ children: [], spacing: { after: 120 } }));
      continue;
    }

    const isDate = /^(Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday),?\s+\d{1,2}\s+\w+\s+\d{4}$/i.test(line) || /^\d{1,2}\s+\w+\s+\d{4}$/i.test(line);
    const isSubject = /^Subject:/i.test(line);
    const isSalutation = /^Dear\s/i.test(line);
    const isSignOff = /^Yours\s+(sincerely|faithfully)/i.test(line) || /^Kind regards/i.test(line);
    const isBullet = /^[-*]\s/.test(line);

    // Letterhead (first lines before date)
    if (!letterheadDone && !isDate && letterheadCount < 7) {
      const isFirst = letterheadCount === 0;
      paragraphs.push(new Paragraph({
        children: [new TextRun({ text: line, bold: isFirst, size: isFirst ? 24 : 22, font: 'Arial' })],
        spacing: { after: isFirst ? 80 : 40 }
      }));
      letterheadCount++;
      continue;
    }

    if (isDate) {
      letterheadDone = true;
      paragraphs.push(new Paragraph({
        children: [new TextRun({ text: line, size: 22, font: 'Arial' })],
        spacing: { before: 240, after: 240 }
      }));
      continue;
    }

    // Recipient block (between date and subject)
    if (letterheadDone && !isSubject && !isSalutation && !inBody) {
      paragraphs.push(new Paragraph({
        children: [new TextRun({ text: line, size: 22, font: 'Arial' })],
        spacing: { after: 40 }
      }));
      recipientDone = true;
      continue;
    }

    if (isSubject) {
      paragraphs.push(new Paragraph({
        children: [new TextRun({ text: line, bold: true, size: 22, font: 'Arial' })],
        spacing: { before: 200, after: 200 }
      }));
      continue;
    }

    if (isSalutation) {
      inBody = true;
      paragraphs.push(new Paragraph({
        children: [new TextRun({ text: line, size: 22, font: 'Arial' })],
        spacing: { before: 160, after: 200 }
      }));
      continue;
    }

    if (isSignOff) {
      paragraphs.push(new Paragraph({
        children: [new TextRun({ text: line, size: 22, font: 'Arial' })],
        spacing: { before: 320, after: 80 }
      }));
      while (lineIndex < lines.length) {
        const signLine = lines[lineIndex].trim();
        lineIndex++;
        if (signLine === '') continue;
        paragraphs.push(new Paragraph({
          children: [new TextRun({ text: signLine, size: 22, font: 'Arial' })],
          spacing: { after: 40 }
        }));
      }
      continue;
    }

    if (isBullet) {
      paragraphs.push(new Paragraph({
        numbering: { reference: 'bullets', level: 0 },
        children: [new TextRun({ text: line.replace(/^[-*]\s/, ''), size: 22, font: 'Arial' })],
        spacing: { after: 80 }
      }));
      continue;
    }

    // Normal body paragraph
    paragraphs.push(new Paragraph({
      children: [new TextRun({ text: line, size: 22, font: 'Arial' })],
      spacing: { after: 160 }
    }));
  }

  return paragraphs;
}

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`docx-webhook running on port ${PORT}`));
