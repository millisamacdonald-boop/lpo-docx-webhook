const express = require('express');
const { Document, Packer, Paragraph, TextRun, AlignmentType, LevelFormat, BorderStyle } = require('docx');

const app = express();
app.use(express.json({ limit: '1mb' }));

const WEBHOOK_SECRET = process.env.WEBHOOK_SECRET || 'changeme';

app.get('/health', (req, res) => res.json({ status: 'ok' }));

app.post('/generate-docx', async (req, res) => {
  const secret = req.headers['x-webhook-secret'];
  if (secret !== WEBHOOK_SECRET) {
    return res.status(401).json({ error: 'Unauthorized' });
  }

  const { letterText, filename } = req.body;

  if (!letterText) {
    return res.status(400).json({ error: 'letterText is required' });
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
  const sections = splitLetters(letterText);
  const children = [];

  sections.forEach((section, idx) => {
    if (idx > 0) {
      children.push(new Paragraph({
        children: [new TextRun({ text: '', break: 1 })],
        pageBreakBefore: true
      }));
    }

    const lines = section.trim().split('\n');

    lines.forEach((line, lineIdx) => {
      const trimmed = line.trim();

      if (trimmed === '') {
        children.push(new Paragraph({ children: [], spacing: { after: 80 } }));
        return;
      }

      if (lineIdx < 4 && (trimmed.startsWith('Subject:') || trimmed.startsWith('Re:') || /^(Decline|Future|Protection|Support|The\s)/i.test(trimmed))) {
        children.push(new Paragraph({
          children: [new TextRun({ text: trimmed, bold: true, size: 24, font: 'Arial' })],
          spacing: { before: 0, after: 200 }
        }));
        return;
      }

      if (trimmed.startsWith('Dear ')) {
        children.push(new Paragraph({
          children: [new TextRun({ text: trimmed, size: 22, font: 'Arial' })],
          spacing: { before: 160, after: 160 }
        }));
        return;
      }

      if (trimmed.startsWith('Yours ') || trimmed.startsWith('Kind regards') || trimmed.startsWith('Warm regards')) {
        children.push(new Paragraph({
          children: [new TextRun({ text: trimmed, size: 22, font: 'Arial' })],
          spacing: { before: 240, after: 80 }
        }));
        return;
      }

      if (/^[-*]\s/.test(trimmed)) {
        children.push(new Paragraph({
          numbering: { reference: 'bullets', level: 0 },
          children: [new TextRun({ text: trimmed.replace(/^[-*]\s/, ''), size: 22, font: 'Arial' })],
          spacing: { after: 80 }
        }));
        return;
      }

      children.push(new Paragraph({
        children: [new TextRun({ text: trimmed, size: 22, font: 'Arial' })],
        spacing: { after: 160 }
      }));
    });
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
      children
    }]
  });

  return Packer.toBuffer(doc);
}

function splitLetters(text) {
  const separatorPattern = /\n[-=]{3,}\n|\n\*{3,}\n/;
  const parts = text.split(separatorPattern).filter(p => p.trim().length > 0);
  return parts.length > 0 ? parts : [text];
}

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`docx-webhook running on port ${PORT}`));
