import express from 'express';
import multer from 'multer';
import cors from 'cors';
import dotenv from 'dotenv';
import { execFile } from 'child_process';
import { promisify } from 'util';
import fs from 'fs/promises';
import path from 'path';
import os from 'os';
import Anthropic from '@anthropic-ai/sdk';
import { buildDocx } from './docxBuilder.js';

dotenv.config();
const execFileAsync = promisify(execFile);

const app = express();
app.use(cors());

const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 50 * 1024 * 1024 } });
const anthropic = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY });
const MODEL = process.env.ANTHROPIC_MODEL || 'claude-sonnet-5';

app.get('/health', (req, res) => res.json({ ok: true }));

app.post('/page-count', upload.single('file'), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: 'No file uploaded' });
    const tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), 'pdf2word-count-'));
    const pdfPath = path.join(tmpDir, 'input.pdf');
    await fs.writeFile(pdfPath, req.file.buffer);
    const { stdout } = await execFileAsync('pdfinfo', [pdfPath]);
    await fs.rm(tmpDir, { recursive: true, force: true });
    const match = stdout.match(/Pages:\s+(\d+)/);
    res.json({ pages: match ? parseInt(match[1], 10) : 0 });
  } catch (err) {
    res.status(500).json({ error: err.message || 'Could not read PDF' });
  }
});

// Renders every page of a PDF buffer to PNG bytes using poppler's pdftoppm.
// Returns an array of Buffer, one per page, in order (page 1 first).
async function renderPdfPages(pdfBuffer) {
  const tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), 'pdf2word-'));
  const pdfPath = path.join(tmpDir, 'input.pdf');
  await fs.writeFile(pdfPath, pdfBuffer);
  const prefix = path.join(tmpDir, 'page');
  await execFileAsync('pdftoppm', ['-jpeg', '-r', '150', pdfPath, prefix]);
  const files = (await fs.readdir(tmpDir))
    .filter(f => f.startsWith('page') && f.endsWith('.jpg'))
    .sort((a, b) => {
      const na = parseInt(a.match(/-(\d+)\.jpg$/)?.[1] || '0', 10);
      const nb = parseInt(b.match(/-(\d+)\.jpg$/)?.[1] || '0', 10);
      return na - nb;
    });
  const buffers = [];
  for (const f of files) buffers.push(await fs.readFile(path.join(tmpDir, f)));
  await fs.rm(tmpDir, { recursive: true, force: true });
  return buffers;
}

const PAGE_PROMPT = `You are converting a scanned or digital document page into structured content for a Word document.
Look at the page image and return ONLY valid JSON (no markdown fences, no commentary) matching this shape:

{
  "blocks": [
    { "type": "heading", "text": "..." },
    { "type": "paragraph", "text": "..." },
    { "type": "table", "rows": [["cell","cell"], ["cell","cell"]] }
  ]
}

Rules:
- Preserve the original language of the document exactly — do not translate.
- Preserve reading order top to bottom.
- Use "heading" only for real section titles, not for emphasis.
- For any word or number you are not confident about (blurry, cut off, ambiguous), wrap ONLY that fragment in double curly braces, e.g. "the total is {{4,850.00}}". Do not wrap text you are confident about.
- Never invent text that is not visibly present on the page.
- If the page is blank or has no extractable content, return {"blocks": []}.
- Do not include page numbers, headers/footers that just repeat, as separate noise blocks unless they carry real information.`;

async function parsePageWithClaude(imageBuffer) {
  const base64 = imageBuffer.toString('base64');
  const msg = await anthropic.messages.create({
    model: MODEL,
    max_tokens: 4096,
    messages: [{
      role: 'user',
      content: [
        { type: 'image', source: { type: 'base64', media_type: 'image/jpeg', data: base64 } },
        { type: 'text', text: PAGE_PROMPT }
      ]
    }]
  });
  const text = msg.content.find(b => b.type === 'text')?.text || '{"blocks":[]}';
  const cleaned = text.replace(/^```json\s*/i, '').replace(/```\s*$/i, '').trim();
  try {
    return JSON.parse(cleaned);
  } catch {
    return { blocks: [{ type: 'paragraph', text: '[Не удалось разобрать страницу]' }] };
  }
}

app.post('/convert', upload.single('file'), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: 'No file uploaded' });
    if (!process.env.ANTHROPIC_API_KEY) return res.status(500).json({ error: 'ANTHROPIC_API_KEY not configured on server' });

    const pageImages = await renderPdfPages(req.file.buffer);
    if (pageImages.length === 0) return res.status(422).json({ error: 'Could not render any pages from this PDF' });

    let indices = pageImages.map((_, i) => i);
    if (req.body.pages) {
      const wanted = String(req.body.pages).split(',').map(n => parseInt(n.trim(), 10) - 1);
      const filtered = wanted.filter(i => i >= 0 && i < pageImages.length);
      if (filtered.length) indices = filtered;
    }

    const pageResults = [];
    for (const i of indices) {
      const parsed = await parsePageWithClaude(pageImages[i]);
      pageResults.push({ pageNumber: i + 1, blocks: parsed.blocks || [] });
    }

    const docxBuffer = await buildDocx(pageResults);

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', `attachment; filename="converted.docx"`);
    res.setHeader('X-Pages-Converted', String(indices.length));
    res.send(docxBuffer);
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message || 'Conversion failed' });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`pdf-to-word backend listening on :${PORT}`));
