const express = require('express');
const multer = require('multer');
const PDFDocument = require('pdfkit');
const fs = require('fs');
const https = require('https');
const OpenAI = require('openai');
const path = require('path');
require('dotenv').config();

const app = express();
const upload = multer({ storage: multer.memoryStorage() });
const client = new OpenAI({ apiKey: process.env.OPENAI_API_KEY });

app.use(express.static(path.join(__dirname, 'public')));
app.use(express.urlencoded({ extended: true, limit: '10mb' }));

// Filename helpers for UTF-8 and safe ASCII fallback
function normalizeTurkishToAscii(input) {
  const map = {
    'ç': 'c', 'Ç': 'C', 'ğ': 'g', 'Ğ': 'G', 'ı': 'i', 'I': 'I', 'İ': 'I',
    'ö': 'o', 'Ö': 'O', 'ş': 's', 'Ş': 'S', 'ü': 'u', 'Ü': 'U'
  };
  return (input || '').split('').map((ch) => map[ch] || ch).join('');
}

function makeSafeFilenameBase(name) {
  const ascii = normalizeTurkishToAscii(String(name || '').trim());
  const replaced = ascii.replace(/\s+/g, '_');
  return replaced.replace(/[^a-zA-Z0-9._-]/g, '').slice(0, 100) || 'degerlendirme';
}

// Download and cache a Unicode-capable font (NotoSans/DejaVu) if missing
function downloadFile(url, destPath) {
  return new Promise((resolve, reject) => {
    const file = fs.createWriteStream(destPath);
    https
      .get(url, (res) => {
        if (res.statusCode !== 200) {
          file.close(() => fs.existsSync(destPath) && fs.unlinkSync(destPath));
          return reject(new Error('HTTP ' + res.statusCode));
        }
        res.pipe(file);
        file.on('finish', () => file.close(() => resolve(destPath)));
      })
      .on('error', (err) => {
        file.close(() => fs.existsSync(destPath) && fs.unlinkSync(destPath));
        reject(err);
      });
  });
}

async function ensureLocalUnicodeFontPath() {
  const fontsDir = path.join(__dirname, 'fonts');
  try { if (!fs.existsSync(fontsDir)) fs.mkdirSync(fontsDir, { recursive: true }); } catch (_) {}
  const notoPath = path.join(fontsDir, 'NotoSans-Regular.ttf');
  if (fs.existsSync(notoPath)) return notoPath;
  try {
    await downloadFile('https://github.com/googlefonts/noto-fonts/raw/main/hinted/ttf/NotoSans/NotoSans-Regular.ttf', notoPath);
    return notoPath;
  } catch (_) {}
  const dejavuPath = path.join(fontsDir, 'DejaVuSans.ttf');
  if (fs.existsSync(dejavuPath)) return dejavuPath;
  try {
    await downloadFile('https://github.com/dejavu-fonts/dejavu-fonts/raw/version_2_37/ttf/DejaVuSans.ttf', dejavuPath);
    return dejavuPath;
  } catch (_) {}
  return null;
}

// Try to resolve a Unicode-capable font from system fonts
function resolveSystemUnicodeFontPath() {
  const winDir = process.env.WINDIR || 'C\\Windows';
  const candidates = [
    path.join(__dirname, 'fonts', 'NotoSans-Regular.ttf'),
    path.join(__dirname, 'fonts', 'DejaVuSans.ttf'),
    path.join(winDir, 'Fonts', 'calibri.ttf'),
    path.join(winDir, 'Fonts', 'tahoma.ttf'),
    path.join(winDir, 'Fonts', 'verdana.ttf'),
    path.join(winDir, 'Fonts', 'arial.ttf'),
    path.join(winDir, 'Fonts', 'arialuni.ttf'),
    path.join(winDir, 'Fonts', 'segoeui.ttf')
  ];
  for (const p of candidates) {
    try { if (fs.existsSync(p)) return p; } catch (_) {}
  }
  return null;
}

app.post(['/evaluate','/api/evaluate'], upload.single('file'), async (req, res) => {
  try {
    if (!req.file) return res.status(400).send('Dosya gerekli.');

    const { examCode, criteriaJson } = req.body;
    if (!process.env.OPENAI_API_KEY) {
      return res.status(500).send('OPENAI_API_KEY .env dosyasında tanımlı olmalı.');
    }

    let criteria = [];
    let subject = '';
    if (criteriaJson) {
      try {
        const parsed = JSON.parse(criteriaJson);
        if (Array.isArray(parsed)) {
          criteria = parsed;
        } else if (parsed && Array.isArray(parsed.criteria)) {
          subject = parsed.subject || '';
          criteria = parsed.criteria;
        }
      } catch (e) {}
    }
    if (!subject) subject = 'Kimya';

    const mime = req.file.mimetype || 'image/png';
    const base64 = req.file.buffer.toString('base64');
    const dataUrl = `data:${mime};base64,${base64}`;

    const systemPrompt = `Sadece geçerli JSON üret. Bir sınav kağıdı görselinden öğrenci bilgilerini (varsa) ve cevapları ayıkla, verilen kriterlere göre objektif puanla. Her kriter için: raw_score (0..max_points), weighted_score = raw_score * weight. Sonunda final_score_100 üret (0..100, tam sayı).`;
    const userText = `Ders: ${subject}\nExam code: ${examCode || ''}\n\nKriterler (JSON):\n${JSON.stringify(criteria)}\n\nJSON Şema:\n{\n  "student": {"ogrenci_ad": "string|null", "ogrenci_no": "string|null", "sinif": "string|null"},\n  "items": [{\n    "criterion_id": "string",\n    "name": "string",\n    "max_points": number,\n    "weight": number,\n    "raw_score": number,\n    "weighted_score": number,\n    "justification": "string",\n    "flags": ["string"]\n  }],\n  "final_score_100": number,\n  "notes": "string"\n}\n\nKurallar:\n- Öğrenci bilgileri kağıtta okunabiliyorsa doldur, değilse null bırak.\n- Kriterlerde isim/desc girdi JSON'undan alınır; puanlamayı sadece öğrenci cevabına dayanarak yap.\n- Sadece JSON döndür, başka metin ekleme.`;

    const completion = await client.chat.completions.create({
      model: 'gpt-4o-mini',
      response_format: { type: 'json_object' },
      temperature: 0.2,
      messages: [
        { role: 'system', content: systemPrompt },
        { role: 'user', content: [ { type: 'text', text: userText }, { type: 'image_url', image_url: { url: dataUrl } } ] }
      ]
    });

    const raw = completion.choices?.[0]?.message?.content || '{}';
    let result = {};
    try { result = JSON.parse(raw); } catch (e) { result = {}; }

    // Build DOCX instead of PDF
    const { Document, Packer, Paragraph, TextRun, HeadingLevel } = await import('docx');
    const st = result.student || {};
    const items = Array.isArray(result.items) ? result.items : [];
    const finalScore = Math.round(result.final_score_100 ?? 0);

    const docx = new Document({
      sections: [
        {
          children: [
            new Paragraph({ text: 'Yazılı Değerlendirme Formu', heading: HeadingLevel.TITLE }),
            new Paragraph({
              children: [
                new TextRun(`Ders: ${subject}   |   Sınav Kodu: ${examCode || '-'}`),
              ],
            }),
            new Paragraph({
              children: [
                new TextRun(`Öğrenci: ${st.ogrenci_ad || '-'}`),
                new TextRun('  |  '),
                new TextRun(`No: ${st.ogrenci_no || '-'}`),
                new TextRun('  |  '),
                new TextRun(`Sınıf: ${st.sinif || '-'}`),
              ],
            }),
            new Paragraph({ text: 'Kriter Sonuçları', heading: HeadingLevel.HEADING_2 }),
            ...items.flatMap((it, idx) => [
              new Paragraph({
                children: [ new TextRun({ text: `${idx + 1}) ${it.criterion_id} - ${it.name}`, bold: true }) ],
              }),
              new Paragraph(`Maks Puan: ${it.max_points} | Ağırlık: ${it.weight} | Puan: ${it.raw_score} | Ağırlıklı: ${it.weighted_score}`),
              ...(it.justification ? [ new Paragraph(`Gerekçe: ${it.justification}`) ] : []),
              ...(it.flags && it.flags.length ? [ new Paragraph(`Notlar: ${it.flags.join(', ')}`) ] : []),
            ]),
            new Paragraph({ text: `Final Notu (100): ${finalScore}`, heading: HeadingLevel.HEADING_3 }),
            ...(result.notes ? [ new Paragraph(`Açıklama: ${result.notes}`) ] : []),
          ],
        },
      ],
    });

    const buffer = await Packer.toBuffer(docx);
    const studentName = (st.ogrenci_ad || '').toString();
    const base = makeSafeFilenameBase(studentName || 'degerlendirme');
    const filename = `${base}.docx`;
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
    res.setHeader('Cache-Control', 'no-store, no-cache, must-revalidate, proxy-revalidate');
    res.setHeader('Pragma', 'no-cache');
    res.setHeader('Expires', '0');
    console.log('Responding with DOCX (application/vnd.openxmlformats-officedocument.wordprocessingml.document)');
    res.end(buffer);
  } catch (err) {
    console.error(err);
    res.status(500).send('Sunucu hatası: ' + err.message);
  }
});

// Export app for serverless (Vercel). Only listen when run locally.
if (require.main === module) {
  const port = process.env.PORT || 3000;
  app.listen(port, () => {
    console.log('Server running on http://localhost:' + port);
  });
}

module.exports = app;
