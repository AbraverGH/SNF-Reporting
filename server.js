/**
 * AR Collections Explorer — backend
 * ----------------------------------
 * A minimal server that:
 *   1. Serves the dashboard (public/index.html)
 *   2. Accepts an admin-only upload of the AR Collections workbook,
 *      parses it server-side, and writes the result to data/dataset.json
 *   3. Exposes that dataset to every visitor via GET /api/dataset
 *
 * There's no external database — the parsed dataset is a single JSON file
 * on disk. That's plenty for this size of data (tens of thousands of rows,
 * a few MB of JSON) and keeps deployment simple. IMPORTANT: whatever host
 * you deploy this to must give the `data/` folder a *persistent* disk/volume,
 * or the dataset will disappear on every redeploy/restart. See README.md.
 */
const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

const app = express();
const PORT = process.env.PORT || 3000;
const ADMIN_KEY = process.env.ADMIN_KEY || 'changeme'; // set this in your host's env vars!
const DATA_DIR = path.join(__dirname, 'data');
const DATA_FILE = path.join(DATA_DIR, 'dataset.json');

if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR, { recursive: true });
if (ADMIN_KEY === 'changeme') {
  console.warn('\n⚠️  ADMIN_KEY is not set — using the default "changeme".');
  console.warn('   Set a real ADMIN_KEY environment variable before going live!\n');
}

const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 25 * 1024 * 1024 } });

app.use(express.static(path.join(__dirname, 'public')));
app.use(express.json());

/* ------------------------------------------------------------------
   Same parsing logic as the browser version (kept in sync intentionally)
------------------------------------------------------------------ */
const MONTH_IDX = {Jan:0,Feb:1,Mar:2,Apr:3,May:4,Jun:5,Jul:6,Aug:7,Sep:8,Oct:9,Nov:10,Dec:11};
const PAYER_COLUMNS = [
  {key:'allGroups', header:'All  Groups'},
  {key:'medicare', header:'Medicare'},
  {key:'neMedicaid', header:'NE Medicaid'},
  {key:'kyMedicaid', header:'Kentucky Medicaid'},
  {key:'okMedicaid', header:'OK Medicaid'},
  {key:'privatePay', header:'Private Pay'},
  {key:'hmo', header:'HMO/Private Insurance'},
  {key:'hospice', header:'Hospice'},
  {key:'veterans', header:'Veterans'},
  {key:'mltcMedicaid', header:'MLTC Medicaid'},
  {key:'respite', header:'Respite'},
  {key:'budget', header:'Budget'}
];

function normHeader(h) {
  return String(h || '').replace(/\s+/g, ' ').trim().toLowerCase();
}

function parseWorkbookBuffer(buffer) {
  const wb = XLSX.read(buffer, { type: 'buffer', cellDates: true });
  const sheetName = wb.SheetNames.includes('Sheet1') ? 'Sheet1' : wb.SheetNames[0];
  const ws = wb.Sheets[sheetName];
  const grid = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
  if (!grid.length) throw new Error('The selected sheet appears to be empty.');

  const headerRow = grid[0].map(normHeader);
  function col(name) { return headerRow.indexOf(normHeader(name)); }
  const idx = {
    company: col('Company'), metric: col('Metric'), month: col('Month'),
    year: col('Year'), asOf: col('As Of'), code: col('Code'), facility: col('Facility')
  };
  const payerIdx = PAYER_COLUMNS.map(p => col(p.header));

  if (idx.company < 0 || idx.metric < 0 || idx.facility < 0) {
    throw new Error('This file does not look like the expected AR Collections layout.');
  }

  const records = [];
  for (let r = 1; r < grid.length; r++) {
    const row = grid[r];
    if (!row || row[idx.company] == null || row[idx.company] === '') continue;
    const monthAbbr = row[idx.month];
    const year = row[idx.year];
    if (monthAbbr === 'Total' || year === 'Total') continue;
    if (!(monthAbbr in MONTH_IDX)) continue;
    const asOfVal = row[idx.asOf];
    if (!(asOfVal instanceof Date)) continue;

    const values = {};
    PAYER_COLUMNS.forEach((p, i) => {
      const ci = payerIdx[i];
      const v = ci >= 0 ? row[ci] : null;
      values[p.key] = (typeof v === 'number') ? v : null;
    });

    records.push({
      company: row[idx.company],
      metric: row[idx.metric],
      facility: row[idx.facility],
      code: row[idx.code],
      monthAbbr,
      year: Number(year),
      asOfISO: asOfVal.toISOString(),
      values
    });
  }

  if (!records.length) throw new Error('No usable rows were found in this file.');
  return records;
}

/* ------------------------------------------------------------------
   API
------------------------------------------------------------------ */

// Anyone with the link can read the current shared dataset.
app.get('/api/dataset', (req, res) => {
  if (!fs.existsSync(DATA_FILE)) return res.json(null);
  fs.createReadStream(DATA_FILE).pipe(res.type('application/json'));
});

// Only requests with the correct x-admin-key header can replace it.
app.post('/api/upload', upload.single('file'), (req, res) => {
  try {
    if (req.headers['x-admin-key'] !== ADMIN_KEY) {
      return res.status(401).json({ error: 'Invalid admin key.' });
    }
    if (!req.file) {
      return res.status(400).json({ error: 'No file was uploaded.' });
    }
    const records = parseWorkbookBuffer(req.file.buffer);
    const payload = {
      fileName: req.file.originalname,
      uploadedAt: new Date().toISOString(),
      rowCount: records.length,
      records
    };
    fs.writeFileSync(DATA_FILE, JSON.stringify(payload));
    res.json({ ok: true, rowCount: records.length, fileName: req.file.originalname });
  } catch (err) {
    res.status(400).json({ error: err.message });
  }
});

app.get('/api/health', (req, res) => res.json({ ok: true }));

app.listen(PORT, () => {
  console.log('AR Collections Explorer listening on port ' + PORT);
});
