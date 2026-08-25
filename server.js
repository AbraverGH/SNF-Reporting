/**
 * AR Collections Explorer — backend
 * ----------------------------------
 * A minimal server that:
 *   1. Serves the dashboard (public/index.html)
 *   2. Accepts admin-only uploads of FOUR separate workbooks — Part A,
 *      Part A Coinsurance, Part B, and Part B Coinsurance — each parsed
 *      server-side and stored independently, since each is maintained
 *      and updated on its own schedule.
 *   3. Exposes all four to every visitor via a single GET /api/dataset,
 *      which the frontend merges into one combined view.
 *
 * There's no external database — each source's parsed dataset is its own
 * JSON file on disk (data/<sourceKey>.json). That's plenty for this size of
 * data and keeps deployment simple. IMPORTANT: whatever host you deploy this
 * to must give the `data/` folder a *persistent* disk/volume, or published
 * data will disappear on every redeploy/restart. See README.md.
 */
const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');

const app = express();
const PORT = process.env.PORT || 3000;
const ADMIN_KEY = process.env.ADMIN_KEY || 'changeme'; // set this in your host's env vars!
const DATA_DIR = path.join(__dirname, 'data');

if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR, { recursive: true });
if (ADMIN_KEY === 'changeme') {
  console.warn('\n⚠️  ADMIN_KEY is not set — using the default "changeme".');
  console.warn('   Set a real ADMIN_KEY environment variable before going live!\n');
}

const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 50 * 1024 * 1024 } }); // 50MB — PCC's journal entry export runs much larger than the other sources

app.use(express.static(path.join(__dirname, 'public')));
app.use(express.json());

/* ------------------------------------------------------------------
   Access Profiles — scoped, tokenized links an admin creates and
   shares (e.g. yourapp.com/?access=abc123...). Each profile controls
   which home-screen tiles and which facilities that link can see.
   There's no username/password login for regular viewers — the link
   itself IS the credential — but enforcement happens here on the
   backend, not just by hiding things in the UI, so someone without a
   valid link genuinely can't pull the data another way. The plain
   base URL with no token is locked down by default (returns nothing)
   unless the request carries the real ADMIN_KEY, which always has
   full access to everything.
------------------------------------------------------------------ */
const ACCESS_PROFILES_FILE = path.join(DATA_DIR, 'access-profiles.json');
const VALID_TILES = ['ar', 'census', 'skilled', 'medicare'];

function loadAccessProfiles() {
  if (!fs.existsSync(ACCESS_PROFILES_FILE)) return [];
  try { return JSON.parse(fs.readFileSync(ACCESS_PROFILES_FILE, 'utf8')); } catch (e) { return []; }
}
function saveAccessProfiles(profiles) {
  fs.writeFileSync(ACCESS_PROFILES_FILE, JSON.stringify(profiles, null, 2));
}
function isAdminRequest(req) {
  return req.headers['x-admin-key'] === ADMIN_KEY;
}
// Resolves the requesting profile from either the real admin key (full
// access, no restrictions) or an x-access-token header. Returns null if
// neither is present/valid — callers should treat that as "no access".
function resolveAccessProfile(req) {
  if (isAdminRequest(req)) return { isAdmin: true, tiles: VALID_TILES, facilities: 'ALL' };
  const token = req.headers['x-access-token'];
  if (!token) return null;
  const profiles = loadAccessProfiles();
  const match = profiles.find(function (p) { return p.token === token; });
  return match ? { isAdmin: false, tiles: match.tiles, facilities: match.facilities } : null;
}
// Filters an array of records down to only the allowed facilities.
// facilities === 'ALL' (or the requester is admin) short-circuits to no
// filtering. facilityField is whichever key holds the facility name on
// that record shape ('facility' for AR sources, 'company' for Census/
// Skilled/Medicare ones).
function filterByFacilities(records, facilityField, allowedFacilities) {
  if (!Array.isArray(records)) return records;
  if (allowedFacilities === 'ALL') return records;
  const allowedSet = new Set(allowedFacilities || []);
  return records.filter(function (r) { return allowedSet.has(r[facilityField]); });
}

/* ------------------------------------------------------------------
   Sources — four independently-uploaded workbooks that get merged
   into one dataset on the frontend. Keep this list in sync with the
   SOURCES array in public/index.html.
------------------------------------------------------------------ */
const SOURCES = [
  { key: 'partA', label: 'Part A' },
  { key: 'partACoinsurance', label: 'Part A Coinsurance' },
  { key: 'partB', label: 'Part B' },
  { key: 'partBCoinsurance', label: 'Part B Coinsurance' }
];
const SOURCE_KEYS = SOURCES.map(s => s.key);
function dataFileFor(sourceKey) { return path.join(DATA_DIR, sourceKey + '.json'); }

// PCC AR sources — a facility that bills out of PointClickCare instead of
// the primary system. These are kept in their own files, separate from
// the four SOURCE_KEYS above, and merged into partA/partACoinsurance/
// partB/partBCoinsurance records only when /api/dataset is read — so a
// PCC upload can never overwrite the other facilities' own data, and vice
// versa. See buildPCCARRecords() below for the merge itself.
const PCC_AR_SOURCES = [
  { key: 'pccJournal', label: 'PCC Journal Entries' },
  { key: 'pccAging', label: 'PCC AR Aging Report' }
];
const PCC_AR_SOURCE_KEYS = PCC_AR_SOURCES.map(s => s.key);
function pccArFileFor(sourceKey) { return path.join(DATA_DIR, sourceKey + '.json'); }
// Company name (exactly as PCC's own exports spell it) -> [short display
// name, short code] — same convention as the frontend's facility
// crosswalk, just the 3 rows relevant to billing/AR.
const PCC_FACILITY_IDENTITY = {
  'Paramount - SNF': { facility: 'Paramount SNF', code: 'KSPA' },
  'Paramount  - ALF': { facility: 'Paramount ALF', code: 'KSPA ALF' },
  'Paramount - ILF': { facility: 'Paramount ILF', code: 'KSPA ILF' }
};

// PCC Medicare Sales — feeds Medicare Days only. Kept separate from the
// shared "salesJournal" source (which Skilled Analysis and Census Trends
// also read from) rather than merged into it, since this file only ever
// covers Medicare A specifically — merging it into salesJournal would make
// Skilled Analysis' Level Detail look wrong for Paramount (missing every
// other payer/rate code that a real Sales Journal would carry).
const PCC_MEDICARE_SOURCES = [
  { key: 'pccMedicareSales', label: 'PCC Medicare Sales' }
];
const PCC_MEDICARE_SOURCE_KEYS = PCC_MEDICARE_SOURCES.map(s => s.key);
function pccMedicareFileFor(sourceKey) { return path.join(DATA_DIR, sourceKey + '.json'); }

/* ------------------------------------------------------------------
   Same parsing logic as the browser version (kept in sync intentionally).
   Tolerant of both the original file layout (Sheet1 / "As Of" / "Month" /
   "Year") and the newer per-part layout ("Percentage Collected Summary" /
   "As Of Date" / "Month Of Service" / "Year of Service") — whichever a
   given workbook uses.
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
  const preferredSheets = ['Sheet1', 'Percentage Collected Summary'];
  const sheetName = preferredSheets.find(n => wb.SheetNames.includes(n)) || wb.SheetNames[0];
  const ws = wb.Sheets[sheetName];
  const grid = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
  if (!grid.length) throw new Error('The selected sheet appears to be empty.');

  const headerRow = grid[0].map(normHeader);
  function col(names) {
    var list = Array.isArray(names) ? names : [names];
    for (var i = 0; i < list.length; i++) {
      var idx = headerRow.indexOf(normHeader(list[i]));
      if (idx >= 0) return idx;
    }
    return -1;
  }
  const idx = {
    company: col('Company'), metric: col('Metric'),
    month: col(['Month', 'Month Of Service']),
    year: col(['Year', 'Year of Service']),
    asOf: col(['As Of', 'As Of Date']),
    code: col('Code'), facility: col('Facility')
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
    if (monthAbbr === 'Total' || year === 'Total') continue; // tolerate leftover Total rows
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
   Census Trends — a second, independent dataset from AR Collections,
   served from the same backend. Six import slots: the Length of Stay
   report, the Sales Journal (for Skilled Analysis), the Census Report
   (for the Census Days / Census % tabs), the Payer Mapping (built
   from the NCS Payer Listing's "List of Payors Detail" + "Sheet1"
   tabs), the Rate Key (Rate Code -> Rate Type, e.g. PDPM/Skilled/
   NE MCD/Levels/Bariatric — the primary source for the Skilled
   Analysis "Level Detail" tab, since it covers far more rate codes
   than the Payer Mapping's per-payor level descriptions do), and
   Available Beds (Code -> actual available bed count per facility —
   the accurate denominator for Census % occupancy, since the Census
   Report's own capacity figure is closer to total licensed beds) —
   these ship with built-in defaults (see public/index.html) but can
   be refreshed here whenever rates/contracts/bed counts change,
   without a code redeploy.
------------------------------------------------------------------ */
const CENSUS_SOURCES = [
  { key: 'losReport', label: 'Length of Stay Report' },
  { key: 'salesJournal', label: 'Sales Journal' },
  { key: 'censusReport', label: 'Census Report' },
  { key: 'payerMap', label: 'Payer Mapping' },
  { key: 'rateKey', label: 'Rate Key' },
  { key: 'availableBeds', label: 'Available Beds' },
  { key: 'pccJournal', label: 'PCC Journal Entries' }
];
const CENSUS_SOURCE_KEYS = CENSUS_SOURCES.map(s => s.key);
function censusFileFor(sourceKey) { return path.join(DATA_DIR, 'census-' + sourceKey + '.json'); }

function findHeaderRow(grid, firstColName) {
  var target = normHeader(firstColName);
  for (var i = 0; i < Math.min(grid.length, 10); i++) {
    if (grid[i] && normHeader(grid[i][0]) === target) return i;
  }
  return 0;
}

function parseDateFlexible(v) {
  if (v instanceof Date) return v;
  if (typeof v === 'string') {
    var m = v.trim().match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    // UTC, not the local-time Date constructor — these are calendar dates
    // with no real time-of-day, and building them via new Date(y,m,d) uses
    // whatever timezone this server process happens to run in, which is an
    // unnecessary and fragile dependency. Pinning to UTC midnight makes the
    // resulting ISO string deterministic regardless of server timezone —
    // paired with the frontend reading these back via getUTCFullYear()/
    // getUTCMonth() instead of the local-time equivalents.
    if (m) return new Date(Date.UTC(Number(m[3]), Number(m[1]) - 1, Number(m[2])));
  }
  return null;
}

// Parses the "CAVGLENSTAYCENS Detail" length-of-stay report: one row per
// completed stay segment, with Company/ID/Name/Days/Begin/End/DischargeTo/Payor.
function parseLOSWorkbookBuffer(buffer) {
  const wb = XLSX.read(buffer, { type: 'buffer', cellDates: true });
  const sheetName = wb.SheetNames.includes('CAVGLENSTAYCENS Detail') ? 'CAVGLENSTAYCENS Detail' : wb.SheetNames[0];
  const ws = wb.Sheets[sheetName];
  const grid = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
  if (!grid.length) throw new Error('The selected sheet appears to be empty.');

  const headerIdx = findHeaderRow(grid, 'Company Name');
  const headerRow = grid[headerIdx].map(normHeader);
  function col(name) { return headerRow.indexOf(normHeader(name)); }
  const idx = {
    company: col('Company Name'), id: col('ID #'), name: col('Name'), days: col('Days'),
    begin: col('Begin Date'), end: col('End Date'), dischargeTo: col('Discharge To'), payor: col('Payor')
  };
  if (idx.company < 0 || idx.days < 0 || idx.end < 0 || idx.payor < 0) {
    throw new Error('This file does not look like the expected Length of Stay layout.');
  }

  const records = [];
  for (let r = headerIdx + 1; r < grid.length; r++) {
    const row = grid[r];
    if (!row || row[idx.company] == null || row[idx.company] === '') continue;
    const endDate = parseDateFlexible(row[idx.end]);
    const beginDate = parseDateFlexible(row[idx.begin]);
    const days = row[idx.days];
    if (!endDate || typeof days !== 'number') continue;

    // Deliberately not carrying resident name or ID# into what gets stored/
    // published — only what's needed for the aggregate trend report, to keep
    // the shared payload free of resident PII (this endpoint has no auth on
    // reads, same as the AR Collections dataset).
    records.push({
      company: row[idx.company],
      days: days,
      beginISO: beginDate ? beginDate.toISOString() : null,
      endISO: endDate.toISOString(),
      dischargeTo: row[idx.dischargeTo],
      payorCode: row[idx.payor]
    });
  }
  if (!records.length) throw new Error('No usable rows were found in this file.');
  return records;
}

// Parses the "Inhouse Sales Journal Detail" sheet: one row per billing
// transaction line (a rate code applied over a date range), with Units
// (billed days for that line), Rate (per-day rate), Gross/Net revenue,
// and the payor code + the file's own Payor Name. The frontend combines
// Company+Code against the payer mapping to get a fine-grained category
// (e.g. "HMO Medicare PDPM") for the Skilled Analysis tab.
function parseSalesJournalWorkbookBuffer(buffer) {
  const wb = XLSX.read(buffer, { type: 'buffer', cellDates: true });
  const sheetName = wb.SheetNames.find(n => normHeader(n).indexOf('sales journal detail') >= 0) || wb.SheetNames[0];
  const ws = wb.Sheets[sheetName];
  const grid = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
  if (!grid.length) throw new Error('The selected sheet appears to be empty.');

  const headerIdx = findHeaderRow(grid, 'Company');
  const headerRow = grid[headerIdx].map(normHeader);
  function col(name) { return headerRow.indexOf(normHeader(name)); }
  const idx = {
    company: col('Company'), payor: col('Payor'), payorName: col('Payor Name'),
    start: col('Start Date'), end: col('End Date'), units: col('Units'),
    rate: col('Rate'), gross: col('Gross'), net: col('Net'), rateCode: col('Rate Code'),
    rawPayorType: col('Payor Type'), oav: col('OAV'), coinsuranceDed: col('Coinsurance Ded')
  };
  if (idx.company < 0 || idx.payor < 0 || idx.start < 0 || idx.units < 0 || idx.rate < 0) {
    throw new Error('This file does not look like the expected Sales Journal layout (missing Company / Payor / Start Date / Units / Rate columns).');
  }

  const records = [];
  for (let r = headerIdx + 1; r < grid.length; r++) {
    const row = grid[r];
    if (!row || row[idx.company] == null || row[idx.company] === '') continue;
    // OAV marks voided lines — both the original charge (lowercase "v")
    // and its reversing adjustment (uppercase "V") get excluded, since
    // together they're a matched pair that nets to zero and were never a
    // real stay.
    if (idx.oav >= 0 && row[idx.oav] != null && String(row[idx.oav]).trim() !== '') continue;
    const startDate = parseDateFlexible(row[idx.start]);
    const units = row[idx.units];
    if (!startDate || typeof units !== 'number') continue;

    records.push({
      company: row[idx.company],
      payorCode: row[idx.payor],
      payorName: row[idx.payorName] || '',
      rateCode: idx.rateCode >= 0 ? (row[idx.rateCode] || null) : null,
      rawPayorType: idx.rawPayorType >= 0 ? (row[idx.rawPayorType] || null) : null,
      startISO: startDate.toISOString(),
      units: units,
      rate: typeof row[idx.rate] === 'number' ? row[idx.rate] : null,
      gross: typeof row[idx.gross] === 'number' ? row[idx.gross] : null,
      net: typeof row[idx.net] === 'number' ? row[idx.net] : null,
      coinsuranceDed: (idx.coinsuranceDed >= 0 && typeof row[idx.coinsuranceDed] === 'number') ? row[idx.coinsuranceDed] : null
    });
  }
  if (!records.length) throw new Error('No usable rows were found in this file.');
  return records;
}

// Parses the Rate Key workbook: a flat (Rate Code, Rate Type) list — e.g.
// "0191" -> "Skilled", "203" -> "NE MCD", most codes -> "PDPM". This covers
// far more rate codes than the Payer Mapping's per-payor level descriptions
// do, so it's the primary source for the Level Detail tab; the Payer
// Mapping's named levels (Level 1, Bariatric, etc.) still take priority
// when both are available, since they're more specific.
function parseRateKeyWorkbookBuffer(buffer) {
  const wb = XLSX.read(buffer, { type: 'buffer', cellDates: true });
  const sheetName = wb.SheetNames[0];
  const ws = wb.Sheets[sheetName];
  const grid = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
  if (!grid.length) throw new Error('The selected sheet appears to be empty.');

  const headerIdx = findHeaderRow(grid, 'Rate Code');
  const headerRow = grid[headerIdx].map(normHeader);
  const iRateCode = headerRow.indexOf(normHeader('Rate Code'));
  const iRateType = headerRow.indexOf(normHeader('Rate Type'));
  if (iRateCode < 0 || iRateType < 0) {
    throw new Error('This file does not look like the expected Rate Key layout (missing Rate Code / Rate Type columns).');
  }

  const rows = [];
  for (let r = headerIdx + 1; r < grid.length; r++) {
    const row = grid[r];
    if (!row || row[iRateCode] == null || row[iRateType] == null) continue;
    rows.push({ rateCode: String(row[iRateCode]).trim(), rateType: String(row[iRateType]).trim() });
  }
  if (!rows.length) throw new Error('No usable rows were found in this file.');
  return rows;
}

// Parses the Available Beds workbook: a flat (Code, Facility, Available
// Beds) list, one row per facility — the accurate denominator for Census %
// occupancy. Skips the "Total" row (no Code) some exports include at the
// bottom, since that's a sum, not a facility.
function parseAvailableBedsWorkbookBuffer(buffer) {
  const wb = XLSX.read(buffer, { type: 'buffer', cellDates: true });
  const sheetName = wb.SheetNames[0];
  const ws = wb.Sheets[sheetName];
  const grid = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
  if (!grid.length) throw new Error('The selected sheet appears to be empty.');

  const headerIdx = findHeaderRow(grid, 'Code');
  const headerRow = grid[headerIdx].map(normHeader);
  const iCode = headerRow.indexOf(normHeader('Code'));
  const iFacility = headerRow.indexOf(normHeader('Facility'));
  const iBeds = headerRow.indexOf(normHeader('Available Beds'));
  if (iCode < 0 || iBeds < 0) {
    throw new Error('This file does not look like the expected Available Beds layout (missing Code / Available Beds columns).');
  }

  const rows = [];
  for (let r = headerIdx + 1; r < grid.length; r++) {
    const row = grid[r];
    if (!row || row[iCode] == null || row[iBeds] == null) continue;
    if (iFacility >= 0 && row[iFacility] == null) continue; // the "Total" row has no Facility name — skip it, it's a sum not a facility
    const beds = Number(row[iBeds]);
    if (isNaN(beds)) continue;
    rows.push({ code: String(row[iCode]).trim(), facility: iFacility >= 0 ? row[iFacility] : null, beds: beds });
  }
  if (!rows.length) throw new Error('No usable rows were found in this file.');
  return rows;
}

/* ------------------------------------------------------------------
   PCC (PointClickCare) Journal Entries — a single raw general-ledger
   export covering facilities that bill out of PCC instead of the
   primary system. One row per transaction line (charge, contractual
   adjustment, or cash), not pre-aggregated like the other sources.
   This derives TWO different things from the same file:

   1. Census days by facility/month/payer — from the dedicated "X Days"
      statistical GL accounts (Private Pay Days, Medicaid Days,
      Medicare A Days, Man Care/Private Ins. Days), which carry the
      actual day count in Units with $0 Amount (a memo entry separate
      from the real dollar charge on the same resident/period).

   2. AR Collections — Revenue, Cash Receipts, and Aging Balance,
      split into the same four buckets as the other facilities (Part A
      / Part A Coinsurance / Part B / Part B Coinsurance):
        - Revenue Code identifies Part A vs Part B directly: therapy
          codes 420/424/430/434/440/444 (PT/OT/Speech, standard UB-04
          revenue codes) are Part B; Room & Board charges (blank or
          101/110/120/183/185) are Part A. Verified against this
          file: zero overlap between the two code sets.
        - Coinsurance is identified by Payer Name explicitly saying
          "Medicare A coinsurance..." or "Medicare B coinsurance...",
          which also states which payer absorbed it (from Medicaid/
          private/insurance) — that "from X" is who the coinsurance
          dollar amount gets attributed to.
        - Revenue is a clean, direct split (each charge line carries
          both its revenue code and its payer). Cash Receipts and
          Aging Balance are NOT split by Part A/B — PCC posts all cash
          to one account and tracks one running A/R balance per payer,
          with no Part A/B distinction at that level, so splitting
          those would mean fabricating a breakdown the ledger doesn't
          actually contain. Both are reported combined under Part A,
          with Part B's cash/aging left empty — no dollars are lost,
          they're just not pre-split the way Revenue is.

   GL accounts observed and how each is used:
     A/R - {Payer}              balance-sheet account; cumulative sum
                                 of every posting is that payer's aging
                                 balance as of any given date.
     Contractual Allow - {Payer} adjustment against A/R — excluded from
                                 revenue, not separately used.
     Cash Depository            cash receipts; carries its own payer.
     {Anything else}            revenue — Room & Board (Part A),
                                 PT/OT/Speech (Part B via revenue code),
                                 plus smaller categories (Telephone,
                                 Meals, Transportation, Misc/Other
                                 Income, Drug Rev, Ancil Allow, etc.)
                                 which are all Part A by default since
                                 none of them carry a 4xx revenue code.
     {*} Days                   statistical only ($0), used for census.
     Bed Debt / Resident Refund / Exchange-Clearing Account
                                 excluded from both revenue and AR —
                                 exception/write-off categories, not
                                 regular billing activity.
------------------------------------------------------------------ */
const PCC_THERAPY_REVENUE_CODES = ['420','424','430','434','440','444'];
const PCC_EXCLUDED_ACCOUNTS = ['Bed Debt', 'Resident Refund', 'Exchange/Clearing Account'];
// Real dollars, but not resident/service revenue in the way a controller
// tracks it — verified against a real month's numbers (see below) that
// excluding these brings the derived total in line with the reported figure.
const PCC_NON_CORE_INCOME_ACCOUNTS = ['Misc Income', 'Other Income', 'Drug Rev Med A', 'Drug Rev Medicaid'];
function pccIsARAccount(desc) { return /^A\/R\s*-/.test(desc); }
function pccIsContractualAccount(desc) { return /^Contractual Allow/i.test(desc) || /^Cont Allow/i.test(desc); }
function pccIsDaysAccount(desc) { return /Days\s*$/.test(desc); }
// Maps PCC's own Payer Type/Name onto the payer columns already used by
// the other four AR Collections sources. Hospice is name-based (PCC's
// "Hospice ..." payors show Payer Type Medicaid/Private, not Hospice) so
// it's checked first; Medicaid gets a new "ksMedicaid" column, matching
// the existing NE/KY/OK state-specific pattern rather than merging into
// one of those.
function pccPayerColumnKey(payerName, payerType) {
  const name = String(payerName || '');
  const type = String(payerType || '').toUpperCase();
  if (/hospice/i.test(name)) return 'hospice';
  if (type === 'MEDICARE A' || type === 'MEDICARE B') return 'medicare';
  if (type === 'OUTPATIENT') return 'medicare'; // "Outpatient Medicare B" — still Medicare-sourced
  if (type === 'MEDICAID') return 'ksMedicaid';
  if (type === 'PRIVATE') return 'privatePay';
  if (type === 'MANAGED CARE') return 'hmo';
  return 'hmo'; // "Other" (Single Case Agreement, etc.) — small residual catch-all; NOT "budget", which is a distinct target/goal feature in AR Collections, not a real payer bucket
}
function pccParseDate(v) {
  if (!v) return null;
  const m = String(v).trim().match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (!m) return null;
  return new Date(Date.UTC(Number(m[3]), Number(m[1]) - 1, Number(m[2])));
}
const PCC_MONTH_ABBR = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];

function parsePCCJournalCSVBuffer(buffer) {
  const wb = XLSX.read(buffer, { type: 'buffer', raw: true });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const grid = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
  if (!grid.length) throw new Error('The selected file appears to be empty.');

  const headerRow = grid[0].map(normHeader);
  function col(name) { return headerRow.indexOf(normHeader(name)); }
  const idx = {
    facility: col('Facility Name'), year: col('Period Year'), month: col('Period Month'),
    payerName: col('Payer Name'), payerType: col('Payer Type'),
    revCode: col('Revenue Code'), glDesc: col('GL Account Desc'), units: col('Units'), amount: col('Amount')
  };
  if (idx.facility < 0 || idx.glDesc < 0 || idx.amount < 0) {
    throw new Error('This file does not look like the expected PCC Journal Entries layout.');
  }

  // Census days: sum Units on "X Days" accounts, grouped by facility/month/payer.
  const daysMap = {}; // key -> {company, monthAbbr, year, payorName, payorType, days}
  // AR Revenue: for each of the 4 buckets, sum Amount grouped by facility/
  // month/payer column. Values accumulate as SIGNED amounts here and only
  // get abs()'d once, on the final total, when building the output below
  // — summing abs() per-row was the bug that double-counted reversals/
  // corrections as if they were additional revenue instead of an offset.
  const arRevenue = { partA: {}, partACoinsurance: {}, partB: {}, partBCoinsurance: {} };
  const arCash = {}; // key -> {company, monthAbbr, year, values} — combined, reported under Part A

  for (let r = 1; r < grid.length; r++) {
    const row = grid[r];
    if (!row || row[idx.facility] == null) continue;
    const facility = String(row[idx.facility]).trim();
    const year = Number(row[idx.year]);
    const monthNum = Number(row[idx.month]); // 1-12
    if (!year || !monthNum || monthNum < 1 || monthNum > 12) continue;
    const monthAbbr = PCC_MONTH_ABBR[monthNum - 1];
    const glDesc = String(row[idx.glDesc] || '').trim();
    const payerName = row[idx.payerName] != null ? String(row[idx.payerName]).trim() : '';
    const payerType = row[idx.payerType] != null ? String(row[idx.payerType]).trim() : '';
    const amount = Number(row[idx.amount]);
    const monthKey = facility + '||' + year + '||' + monthAbbr;

    // 1) Census days
    if (pccIsDaysAccount(glDesc)) {
      const units = Number(row[idx.units]);
      if (!isNaN(units) && units !== 0) {
        const key = monthKey + '||' + payerName;
        if (!daysMap[key]) daysMap[key] = { company: facility, monthAbbr, year, payorName: payerName, payorType: payerType, days: 0 };
        daysMap[key].days += units;
      }
      continue;
    }

    if (isNaN(amount) || amount === 0) continue;

    // 2) AR — balance-sheet accounts, not revenue. Aging Balance is no
    // longer derived from this file at all (see parsePCCAgingReportCSVBuffer
    // below) — a journal extract only covers whatever date range it was
    // run for, so any balance a resident carried in from before that
    // window is invisible here, making a running balance built this way
    // structurally guaranteed to understate the true aging balance.
    if (pccIsARAccount(glDesc)) continue;

    if (pccIsContractualAccount(glDesc)) continue; // adjustment against A/R, not counted as revenue directly
    if (PCC_EXCLUDED_ACCOUNTS.indexOf(glDesc) >= 0) continue;
    if (PCC_NON_CORE_INCOME_ACCOUNTS.indexOf(glDesc) >= 0) continue; // Misc/Other Income, Drug Rev — not resident service revenue

    const payerCol = pccPayerColumnKey(payerName, payerType);
    if (glDesc === 'Cash Depository') {
      const key = monthKey;
      if (!arCash[key]) arCash[key] = { company: facility, monthAbbr, year, values: {} };
      arCash[key].values[payerCol] = (arCash[key].values[payerCol] || 0) + amount;
      continue;
    }

    // Remaining rows are revenue — split Part A/B by revenue code, or by
    // the payer name's explicit "Medicare A/B coinsurance" designation.
    const revCode = row[idx.revCode] != null ? String(row[idx.revCode]).trim() : '';
    const isCoinsurance = /coinsurance/i.test(payerName);
    let bucket;
    if (isCoinsurance) {
      bucket = /medicare a coinsurance/i.test(payerName) ? 'partACoinsurance' : 'partBCoinsurance';
    } else {
      bucket = PCC_THERAPY_REVENUE_CODES.indexOf(revCode) >= 0 ? 'partB' : 'partA';
    }
    // Coinsurance dollars are attributed to whichever payer absorbed them
    // (the "from X" part of the name), not to a Medicare column.
    let coinsurancePayerCol = payerCol;
    if (isCoinsurance) {
      const m = payerName.match(/from\s+(.+)$/i);
      const from = m ? m[1].trim().toLowerCase() : '';
      if (from === 'medicaid') coinsurancePayerCol = 'ksMedicaid';
      else if (from === 'private') coinsurancePayerCol = 'privatePay';
      else if (from === 'insurance') coinsurancePayerCol = 'hmo';
      else coinsurancePayerCol = payerCol;
    }
    const finalPayerCol = isCoinsurance ? coinsurancePayerCol : payerCol;
    const key = monthKey;
    if (!arRevenue[bucket][key]) arRevenue[bucket][key] = { company: facility, monthAbbr, year, values: {} };
    // Signed accumulation — NOT Math.abs() here. See the abs() applied
    // once, on the final total, when building the output records below.
    arRevenue[bucket][key].values[finalPayerCol] = (arRevenue[bucket][key].values[finalPayerCol] || 0) + amount;
  }

  // Revenue posts negative in this file's convention (a $100 charge shows
  // as -100 on the revenue account); abs() the final signed total once,
  // here, rather than per-row above, so reversals/corrections correctly
  // offset instead of compounding.
  function finalizeRevenueBucket(bucketMap) {
    return Object.keys(bucketMap).map(function (k) {
      const rec = bucketMap[k];
      const values = {};
      Object.keys(rec.values).forEach(function (payerCol) { values[payerCol] = Math.abs(rec.values[payerCol]); });
      return { company: rec.company, monthAbbr: rec.monthAbbr, year: rec.year, values: values };
    });
  }

  return {
    daysRecords: Object.keys(daysMap).map(function (k) { return daysMap[k]; }),
    arRevenue: {
      partA: finalizeRevenueBucket(arRevenue.partA),
      partACoinsurance: finalizeRevenueBucket(arRevenue.partACoinsurance),
      partB: finalizeRevenueBucket(arRevenue.partB),
      partBCoinsurance: finalizeRevenueBucket(arRevenue.partBCoinsurance)
    },
    arCash: Object.keys(arCash).map(function (k) { return arCash[k]; })
  };
}

/* ------------------------------------------------------------------
   PCC AR Aging Report — a real, authoritative per-resident/per-payer
   aging export (not derived from the journal), used directly as the
   source for the Aging Balance metric. Sum "Total" by Facility + Payer
   Type + As Of Date. Uploaded periodically, same as the other sources.
------------------------------------------------------------------ */
function parsePCCAgingReportCSVBuffer(buffer) {
  const wb = XLSX.read(buffer, { type: 'buffer', raw: true });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const grid = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
  if (!grid.length) throw new Error('The selected file appears to be empty.');

  const headerRow = grid[0].map(normHeader);
  function col(name) { return headerRow.indexOf(normHeader(name)); }
  const idx = {
    facility: col('Facility Name'), payerName: col('Payer Desc'), payerType: col('Payer Type'),
    asOf: col('As Of Date'), total: col('Total')
  };
  if (idx.facility < 0 || idx.total < 0 || idx.asOf < 0) {
    throw new Error('This file does not look like the expected PCC AR Aging Report layout.');
  }

  const agingMap = {}; // key -> {company, monthAbbr, year, asOfISO, values}
  for (let r = 1; r < grid.length; r++) {
    const row = grid[r];
    if (!row || row[idx.facility] == null) continue;
    const facility = String(row[idx.facility]).trim();
    const total = Number(row[idx.total]);
    if (isNaN(total)) continue;
    const asOfDate = pccParseDate(row[idx.asOf]);
    if (!asOfDate) continue;
    const asOfISO = asOfDate.toISOString();
    const monthAbbr = PCC_MONTH_ABBR[asOfDate.getUTCMonth()];
    const year = asOfDate.getUTCFullYear();
    const payerName = row[idx.payerName] != null ? String(row[idx.payerName]).trim() : '';
    const payerType = row[idx.payerType] != null ? String(row[idx.payerType]).trim() : '';
    const payerCol = pccPayerColumnKey(payerName, payerType);

    const key = facility + '||' + asOfISO;
    if (!agingMap[key]) agingMap[key] = { company: facility, monthAbbr: monthAbbr, year: year, asOfISO: asOfISO, values: {} };
    agingMap[key].values[payerCol] = (agingMap[key].values[payerCol] || 0) + total;
  }
  if (!Object.keys(agingMap).length) throw new Error('No usable rows were found in this file.');
  return Object.keys(agingMap).map(function (k) { return agingMap[k]; });
}

/* ------------------------------------------------------------------
   PCC Medicare Sales — a focused export of "A/R - Medicare A" postings
   only (not the general journal), used to build Medicare Days for a PCC
   facility. Formula validated against a real PCC "Revenue Reconciliation"
   report for July 2026 — Days and Gross matched exactly (144 days,
   $95,403.38), so this is trusted, not just a best guess:
     - GROSS = "R" (Room Charge) + "CR" (Contractual Adjustment/PDPM).
       R alone uses a placeholder daily rate; CR corrects it up to the
       real PDPM rate, so the two together are the true billed amount.
     - DAYS = sum of Units on the "R" rows.
     - COINSURANCE = "CO" rows where Payer Type is "Other" — the portion
       of the bill shifted to a secondary payer's coinsurance
       responsibility, tracked as its own mirrored entry.
     - NET = GROSS + "CO" rows where Payer Type is "Medicare A" (already
       negative — this is every reduction against Medicare's own A/R,
       coinsurance-shift included, so Net ends up correctly below Gross).
------------------------------------------------------------------ */
function parsePCCMedicareSalesCSVBuffer(buffer) {
  const wb = XLSX.read(buffer, { type: 'buffer', raw: true });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const grid = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
  if (!grid.length) throw new Error('The selected file appears to be empty.');

  const headerRow = grid[0].map(normHeader);
  function col(name) { return headerRow.indexOf(normHeader(name)); }
  const idx = {
    facility: col('Facility Name'), year: col('Period Year'), month: col('Period Month'),
    transType: col('Transaction Type'), payerType: col('Payer Type'), units: col('Units'), amount: col('Amount')
  };
  if (idx.facility < 0 || idx.transType < 0 || idx.amount < 0) {
    throw new Error('This file does not look like the expected PCC Medicare Sales layout.');
  }

  const byMonth = {}; // key -> {company, monthAbbr, year, days, gross, net, coinsurance}
  for (let r = 1; r < grid.length; r++) {
    const row = grid[r];
    if (!row || row[idx.facility] == null) continue;
    const facility = String(row[idx.facility]).trim();
    const year = Number(row[idx.year]);
    const monthNum = Number(row[idx.month]);
    if (!year || !monthNum || monthNum < 1 || monthNum > 12) continue;
    const monthAbbr = PCC_MONTH_ABBR[monthNum - 1];
    const transType = String(row[idx.transType] || '').trim();
    const payerType = String(row[idx.payerType] || '').trim();
    const amount = Number(row[idx.amount]);
    if (isNaN(amount)) continue;

    const key = facility + '||' + year + '||' + monthAbbr;
    if (!byMonth[key]) byMonth[key] = { company: facility, monthAbbr: monthAbbr, year: year, days: 0, gross: 0, net: 0, coinsurance: 0 };
    const bucket = byMonth[key];

    if (transType === 'R') {
      const units = Number(row[idx.units]);
      if (!isNaN(units)) bucket.days += units;
      bucket.gross += amount;
      bucket.net += amount;
    } else if (transType === 'CR') {
      bucket.gross += amount;
      bucket.net += amount;
    } else if (transType === 'CO' && payerType === 'Medicare A') {
      bucket.net += amount; // already negative — reduces Net below Gross
    } else if (transType === 'CO' && payerType === 'Other') {
      bucket.coinsurance += amount;
    }
    // Transaction Type "X" (rare, 1 row in the validation file) isn't part
    // of any of these four buckets — excluded rather than guessed at.
  }

  const records = Object.keys(byMonth).map(function (k) { return byMonth[k]; });
  if (!records.length) throw new Error('No usable rows were found in this file.');
  return records;
}


//   "Census Summary Detail": one row per facility/month/payor, with Days
//     (census days attributed to that payor that month). Summed per
//     facility+month, this is the "Census Days by month" tab.
//   "Census Summary Summary": an irregular block layout — 6 rows per
//     facility per month (Census/BedHolds/LOA/Day of Discharge/blank/blank
//     paired with Capacity/Total Beds Charged/Avg Beds Charged/Beds Not
//     Charged/Avg Beds Not Charged/% of Capacity Used), with NO explicit
//     month column. We match each block to a month by relying on a
//     verified property of this report: for every facility, the number of
//     6-row blocks always equals the number of distinct months for that
//     facility in the Detail sheet, and the blocks appear in the same
//     chronological order as those months — so we sort each facility's
//     months and zip them positionally with that facility's blocks. This
//     gives us "% of Capacity Used" (occupancy) per facility+month for the
//     "Census % by month" tab.
const MONTH_ABBR_LIST = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
function parseCensusReportWorkbookBuffer(buffer) {
  const wb = XLSX.read(buffer, { type: 'buffer', cellDates: true });
  const monthRe = /^([A-Za-z]{3})\s*-\s*(\d{4})$/;

  const detailSheetName = wb.SheetNames.find(n => normHeader(n).indexOf('census summary detail') >= 0) || wb.SheetNames[0];
  const wsD = wb.Sheets[detailSheetName];
  const gridD = XLSX.utils.sheet_to_json(wsD, { header: 1, raw: true, defval: null });
  if (!gridD.length) throw new Error('The Census Summary Detail sheet appears to be empty.');
  const headerD = gridD[0].map(normHeader);
  function colD(name) { return headerD.indexOf(normHeader(name)); }
  const idxD = { company: colD('Company Name'), month: colD('Month'), payorType: colD('Payor Type'), payor: colD('Payor'), days: colD('Days') };
  if (idxD.company < 0 || idxD.month < 0 || idxD.days < 0) {
    throw new Error('This file does not look like the expected Census Report layout (missing Company Name / Month / Days columns).');
  }

  const detailRecords = [];
  const monthsByCompany = {}; // company -> Set of "year-monthIdx" sort keys -> monthAbbr/year
  for (let r = 1; r < gridD.length; r++) {
    const row = gridD[r];
    if (!row || row[idxD.company] == null) continue;
    const m = monthRe.exec(String(row[idxD.month] || ''));
    if (!m) continue;
    const monthAbbr = m[1].charAt(0).toUpperCase() + m[1].slice(1, 3).toLowerCase();
    const monthIdx = MONTH_ABBR_LIST.indexOf(monthAbbr);
    if (monthIdx < 0) continue;
    const year = Number(m[2]);
    const days = row[idxD.days];
    if (typeof days !== 'number') continue;

    detailRecords.push({
      company: row[idxD.company], monthAbbr, year,
      payorType: row[idxD.payorType], payorName: row[idxD.payor], days
    });

    const company = row[idxD.company];
    monthsByCompany[company] = monthsByCompany[company] || {};
    monthsByCompany[company][year * 12 + monthIdx] = { monthAbbr, year };
  }
  if (!detailRecords.length) throw new Error('No usable rows were found in the Census Summary Detail sheet.');

  // ---- Summary sheet (occupancy %) — best-effort; skipped gracefully if absent ----
  const summaryRecords = [];
  const summarySheetName = wb.SheetNames.find(n => normHeader(n).indexOf('census summary summary') >= 0);
  if (summarySheetName) {
    const wsS = wb.Sheets[summarySheetName];
    const gridS = XLSX.utils.sheet_to_json(wsS, { header: 1, raw: true, defval: null });
    function num(v) {
      if (v == null) return null;
      const n = Number(String(v).replace(/,/g, ''));
      return isNaN(n) ? null : n;
    }
    const blocksByCompany = {};
    let i = 1; // skip header row
    while (i < gridS.length) {
      const row = gridS[i];
      if (!row || row[0] == null) { i++; continue; }
      const company = row[0];
      const block = gridS.slice(i, i + 6);
      i += 6;
      if (block.length < 6) break;
      (blocksByCompany[company] = blocksByCompany[company] || []).push({
        census: num(block[0][2]), capacity: num(block[0][5]),
        bedHolds: num(block[1][2]), totalBedsCharged: num(block[1][5]),
        loa: num(block[2][2]), avgBedsCharged: num(block[2][5]),
        dayOfDischarge: num(block[3][2]), bedsNotCharged: num(block[3][5]),
        avgBedsNotCharged: num(block[4][5]), pctUsed: num(block[5][5])
      });
    }
    Object.keys(blocksByCompany).forEach(function (company) {
      const sortedKeys = Object.keys(monthsByCompany[company] || {}).map(Number).sort((a, b) => a - b);
      const blocks = blocksByCompany[company];
      if (sortedKeys.length !== blocks.length) return; // counts don't line up — skip rather than guess
      blocks.forEach(function (b, idx) {
        const mo = monthsByCompany[company][sortedKeys[idx]];
        summaryRecords.push({
          company: company, monthAbbr: mo.monthAbbr, year: mo.year,
          census: b.census, capacity: b.capacity, pctUsed: b.pctUsed,
          bedHolds: b.bedHolds, totalBedsCharged: b.totalBedsCharged, avgBedsCharged: b.avgBedsCharged,
          dayOfDischarge: b.dayOfDischarge, bedsNotCharged: b.bedsNotCharged, avgBedsNotCharged: b.avgBedsNotCharged
        });
      });
    });
  }

  return { detailRecords, summaryRecords };
}

// Parses the NCS Payer Listing workbook: "List of Payors Detail" maps
// (Company, Code) -> Payor name (this is the authoritative, granular sheet —
// it's what actually changes when new contracts/rates come in); "Sheet1"
// maps Payor name -> broader Payer Type category.
function parsePayerMapWorkbookBuffer(buffer) {
  const wb = XLSX.read(buffer, { type: 'buffer', cellDates: true });
  const dataRows = [];
  const rateDetailRows = [];
  if (wb.SheetNames.includes('List of Payors Detail')) {
    const ws = wb.Sheets['List of Payors Detail'];
    const grid = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
    const headerIdx = findHeaderRow(grid, 'State');
    const headerRow = grid[headerIdx].map(normHeader);
    const iCompany = headerRow.indexOf(normHeader('Company Name'));
    const iCode = headerRow.indexOf(normHeader('Code'));
    const iPayor = headerRow.indexOf(normHeader('Payor'));
    const iRateCode = headerRow.indexOf(normHeader('Rate Code'));
    const iDescription = headerRow.indexOf(normHeader('Description'));
    // Optional — a dedicated override column for what the Level Detail tab
    // should show for that rate code. Unlike Description (which is also
    // used for other things and sometimes just echoes the code back),
    // anything here is used as-is, no filtering/guessing needed, since
    // it's a deliberate label rather than a repurposed field.
    const iLevelDetail = headerRow.indexOf(normHeader('Level Detail'));
    const seen = {};
    for (let r = headerIdx + 1; r < grid.length; r++) {
      const row = grid[r];
      if (!row || row[iCompany] == null || row[iCode] == null) continue;
      const key = row[iCompany] + '||' + row[iCode];
      if (!seen[key]) {
        seen[key] = true; // de-dupe — this sheet has one row per rate code, many per payor
        dataRows.push({ company: row[iCompany], code: row[iCode], payorName: row[iPayor] });
      }
      // Rate/level detail keeps full granularity — every rate code the payor
      // has, not just the first — this is what the Skilled Analysis "Level
      // Detail" tab uses (e.g. Level 1/2/3, Bariatric, Trach).
      const hasDescription = iDescription >= 0 && row[iDescription] != null;
      const hasLevelDetail = iLevelDetail >= 0 && row[iLevelDetail] != null;
      if (iRateCode >= 0 && row[iRateCode] != null && (hasDescription || hasLevelDetail)) {
        rateDetailRows.push({
          company: row[iCompany], code: row[iCode], rateCode: row[iRateCode],
          description: hasDescription ? row[iDescription] : null,
          levelDetail: hasLevelDetail ? row[iLevelDetail] : null
        });
      }
    }
  }
  const typeRows = [];
  if (wb.SheetNames.includes('Sheet1')) {
    const ws2 = wb.Sheets['Sheet1'];
    const grid2 = XLSX.utils.sheet_to_json(ws2, { header: 1, raw: true, defval: null });
    const headerIdx2 = findHeaderRow(grid2, 'Payer Type');
    const headerRow2 = grid2[headerIdx2].map(normHeader);
    const iType = headerRow2.indexOf(normHeader('Payer Type'));
    const iPayor2 = headerRow2.indexOf(normHeader('Payor'));
    for (let r = headerIdx2 + 1; r < grid2.length; r++) {
      const row = grid2[r];
      if (!row || row[iType] == null) continue;
      typeRows.push({ payerType: row[iType], payorName: row[iPayor2] });
    }
  }
  if (!dataRows.length && !typeRows.length) throw new Error('No usable rows were found in this file.');
  return { dataRows, typeRows, rateDetailRows };
}

/* ------------------------------------------------------------------
   API
------------------------------------------------------------------ */

function pccMonthEndISO(year, monthAbbr) {
  const monthNum = PCC_MONTH_ABBR.indexOf(monthAbbr); // 0-11
  if (monthNum < 0) return null;
  return new Date(Date.UTC(year, monthNum + 1, 0, 23, 59, 59)).toISOString();
}

// Builds the PCC-derived records for one of the four AR Collections
// sources, in the exact shape the frontend expects — read fresh on every
// request rather than written into partA.json etc, so uploading PCC data
// can never overwrite or collide with the other facilities' own uploads.
// Cash Receipts and Aging Balance aren't split by Part A/B (see the notes
// on parsePCCJournalCSVBuffer) — both are only added under 'partA'.
function buildPCCARRecords(bucketKey) {
  const records = [];
  const journalFile = pccArFileFor('pccJournal');
  if (fs.existsSync(journalFile)) {
    const payload = JSON.parse(fs.readFileSync(journalFile, 'utf8'));
    const parsed = payload.data;
    (parsed.arRevenue[bucketKey] || []).forEach(function (rec) {
      const identity = PCC_FACILITY_IDENTITY[rec.company];
      if (!identity) return;
      records.push({
        company: rec.company, metric: 'Revenue', facility: identity.facility, code: identity.code,
        monthAbbr: rec.monthAbbr, year: rec.year, asOfISO: pccMonthEndISO(rec.year, rec.monthAbbr), values: rec.values
      });
    });
    if (bucketKey === 'partA') {
      (parsed.arCash || []).forEach(function (rec) {
        const identity = PCC_FACILITY_IDENTITY[rec.company];
        if (!identity) return;
        // The other AR Collections sources store Cash Receipts as negative
        // (the frontend flips the sign when displaying it) — PCC's own
        // data is genuinely positive (real cash in), so it needs negating
        // here to match that stored convention, not just at display time.
        const negatedValues = {};
        Object.keys(rec.values).forEach(function (k) { negatedValues[k] = -rec.values[k]; });
        records.push({
          company: rec.company, metric: 'Cash Receipts', facility: identity.facility, code: identity.code,
          monthAbbr: rec.monthAbbr, year: rec.year, asOfISO: pccMonthEndISO(rec.year, rec.monthAbbr), values: negatedValues
        });
      });
    }
  }
  if (bucketKey === 'partA') {
    const agingFile = pccArFileFor('pccAging');
    if (fs.existsSync(agingFile)) {
      const payload = JSON.parse(fs.readFileSync(agingFile, 'utf8'));
      (payload.data || []).forEach(function (rec) {
        const identity = PCC_FACILITY_IDENTITY[rec.company];
        if (!identity) return;
        records.push({
          company: rec.company, metric: 'Aging Bals', facility: identity.facility, code: identity.code,
          monthAbbr: rec.monthAbbr, year: rec.year, asOfISO: rec.asOfISO, values: rec.values
        });
      });
    }
  }
  return records;
}

// Anyone with the link can read the current shared dataset — all four
// sources combined into one object, keyed by source. A source that
// hasn't been published yet comes back as null. PCC-derived records (if
// any PCC file has been uploaded) are appended into the matching source
// here, not stored in partA.json etc directly.
app.get('/api/dataset', (req, res) => {
  const profile = resolveAccessProfile(req);
  if (!profile) return res.status(401).json({ error: 'This link is not authorized to view this data.' });
  const out = {};
  if (profile.tiles.indexOf('ar') === -1) return res.json(out); // valid link, just not for this tile — empty, not an error
  SOURCE_KEYS.forEach(function(key) {
    const file = dataFileFor(key);
    const basePayload = fs.existsSync(file) ? JSON.parse(fs.readFileSync(file, 'utf8')) : { fileName: null, uploadedAt: null, records: [] };
    const combinedRecords = (basePayload.records || []).concat(buildPCCARRecords(key));
    if (!combinedRecords.length) { out[key] = null; return; }
    const filteredRecords = filterByFacilities(combinedRecords, 'facility', profile.facilities);
    out[key] = Object.assign({}, basePayload, { records: filteredRecords, rowCount: filteredRecords.length });
  });
  // Upload status for the two PCC AR sources (filename/timestamp only, not
  // figures) — only admin sessions need this, to show "already loaded" in
  // the Admin panel rather than always saying "Import file...".
  if (profile.isAdmin) {
    out._pccArStatus = {};
    PCC_AR_SOURCE_KEYS.forEach(function (key) {
      const file = pccArFileFor(key);
      if (!fs.existsSync(file)) { out._pccArStatus[key] = null; return; }
      const payload = JSON.parse(fs.readFileSync(file, 'utf8'));
      out._pccArStatus[key] = { fileName: payload.fileName, uploadedAt: payload.uploadedAt, rowCount: payload.rowCount };
    });
  }
  res.json(out);
});

// Only requests with the correct x-admin-key header can publish. Each of
// the four sources is uploaded and published independently of the others.
app.post('/api/upload/:sourceKey', upload.single('file'), (req, res) => {
  try {
    const sourceKey = req.params.sourceKey;
    if (SOURCE_KEYS.indexOf(sourceKey) === -1) {
      return res.status(400).json({ error: 'Unknown data source "' + sourceKey + '".' });
    }
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
    fs.writeFileSync(dataFileFor(sourceKey), JSON.stringify(payload));
    res.json({ ok: true, sourceKey: sourceKey, rowCount: records.length, fileName: req.file.originalname });
  } catch (err) {
    res.status(400).json({ error: err.message });
  }
});

// PCC AR sources — Journal Entries (Revenue/Cash) and the AR Aging Report
// (Aging Balance). Stored separately from the four SOURCE_KEYS files;
// buildPCCARRecords() merges them in at read time. rowCount here reflects
// something meaningful per source (days records for the journal, aging
// rows for the aging report) rather than a raw CSV line count.
// Merges two arrays of records by a composite key — used so uploading a
// new PCC file (a new month's journal, a new aging snapshot) ADDS to what's
// already stored instead of replacing it wholesale. New records win when
// the same key appears in both (e.g. re-uploading a corrected month), but
// anything only in the existing array (other months/snapshots) is kept.
function mergeRecordsByKey(existingArray, newArray, keyFn) {
  const map = {};
  (existingArray || []).forEach(function (r) { map[keyFn(r)] = r; });
  (newArray || []).forEach(function (r) { map[keyFn(r)] = r; }); // new wins on conflict
  return Object.keys(map).map(function (k) { return map[k]; });
}
function pccJournalRecordKey(r) { return r.company + '||' + r.year + '||' + r.monthAbbr; }
function pccAgingRecordKey(r) { return r.company + '||' + r.asOfISO; }

app.post('/api/upload-pcc/:sourceKey', upload.single('file'), (req, res) => {
  try {
    const sourceKey = req.params.sourceKey;
    if (PCC_AR_SOURCE_KEYS.indexOf(sourceKey) === -1) {
      return res.status(400).json({ error: 'Unknown data source "' + sourceKey + '".' });
    }
    if (req.headers['x-admin-key'] !== ADMIN_KEY) {
      return res.status(401).json({ error: 'Invalid admin key.' });
    }
    if (!req.file) {
      return res.status(400).json({ error: 'No file was uploaded.' });
    }
    const existingFile = pccArFileFor(sourceKey);
    const existing = fs.existsSync(existingFile) ? JSON.parse(fs.readFileSync(existingFile, 'utf8')) : null;

    let data, rowCount;
    if (sourceKey === 'pccJournal') {
      const parsed = parsePCCJournalCSVBuffer(req.file.buffer);
      // Merge each piece separately, by (facility, month) — a monthly
      // upload only ever covers the months present in that file, so this
      // adds/replaces just those months without touching any others
      // already on file from earlier uploads.
      const existingData = (existing && existing.data) || {};
      data = {
        daysRecords: mergeRecordsByKey(existingData.daysRecords, parsed.daysRecords, function (r) { return pccJournalRecordKey(r) + '||' + r.payorName; }),
        arRevenue: {
          partA: mergeRecordsByKey(existingData.arRevenue && existingData.arRevenue.partA, parsed.arRevenue.partA, pccJournalRecordKey),
          partACoinsurance: mergeRecordsByKey(existingData.arRevenue && existingData.arRevenue.partACoinsurance, parsed.arRevenue.partACoinsurance, pccJournalRecordKey),
          partB: mergeRecordsByKey(existingData.arRevenue && existingData.arRevenue.partB, parsed.arRevenue.partB, pccJournalRecordKey),
          partBCoinsurance: mergeRecordsByKey(existingData.arRevenue && existingData.arRevenue.partBCoinsurance, parsed.arRevenue.partBCoinsurance, pccJournalRecordKey)
        },
        arCash: mergeRecordsByKey(existingData.arCash, parsed.arCash, pccJournalRecordKey)
      };
      rowCount = data.daysRecords.length + data.arRevenue.partA.length + data.arRevenue.partACoinsurance.length +
        data.arRevenue.partB.length + data.arRevenue.partBCoinsurance.length + data.arCash.length;
    } else {
      const parsed = parsePCCAgingReportCSVBuffer(req.file.buffer);
      // Merge by (facility, as-of date) — each aging snapshot is its own
      // point in time; uploading a new one adds a new snapshot rather than
      // replacing whatever snapshot was there before. This is exactly what
      // lets a PCC facility build up the same "as of" history over time
      // that the other facilities' richer reports give in one file.
      data = mergeRecordsByKey(existing && existing.data, parsed, pccAgingRecordKey);
      rowCount = data.length;
    }
    const payload = { fileName: req.file.originalname, uploadedAt: new Date().toISOString(), rowCount: rowCount, data: data };
    fs.writeFileSync(pccArFileFor(sourceKey), JSON.stringify(payload));
    res.json({ ok: true, sourceKey: sourceKey, rowCount: rowCount, fileName: req.file.originalname });
  } catch (err) {
    res.status(400).json({ error: err.message });
  }
});

// PCC Medicare Sales — a separate, dedicated route so this stays fully
// independent from the AR sources above (different file, different
// storage, different downstream consumer).
app.post('/api/upload-pcc-medicare/:sourceKey', upload.single('file'), (req, res) => {
  try {
    const sourceKey = req.params.sourceKey;
    if (PCC_MEDICARE_SOURCE_KEYS.indexOf(sourceKey) === -1) {
      return res.status(400).json({ error: 'Unknown data source "' + sourceKey + '".' });
    }
    if (req.headers['x-admin-key'] !== ADMIN_KEY) {
      return res.status(401).json({ error: 'Invalid admin key.' });
    }
    if (!req.file) {
      return res.status(400).json({ error: 'No file was uploaded.' });
    }
    const existingFile = pccMedicareFileFor(sourceKey);
    const existing = fs.existsSync(existingFile) ? JSON.parse(fs.readFileSync(existingFile, 'utf8')) : null;
    const parsed = parsePCCMedicareSalesCSVBuffer(req.file.buffer);
    // Same reasoning as the AR merge above — a monthly upload should add
    // to what's on file, not replace months it doesn't even cover.
    const data = mergeRecordsByKey(existing && existing.data, parsed, pccJournalRecordKey);
    const payload = { fileName: req.file.originalname, uploadedAt: new Date().toISOString(), rowCount: data.length, data: data };
    fs.writeFileSync(pccMedicareFileFor(sourceKey), JSON.stringify(payload));
    res.json({ ok: true, sourceKey: sourceKey, rowCount: data.length, fileName: req.file.originalname });
  } catch (err) {
    res.status(400).json({ error: err.message });
  }
});

app.get('/api/health', (req, res) => res.json({ ok: true }));

// A one-time "log in" check for the Admin tab — the frontend stores the
// key client-side (sessionStorage) after this succeeds, so the person
// doesn't have to re-type it for every subsequent import that session.
// The real enforcement is still per-request on each upload endpoint below;
// this just gives immediate pass/fail feedback at the login step.
app.post('/api/admin/verify', (req, res) => {
  if (req.headers['x-admin-key'] === ADMIN_KEY) {
    res.json({ ok: true });
  } else {
    res.status(401).json({ error: 'Incorrect admin password.' });
  }
});

// --- Access Profiles: admin-only CRUD, plus one public lookup route ---

app.get('/api/admin/access-profiles', (req, res) => {
  if (!isAdminRequest(req)) return res.status(401).json({ error: 'Invalid admin key.' });
  res.json({ profiles: loadAccessProfiles() });
});

app.post('/api/admin/access-profiles', (req, res) => {
  if (!isAdminRequest(req)) return res.status(401).json({ error: 'Invalid admin key.' });
  const body = req.body || {};
  const label = (body.label || '').trim();
  if (!label) return res.status(400).json({ error: 'A label is required.' });
  const tiles = Array.isArray(body.tiles) ? body.tiles.filter(function (t) { return VALID_TILES.indexOf(t) >= 0; }) : [];
  if (!tiles.length) return res.status(400).json({ error: 'At least one tile must be selected.' });
  const facilities = body.facilities === 'ALL' ? 'ALL' : (Array.isArray(body.facilities) ? body.facilities : []);
  if (facilities !== 'ALL' && !facilities.length) return res.status(400).json({ error: 'Select at least one facility, or choose all facilities.' });
  const email = (body.email || '').trim(); // record-keeping only — who this link was meant for. Not enforced: whoever has the link can use it, this is just a paper trail.
  // Just declutters that link's home screen — Admin is still protected by
  // the real admin password regardless of this setting, so it's not a
  // security boundary, only a "don't show them a tile they don't need" one.
  const showAdmin = body.showAdmin !== false;
  // Opt-IN, unlike showAdmin above — exporting data means it leaves the
  // app entirely, so a scoped link shouldn't get that by default just
  // because nothing was specified. Must be explicitly granted.
  const canExport = body.canExport === true;

  const profiles = loadAccessProfiles();
  const profile = {
    id: crypto.randomBytes(8).toString('hex'),
    token: crypto.randomBytes(20).toString('hex'),
    label: label,
    email: email,
    tiles: tiles,
    facilities: facilities,
    showAdmin: showAdmin,
    canExport: canExport,
    createdAt: new Date().toISOString()
  };
  profiles.push(profile);
  saveAccessProfiles(profiles);
  res.json({ ok: true, profile: profile });
});

app.put('/api/admin/access-profiles/:id', (req, res) => {
  if (!isAdminRequest(req)) return res.status(401).json({ error: 'Invalid admin key.' });
  const profiles = loadAccessProfiles();
  const idx = profiles.findIndex(function (p) { return p.id === req.params.id; });
  if (idx === -1) return res.status(404).json({ error: 'Access link not found.' });
  const body = req.body || {};
  if (typeof body.label === 'string' && body.label.trim()) profiles[idx].label = body.label.trim();
  if (typeof body.email === 'string') profiles[idx].email = body.email.trim();
  if (Array.isArray(body.tiles)) {
    const tiles = body.tiles.filter(function (t) { return VALID_TILES.indexOf(t) >= 0; });
    if (tiles.length) profiles[idx].tiles = tiles;
  }
  if (body.facilities === 'ALL' || Array.isArray(body.facilities)) {
    if (body.facilities === 'ALL' || body.facilities.length) profiles[idx].facilities = body.facilities;
  }
  if (typeof body.showAdmin === 'boolean') profiles[idx].showAdmin = body.showAdmin;
  if (typeof body.canExport === 'boolean') profiles[idx].canExport = body.canExport;
  saveAccessProfiles(profiles);
  res.json({ ok: true, profile: profiles[idx] });
});

app.delete('/api/admin/access-profiles/:id', (req, res) => {
  if (!isAdminRequest(req)) return res.status(401).json({ error: 'Invalid admin key.' });
  const profiles = loadAccessProfiles();
  const next = profiles.filter(function (p) { return p.id !== req.params.id; });
  if (next.length === profiles.length) return res.status(404).json({ error: 'Access link not found.' });
  saveAccessProfiles(next);
  res.json({ ok: true });
});

// Public — the frontend calls this with whatever token is in the URL
// (?access=...) to find out what that link is allowed to see. No admin
// key needed here since this IS the regular-viewer "login" — the token
// itself is the credential, resolved against the stored profiles.
app.get('/api/access/:token', (req, res) => {
  const profiles = loadAccessProfiles();
  const match = profiles.find(function (p) { return p.token === req.params.token; });
  if (!match) return res.status(404).json({ error: 'This access link is not valid or has been revoked.' });
  // Profiles created before this setting existed didn't store it — default
  // to showing the tile so old links keep behaving exactly as they did.
  const showAdmin = match.showAdmin !== false;
  // Opposite default from showAdmin — old profiles never had export as a
  // concept at all, so there's no prior behavior to preserve here; only
  // profiles explicitly granted it get it.
  const canExport = match.canExport === true;
  res.json({ ok: true, label: match.label, tiles: match.tiles, facilities: match.facilities, showAdmin: showAdmin, canExport: canExport });
});

// Census Trends — same read-open / publish-gated pattern as the AR sources.
app.get('/api/census-dataset', (req, res) => {
  const profile = resolveAccessProfile(req);
  if (!profile) return res.status(401).json({ error: 'This link is not authorized to view this data.' });
  const out = {};
  const canCensus = profile.tiles.indexOf('census') >= 0;
  const canSkilledOrMedicare = profile.tiles.indexOf('skilled') >= 0 || profile.tiles.indexOf('medicare') >= 0;

  function readFile(key) {
    const file = censusFileFor(key);
    return fs.existsSync(file) ? JSON.parse(fs.readFileSync(file, 'utf8')) : null;
  }

  out.losReport = null;
  if (canCensus) {
    const payload = readFile('losReport');
    if (payload) {
      const records = filterByFacilities(payload.records, 'company', profile.facilities);
      out.losReport = Object.assign({}, payload, { records: records, rowCount: records.length });
    }
  }
  out.censusReport = null;
  if (canCensus) {
    const payload = readFile('censusReport');
    if (payload) {
      const detailRecords = filterByFacilities(payload.detailRecords, 'company', profile.facilities);
      const summaryRecords = filterByFacilities(payload.summaryRecords, 'company', profile.facilities);
      out.censusReport = Object.assign({}, payload, { detailRecords: detailRecords, summaryRecords: summaryRecords, rowCount: detailRecords.length });
    }
  }
  out.salesJournal = null;
  if (canSkilledOrMedicare) {
    const payload = readFile('salesJournal');
    if (payload) {
      const records = filterByFacilities(payload.records, 'company', profile.facilities);
      out.salesJournal = Object.assign({}, payload, { records: records, rowCount: records.length });
    }
  }
  // Reference/lookup data, not per-resident data — not facility-specific,
  // so no filtering, just the same tile gate as the sources that use it.
  out.payerMap = (canCensus || canSkilledOrMedicare) ? readFile('payerMap') : null;
  out.rateKey = canSkilledOrMedicare ? readFile('rateKey') : null;
  out.availableBeds = canCensus ? readFile('availableBeds') : null;

  // PCC Medicare Sales — separate from salesJournal above (see the note
  // by PCC_MEDICARE_SOURCES for why), already pre-aggregated by facility/
  // month, no further filtering needed.
  out.pccMedicareSales = null;
  if (canSkilledOrMedicare) {
    const file = pccMedicareFileFor('pccMedicareSales');
    if (fs.existsSync(file)) out.pccMedicareSales = JSON.parse(fs.readFileSync(file, 'utf8'));
  }

  res.json(out);
});

app.post('/api/census-upload/:sourceKey', upload.single('file'), (req, res) => {
  try {
    const sourceKey = req.params.sourceKey;
    if (CENSUS_SOURCE_KEYS.indexOf(sourceKey) === -1) {
      return res.status(400).json({ error: 'Unknown data source "' + sourceKey + '".' });
    }
    if (req.headers['x-admin-key'] !== ADMIN_KEY) {
      return res.status(401).json({ error: 'Invalid admin key.' });
    }
    if (!req.file) {
      return res.status(400).json({ error: 'No file was uploaded.' });
    }
    let payload;
    if (sourceKey === 'losReport') {
      const records = parseLOSWorkbookBuffer(req.file.buffer);
      payload = { fileName: req.file.originalname, uploadedAt: new Date().toISOString(), rowCount: records.length, records };
    } else if (sourceKey === 'salesJournal') {
      const records = parseSalesJournalWorkbookBuffer(req.file.buffer);
      payload = { fileName: req.file.originalname, uploadedAt: new Date().toISOString(), rowCount: records.length, records };
    } else if (sourceKey === 'censusReport') {
      const parsed = parseCensusReportWorkbookBuffer(req.file.buffer);
      payload = { fileName: req.file.originalname, uploadedAt: new Date().toISOString(), rowCount: parsed.detailRecords.length, detailRecords: parsed.detailRecords, summaryRecords: parsed.summaryRecords };
    } else if (sourceKey === 'rateKey') {
      const rows = parseRateKeyWorkbookBuffer(req.file.buffer);
      payload = { fileName: req.file.originalname, uploadedAt: new Date().toISOString(), rowCount: rows.length, rows: rows };
    } else if (sourceKey === 'availableBeds') {
      const rows = parseAvailableBedsWorkbookBuffer(req.file.buffer);
      payload = { fileName: req.file.originalname, uploadedAt: new Date().toISOString(), rowCount: rows.length, rows: rows };
    } else {
      const parsed = parsePayerMapWorkbookBuffer(req.file.buffer);
      payload = { fileName: req.file.originalname, uploadedAt: new Date().toISOString(), rowCount: parsed.dataRows.length + parsed.typeRows.length, dataRows: parsed.dataRows, typeRows: parsed.typeRows, rateDetailRows: parsed.rateDetailRows };
    }
    fs.writeFileSync(censusFileFor(sourceKey), JSON.stringify(payload));
    res.json({ ok: true, sourceKey: sourceKey, rowCount: payload.rowCount, fileName: req.file.originalname });
  } catch (err) {
    res.status(400).json({ error: err.message });
  }
});

app.listen(PORT, () => {
  console.log('AR Collections Explorer listening on port ' + PORT);
});
