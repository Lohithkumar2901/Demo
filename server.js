
const express = require('express');
const cors = require('cors');
const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');

const app = express();
const PORT = 3000;

// STATIC EXCEL PATH (change file name if needed)
const EXCEL_PATH = 'D:\\demo\\Project-Management-Sample-Data.xlsx';

// "Website storage" = JSON file on server
const STORAGE_DIR = path.join(__dirname, 'storage');
const STORAGE_FILE = path.join(STORAGE_DIR, 'data.json');

// Middlewares
app.use(cors());
app.use(express.json());

// Serve frontend from /public
app.use(express.static(path.join(__dirname, 'public')));

// Helper: ensure storage folder exists
function ensureStorageDir() {
  if (!fs.existsSync(STORAGE_DIR)) {
    fs.mkdirSync(STORAGE_DIR, { recursive: true });
  }
}

// POST: Import Excel from static path and save to website storage (JSON file)
// Helper: Find unique key for a row (customize as needed)
function getRowKey(row) {
  // Try to use 'id', else fallback to stringified row
  if (row.id !== undefined && row.id !== null) return String(row.id);
  // If no 'id', use a composite key of all values (not ideal, but fallback)
  return Object.values(row).join('|');
}

// Helper: Validate a row (customize rules as needed)
function validateRow(row, referenceRow) {
  // If referenceRow is provided, check types match
  if (referenceRow) {
    for (const key of Object.keys(referenceRow)) {
      if (row[key] !== null && row[key] !== undefined) {
        if (typeof row[key] !== typeof referenceRow[key]) {
          // Allow number <-> string conversion if possible
          if (!(typeof row[key] === 'string' && typeof referenceRow[key] === 'number' && !isNaN(Number(row[key])))) {
            return `Type mismatch for field '${key}': expected ${typeof referenceRow[key]}, got ${typeof row[key]}`;
          }
        }
      }
    }
  }
  // Add more validation rules as needed
  return null;
}


app.post('/api/import-excel', (req, res) => {
  try {
    if (!fs.existsSync(EXCEL_PATH)) {
      return res.status(404).json({ message: `Excel file not found at ${EXCEL_PATH}` });
    }

    const workbook = XLSX.readFile(EXCEL_PATH);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const newData = XLSX.utils.sheet_to_json(worksheet, { defval: null });

    ensureStorageDir();
    fs.writeFileSync(STORAGE_FILE, JSON.stringify(newData, null, 2), 'utf-8');

    return res.json({
      message: 'Excel imported and storage overwritten with latest data.',
      rows: newData.length,
    });
  } catch (err) {
    console.error(err);
    return res.status(500).json({ message: 'Error importing Excel', error: err.message });
  }
});

// GET: Return stored data from website storage
app.get('/api/data', (req, res) => {
  try {
    if (!fs.existsSync(STORAGE_FILE)) {
      return res.json([]); // no data yet
    }
    const raw = fs.readFileSync(STORAGE_FILE, 'utf-8');
    const data = JSON.parse(raw || '[]');
    return res.json(data);
  } catch (err) {
    console.error(err);
    return res.status(500).json({ message: 'Error reading stored data', error: err.message });
  }
});

// Health check
app.get('/api/health', (req, res) => {
  res.json({ status: 'ok' });
});

app.listen(PORT, () => {
  console.log(`Server listening on http://localhost:${PORT}`);
}); 