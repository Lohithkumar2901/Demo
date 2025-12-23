const express = require('express');
const cors = require('cors');
const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');

const app = express();
const PORT = 3000;

// STATIC EXCEL PATH (change file name if needed)
const EXCEL_PATH = 'D:\\demo\\data.csv';

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
app.post('/api/import-excel', (req, res) => {
  try {
    if (!fs.existsSync(EXCEL_PATH)) {
      return res.status(404).json({ message: `Excel file not found at ${EXCEL_PATH}` });
    }

    const workbook = XLSX.readFile(EXCEL_PATH);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];

    const data = XLSX.utils.sheet_to_json(worksheet, { defval: null });

    ensureStorageDir();
    fs.writeFileSync(STORAGE_FILE, JSON.stringify(data, null, 2), 'utf-8');

    return res.json({
      message: 'Excel imported and stored successfully',
      rows: data.length,
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