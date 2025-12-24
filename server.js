const express = require('express');
const cors = require('cors');
const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');

const app = express();
const PORT = 3000;

// STATIC EXCEL PATH
const EXCEL_PATH = 'D:\\demo\\data.xlsx';

// Website storage
const STORAGE_DIR = path.join(__dirname, 'storage');
const STORAGE_FILE = path.join(STORAGE_DIR, 'data.json');

// UNIQUE KEY COLUMN NAME (change this to match your CSV's unique identifier column)
// Common names: 'id', 'ID', 'Id', 'Code', 'ProductID', etc.
// If your CSV doesn't have a unique ID, we'll use the first column
const UNIQUE_KEY = 'id'; // Change this to your actual unique column name

app.use(cors());
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

function ensureStorageDir() {
  if (!fs.existsSync(STORAGE_DIR)) {
    fs.mkdirSync(STORAGE_DIR, { recursive: true });
  }
}

// Helper: Get unique key value from a row
function getUniqueKey(row, uniqueKeyColumn) {
  // Try the specified unique key column
  if (row[uniqueKeyColumn] !== undefined && row[uniqueKeyColumn] !== null) {
    return String(row[uniqueKeyColumn]);
  }
  // Fallback: use first column value
  const firstKey = Object.keys(row)[0];
  return String(row[firstKey] || '');
}

// Helper: Check if two rows are identical (all fields match)
function areRowsIdentical(row1, row2) {
  const keys1 = Object.keys(row1).sort();
  const keys2 = Object.keys(row2).sort();
  
  if (keys1.length !== keys2.length) return false;
  
  return keys1.every(key => {
    const val1 = String(row1[key] || '').trim();
    const val2 = String(row2[key] || '').trim();
    return val1 === val2;
  });
}

// POST: Import Excel from static path and merge with existing data
app.post('/api/import-excel', (req, res) => {
  try {
    if (!fs.existsSync(EXCEL_PATH)) {
      return res.status(404).json({ message: `Excel file not found at ${EXCEL_PATH}` });
    }

    // Read new data from CSV
    const workbook = XLSX.readFile(EXCEL_PATH);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const newData = XLSX.utils.sheet_to_json(worksheet, { defval: null });

    if (newData.length === 0) {
      return res.json({ message: 'CSV file is empty', rows: 0 });
    }

    // Detect unique key column from first row
    const firstRow = newData[0];
    let uniqueKeyColumn = UNIQUE_KEY;
    
    // Check if UNIQUE_KEY exists in the data, otherwise use first column
    if (!firstRow.hasOwnProperty(UNIQUE_KEY)) {
      uniqueKeyColumn = Object.keys(firstRow)[0];
      console.log(`Warning: Column "${UNIQUE_KEY}" not found. Using "${uniqueKeyColumn}" as unique key.`);
    }

    // Load existing stored data
    let existingData = [];
    if (fs.existsSync(STORAGE_FILE)) {
      try {
        const raw = fs.readFileSync(STORAGE_FILE, 'utf-8');
        existingData = JSON.parse(raw || '[]');
      } catch (err) {
        console.log('Could not read existing data, starting fresh:', err.message);
        existingData = [];
      }
    }

    // Create a map of existing data by unique key for fast lookup
    const existingMap = new Map();
    existingData.forEach(row => {
      const key = getUniqueKey(row, uniqueKeyColumn);
      if (key) {
        existingMap.set(key, row);
      }
    });

    // Merge: update existing, add new, track changes
    const mergedMap = new Map(existingMap);
    let added = 0;
    let updated = 0;
    let duplicatesRemoved = 0;

    newData.forEach(newRow => {
      const key = getUniqueKey(newRow, uniqueKeyColumn);
      
      if (!key) {
        // Skip rows without a valid unique key
        return;
      }

      if (mergedMap.has(key)) {
        const existingRow = mergedMap.get(key);
        
        // Check if it's an exact duplicate (all fields identical)
        if (areRowsIdentical(existingRow, newRow)) {
          duplicatesRemoved++;
          return; // Skip exact duplicate
        }
        
        // Update existing record with new values
        mergedMap.set(key, newRow);
        updated++;
      } else {
        // New record
        mergedMap.set(key, newRow);
        added++;
      }
    });

    // Convert map back to array
    const mergedData = Array.from(mergedMap.values());

    // Remove any remaining exact duplicates (by comparing all rows)
    const finalData = [];
    const seenRows = new Set();
    
    mergedData.forEach(row => {
      const rowSignature = JSON.stringify(row);
      if (!seenRows.has(rowSignature)) {
        seenRows.add(rowSignature);
        finalData.push(row);
      } else {
        duplicatesRemoved++;
      }
    });

    // Save merged data to storage
    ensureStorageDir();
    fs.writeFileSync(STORAGE_FILE, JSON.stringify(finalData, null, 2), 'utf-8');

    return res.json({
      message: 'Excel imported and merged successfully',
      totalRows: finalData.length,
      added: added,
      updated: updated,
      duplicatesRemoved: duplicatesRemoved,
      uniqueKeyColumn: uniqueKeyColumn
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
      return res.json([]);
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


// const express = require('express');
// const cors = require('cors');
// const fs = require('fs');
// const path = require('path');
// const XLSX = require('xlsx');

// const app = express();
// const PORT = 3000;

// // STATIC EXCEL PATH (change file name if needed)
// const EXCEL_PATH = 'D:\\demo\\data.xlsx';

// // "Website storage" = JSON file on server
// const STORAGE_DIR = path.join(__dirname, 'storage');
// const STORAGE_FILE = path.join(STORAGE_DIR, 'data.json');

// // Middlewares
// app.use(cors());
// app.use(express.json());

// // Serve frontend from /public
// app.use(express.static(path.join(__dirname, 'public')));

// // Helper: ensure storage folder exists
// function ensureStorageDir() {
//   if (!fs.existsSync(STORAGE_DIR)) {
//     fs.mkdirSync(STORAGE_DIR, { recursive: true });
//   }
// }

// // POST: Import Excel from static path and save to website storage (JSON file)
// app.post('/api/import-excel', (req, res) => {
//   try {
//     if (!fs.existsSync(EXCEL_PATH)) {
//       return res.status(404).json({ message: `Excel file not found at ${EXCEL_PATH}` });
//     }

//     const workbook = XLSX.readFile(EXCEL_PATH);
//     const sheetName = workbook.SheetNames[0];
//     const worksheet = workbook.Sheets[sheetName];

//     const data = XLSX.utils.sheet_to_json(worksheet, { defval: null });

//     ensureStorageDir();
//     fs.writeFileSync(STORAGE_FILE, JSON.stringify(data, null, 2), 'utf-8');

//     return res.json({
//       message: 'Excel imported and stored successfully',
//       rows: data.length,
//     });
//   } catch (err) {
//     console.error(err);
//     return res.status(500).json({ message: 'Error importing Excel', error: err.message });
//   }
// });

// // GET: Return stored data from website storage
// app.get('/api/data', (req, res) => {
//   try {
//     if (!fs.existsSync(STORAGE_FILE)) {
//       return res.json([]); // no data yet
//     }
//     const raw = fs.readFileSync(STORAGE_FILE, 'utf-8');
//     const data = JSON.parse(raw || '[]');
//     return res.json(data);
//   } catch (err) {
//     console.error(err);
//     return res.status(500).json({ message: 'Error reading stored data', error: err.message });
//   }
// });

// // Health check
// app.get('/api/health', (req, res) => {
//   res.json({ status: 'ok' });
// });

// app.listen(PORT, () => {
//   console.log(`Server listening on http://localhost:${PORT}`);
// }); -->