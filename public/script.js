const API_BASE = ''; // same origin (http://localhost:3000)

const importBtn = document.getElementById('importBtn');
const loadDataBtn = document.getElementById('loadDataBtn');
const statusEl = document.getElementById('status');
const tableHead = document.getElementById('tableHead');
const tableBody = document.getElementById('tableBody');

importBtn.addEventListener('click', async () => {
  setStatus('Importing Excel from D:\\demo ...');
  try {
    const res = await fetch(`${API_BASE}/api/import-excel`, { method: 'POST' });
    const data = await res.json();
    if (!res.ok) {
      throw new Error(data.message || 'Failed to import');
    }
    setStatus(`${data.message} (rows: ${data.rows})`);
  } catch (err) {
    setStatus('Error: ' + err.message, true);
  }
});

loadDataBtn.addEventListener('click', async () => {
  setStatus('Loading stored data...');
  try {
    const res = await fetch(`${API_BASE}/api/data`);
    const rows = await res.json();
    setStatus(`Loaded ${rows.length} rows.`);
    renderTable(rows);
  } catch (err) {
    setStatus('Error: ' + err.message, true);
  }
});

function setStatus(message, isError = false) {
  statusEl.textContent = message;
  statusEl.style.color = isError ? 'red' : '#333';
}

function renderTable(rows) {
  tableHead.innerHTML = '';
  tableBody.innerHTML = '';

  if (!rows || rows.length === 0) return;

  const columns = Object.keys(rows[0]);

  // Header
  const headerRow = document.createElement('tr');
  columns.forEach(col => {
    const th = document.createElement('th');
    th.textContent = col;
    headerRow.appendChild(th);
  });
  tableHead.appendChild(headerRow);

  // Body
  rows.forEach(row => {
    const tr = document.createElement('tr');
    columns.forEach(col => {
      const td = document.createElement('td');
      td.textContent = row[col];
      tr.appendChild(td);
    });
    tableBody.appendChild(tr);
  });
}