import './style.css';
import * as XLSX from 'xlsx';
import Chart from 'chart.js/auto';

const DATA_FILE = '/data/pit_2026_dashboard_upload_data.xlsx';

let pitData = [];

document.querySelector('#app').innerHTML = `
  <div class="dashboard">
    <h1>Central Sierra CoC PIT Dashboard</h1>

    <p class="subtitle">
      2026 Point In Time Count
    </p>

    <div class="controls">

      <select id="county-select">
        <option value="Combined">Combined</option>
        <option value="Amador">Amador</option>
        <option value="Calaveras">Calaveras</option>
        <option value="Mariposa">Mariposa</option>
        <option value="Tuolumne">Tuolumne</option>
      </select>

    </div>

    <div class="card">
      <p id="status">Loading Excel workbook...</p>
    </div>
  </div>
`;

document
  .querySelector('#county-select')
  .addEventListener('change', (event) => {

    const selectedCounty = event.target.value;

    const filteredRows = pitData.filter(row => 
      row.geography === selectedCounty
    );
    
    console.log(filteredRows);

    document.querySelector('#status').textContent =
      `${selectedCounty} selected. Geography rows found: ${filteredRows.length}`;
  });


async function loadWorkbook() {
  const response = await fetch(DATA_FILE);
  const arrayBuffer = await response.arrayBuffer();

  const workbook = XLSX.read(arrayBuffer, { type: 'array' });

  console.log('Workbook sheets:', workbook.SheetNames);

  const firstSheetName = 'data_upload';
  const worksheet = workbook.Sheets[firstSheetName];

  const rows = XLSX.utils.sheet_to_json(worksheet);

  pitData = rows;

  console.log('First sheet:', firstSheetName);
  console.log('Rows:', rows);

  document.querySelector('#status').textContent =
    `Loaded workbook. First sheet: ${firstSheetName}. Rows found: ${rows.length}`;
}

loadWorkbook();