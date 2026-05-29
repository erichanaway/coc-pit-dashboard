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
      <h3>Total Households</h3>
      <p id="total-households">--</p>

      <h3>Total People</h3>
      <p id="total-people">--</p>
    </div>
  </div>
`;

document
  .querySelector('#county-select')
  .addEventListener('change', (event) => {

    updateDashboard(event.target.value);

  });

  function updateDashboard(selectedCounty) {
    const filteredRows = pitData.filter(row =>
      row.geography ===selectedCounty
    );

    console.table(
      filteredRows.map(row => ({
        count_type: row.count_type,
        section: row.section,
        metric: row.metric,
        value: row.value
      }))
    );

    const totalHouseholds = filteredRows.find(row =>
  row.section === 'Location and Family Type' &&
  row.metric === 'Total' &&
  row.count_type === 'Households'
);

const totalPeople = filteredRows.find(row =>
  row.section === 'Location and Family Type' &&
  row.metric === 'Total' &&
  row.count_type === 'People'
);

document.querySelector('#total-households').textContent =
  totalHouseholds?.value ?? '--';

document.querySelector('#total-people').textContent =
  totalPeople?.value ?? '--';
  
  }

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

  updateDashboard('Combined');

  
}

loadWorkbook();