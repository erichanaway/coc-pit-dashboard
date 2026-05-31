import './style.css';
import * as XLSX from 'xlsx';
import Chart from 'chart.js/auto';

const DATA_FILE = '/data/pit_2025_2026_dashboard_upload_data.xlsx';

let pitData = [];
let compositionChart;
let countyDonutChart;

document.querySelector('#app').innerHTML = `
  <div class="app-layout">

    <aside class="sidebar">
      <h2>CA-526</h2>
      <nav>
        <a class="active" href="#" data-page="overview">Overview</a>
        <a href="#" data-page="demographics">Demographics</a>
        <a href="#" data-page="veterans">Veterans</a>
        <a href="#" data-page="chronic">Chronic</a>
        <a href="#" data-page="household-types">Household Types</a>
        <a href="#" data-page="county-comparison">County Comparison</a>
        <a href="#" data-page="about">About</a>
      </nav>
    </aside>

    <main class="dashboard">

      <div id="page-content">

    <h1>Central Sierra CoC PIT Dashboard</h1>

    <p class="subtitle">
      2026 Point In Time Count
    </p>

    <div class="controls">
      <label>
        Year
        <select id="year-select">
          <option value="2026">2026</option>
          <option value="2025">2025</option>
        </select>
      </label>

      <label>
        County
        <select id="county-select">
          <option value="Combined">Combined</option>
          <option value="Amador">Amador</option>
          <option value="Calaveras">Calaveras</option>
          <option value="Mariposa">Mariposa</option>
          <option value="Tuolumne">Tuolumne</option>
        </select>
      </label>
    </div>

    <div class="kpi-grid">
      <div class="card">
        <h3>Total Households</h3>
        <p id="total-households">--</p>
      </div>

      <div class="card">
        <h3>Total People</h3>
        <p id="total-people">--</p>
      </div>

      <div class="card">
        <h3>Sheltered Households</h3>
        <p id="sheltered-households">--</p>
      </div>

      <div class="card">
        <h3>Sheltered People</h3>
        <p id="sheltered-people">--</p>
      </div>

      <div class="card">
        <h3>Unsheltered Households</h3>
        <p id="unsheltered-households">--</p>
      </div>

      <div class="card">
        <h3>Unsheltered People</h3>
        <p id="unsheltered-people">--</p>
      </div>

    </div>

    <div class="charts-grid">

     <div class="chart-card">
       <canvas id="composition-chart"></canvas>
     </div>

     <div class="chart-card">
       <canvas id="county-donut-chart"></canvas>
     </div>
    </div>

    <div class="footer">
     Central Sierra CoC PIT Dashboard v1.12<br>
     Developed by Eric Hanaway
    </div>

    </div>
  
  </main>

  </div>
`;

document
  .querySelector('#year-select')
  .addEventListener('change', updateDashboard);

document
  .querySelector('#county-select')
  .addEventListener('change', updateDashboard);

function updateDashboard() {
  const selectedYear = document.querySelector('#year-select').value;
  const selectedCounty = document.querySelector('#county-select').value;

  const filteredRows = pitData.filter(row =>
    row.year == selectedYear &&
    row.geography === selectedCounty
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

  const shelteredPeopleRows = filteredRows.filter(row =>
    row.section === 'Location and Family Type' &&
    row.count_type === 'People' &&
    (
      row.metric === 'ES Adults Only' ||
      row.metric === 'ES Houses w/children' ||
      row.metric === 'TH Adults Only' ||
      row.metric === 'TH Households w/children'
    )
  );

  const shelteredPeople = shelteredPeopleRows.reduce(
    (sum, row) => sum + Number(row.value || 0),
    0
  );

  const unshelteredPeopleRows = filteredRows.filter(row =>
    row.section === 'Location and Family Type' &&
    row.count_type === 'People' &&
    (
      row.metric === 'Unsheltered Adults Only' ||
      row.metric === 'Unsheltered w/children'
    )
  );

  const unshelteredPeople = unshelteredPeopleRows.reduce(
    (sum, row) => sum + Number(row.value || 0),
    0
  );

  const shelteredHouseholdsRows = filteredRows.filter(row =>
    row.section === 'Location and Family Type' &&
    row.count_type === 'Households' &&
    (
      row.metric === 'ES Houses w/children' ||
      row.metric === 'ES Adults Only' ||
      row.metric === 'TH Households w/children' ||
      row.metric === 'TH Adults Only'
    )
  );

  const shelteredHouseholds = shelteredHouseholdsRows.reduce(
    (sum, row) => sum + Number(row.value || 0),
    0
  );

  const unshelteredHouseholdsRows = filteredRows.filter(row =>
    row.section === 'Location and Family Type' &&
    row.count_type === 'Households' &&
   (
     row.metric === 'Unsheltered Adults Only' ||
     row.metric === 'Unsheltered w/children'
   )
  );

  const unshelteredHouseholds = unshelteredHouseholdsRows.reduce(
    (sum, row) => sum + Number(row.value || 0),
    0
  );

  document.querySelector('#total-households').textContent =
    totalHouseholds?.value ?? '--';

  document.querySelector('#total-people').textContent =
    totalPeople?.value ?? '--';

  document.querySelector('#sheltered-people').textContent =
    shelteredPeople;

  document.querySelector('#unsheltered-people').textContent = 
    unshelteredPeople;

  document.querySelector('#sheltered-households').textContent =
    shelteredHouseholds;

  document.querySelector('#unsheltered-households').textContent =
    unshelteredHouseholds;

  if (compositionChart) {
    compositionChart.destroy();
  }

  if (countyDonutChart) {
    countyDonutChart.destroy();
  }
  
  const countyNames = [
    'Amador',
    'Calaveras',
    'Mariposa',
    'Tuolumne'
  ];

  const countyTotals = countyNames.map(county => {

    const countyRows = pitData.filter(row => 
      row.year == selectedYear &&
      row.geography ===county
    );

    const totalPeopleRow = countyRows.find(row =>
      row.section === 'Location and Family Type' &&
      row.metric === 'Total' &&
      row.count_type === 'People'
    );

    return Number(totalPeopleRow?.value || 0);
  });

  const ctx = document
  .getElementById('composition-chart')
  .getContext('2d');

  compositionChart = new Chart(ctx, {
    type: 'bar',
    data: {
      labels: ['Sheltered', 'Unsheltered'],
      datasets: [
        {
          label: 'Households',
          data: [
            shelteredHouseholds,
            unshelteredHouseholds
          ],
          backgroundColor: '#4e79a7',
          borderRadius: 6
        },
        {
          label: 'People',
          data: [
            shelteredPeople,
            unshelteredPeople
          ],
          backgroundColor: '#f28e2b',
          borderRadius: 6
        }
      ]
    },

    options: {
      plugins: {
        title: {
          display: true,
          text: "Sheltered vs Unsheltered"
        }
      }
    }
  })


const donutCtx = document
  .getElementById('county-donut-chart')
  .getContext('2d');

countyDonutChart = new Chart(donutCtx, {
  type: 'doughnut',
  data: {
    labels: ['Amador', 'Calaveras', 'Mariposa', 'Tuolumne'],
    datasets: [
      {
        data: countyTotals,

        backgroundColor: [
          '#4e79a7', // Amador
          '#f28e2b', // Calaveras
          '#59a14f', // Mariposa
          '#e15759', // Tuolumne
        ],

        borderColor: '#ffffff',
        borderWidth:2
      }
    ]
  },

  options: {
    plugins: {
      title: {
        display: true,
        text: "County Share of Total Unhoused People"
      }
    }
  }
});
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

  updateDashboard();
}

document.querySelectorAll('.sidebar a').forEach(link => {

  link.addEventListener('click', (event) => {

    event.preventDefault();

    const page = link.dataset.page;

    document
      .querySelectorAll('.sidebar a')
      .forEach(a => a.classList.remove('active'));

    link.classList.add('active');

    if (page === 'overview') {
      location.reload();
      return;
    }

    document.querySelector('#page-content').innerHTML = `
      <h1>${link.textContent}</h1>

      <div class="coming-soon-card">
        <h2>Coming Soon</h2>

        <p>
          This section is currently under development.
        </p>

        <p>
          Planned for a future dashboard release.
        </p>
      </div>
    `;
  });

});

loadWorkbook();