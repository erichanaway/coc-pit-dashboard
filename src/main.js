// =================================
// IMPORTS 
// =================================

import './style.css';
import * as XLSX from 'xlsx';
import Chart from 'chart.js/auto';

// =================================
// GLOBAL VARIABLES
// =================================

// Excel file used as the dashboard data source
const DATA_FILE = '/data/pit_2025_2026_dashboard_upload_data.xlsx';

// Stores all PIT rows after the Excel workbook is loaded
let pitData = [];

// Chart.js instances used throughout the dashboard
let compositionChart;
let countyDonutChart;
let sexChart;
let raceChart;
let ageChart;
let otherCategoriesChart;

// =================================
// INITIAL PAGE LAYOUT 
// =================================

document.querySelector('#app').innerHTML = `
  <div class="app-layout">

    <aside class="sidebar">
      <h2>CA-526</h2>
      <nav>
        <a class="active" href="#" data-page="overview">Overview</a>
        <a href="#" data-page="demographics">Demographics</a>
        <a href="#" data-page="other-categories">Other Categories</a>
        <a href="#" data-page="data-tables">Data Tables</a>
        <a href="#" data-page="about">About</a>
      </nav>
    </aside>

    <main class="dashboard">

      <div id="page-content">

        <h1>Central Sierra CoC PIT Dashboard</h1>

        <p class="subtitle">
          CA-526 Point-In-Time Count
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

          <div class="card kpi-card kpi-blue">
            <div class="kpi-content">
              <h3>Total Households</h3>
              <p id="total-households">--</p>
            </div>
          </div>

          <div class="card kpi-card kpi-orange">
            <div class="kpi-content">
              <h3>Total People</h3>
              <p id="total-people">--</p>
            </div>
          </div>

          <div class="card kpi-card kpi-blue">
            <div class="kpi-content">
              <h3>Sheltered Households</h3>
              <p id="sheltered-households">--</p>
            </div>
          </div>

          <div class="card kpi-card kpi-orange">
            <div class="kpi-content">
              <h3>Sheltered People</h3>
              <p id="sheltered-people">--</p>
            </div>
          </div>

          <div class="card kpi-card kpi-blue">
            <div class="kpi-content">
              <h3>Unsheltered Households</h3>
              <p id="unsheltered-households">--</p>
            </div>
          </div>

          <div class="card kpi-card kpi-orange">
            <div class="kpi-content">
              <h3>Unsheltered People</h3>
              <p id="unsheltered-people">--</p>
            </div>   
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
          Central Sierra CoC PIT Dashboard v1.5<br>
          Developed by Eric Hanaway
        </div>

      </div>
  
    </main>

  </div>
`;

// section 2

// =================================
// OVERVIEW PAGE EVENT LISTENERS
// =================================

document
  .querySelector('#year-select')
  .addEventListener('change', updateDashboard);

document
  .querySelector('#county-select')
  .addEventListener('change', updateDashboard);

// =================================
// OVERVIEW DASHBOARD
// =================================

function updateDashboard() {
  // Get the selected filter values from the overview dropdowns
  const selectedYear = document.querySelector('#year-select').value;
  const selectedCounty = document.querySelector('#county-select').value;

  // Filter the full PIT dataset to the selected year and county
  const filteredRows = pitData.filter(row =>
    row.year == selectedYear &&
    row.geography === selectedCounty
  );

  // DEBUG: Show demographic rows for the selected filters
  console.table(
    filteredRows.filter(row =>
      row.section === 'Demographics'
    )
  );

  // DEBUG: Show simplified row data for troubleshooting
  console.table(
    filteredRows.map(row => ({
      count_type: row.count_type,
      section: row.section,
      metric: row.metric,
      value: row.value
    }))
  );

  // Find total household and total people rows
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

  // Calculate sheltered people from ES and TH rows
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

  // Calculate unsheltered people
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

  // Calculate sheltered households from ES and TH rows
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

  // Calculate unsheltered households
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

  // Section 3

  // Update overview KPI cards
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

  // Destroy existing charts before redrawing them
  if (compositionChart) {
    compositionChart.destroy();
  }

  if (countyDonutChart) {
    countyDonutChart.destroy();
  }
  
  // Build county totals for the donut chart
  const countyNames = [
    'Amador',
    'Calaveras',
    'Mariposa',
    'Tuolumne'
  ];

  const countyTotals = countyNames.map(county => {
    const countyRows = pitData.filter(row => 
      row.year == selectedYear &&
      row.geography === county
    );

    const totalPeopleRow = countyRows.find(row =>
      row.section === 'Location and Family Type' &&
      row.metric === 'Total' &&
      row.count_type === 'People'
    );

    return Number(totalPeopleRow?.value || 0);
  });

  // Create sheltered vs unsheltered bar chart
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
          backgroundColor: '#0f3a6d',
          borderRadius: 6
        },
        {
          label: 'People',
          data: [
            shelteredPeople,
            unshelteredPeople
          ],
          backgroundColor: '#f97316',
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
  });

  // Create county distribution donut chart
  const donutCtx = document
    .getElementById('county-donut-chart')
    .getContext('2d');

  countyDonutChart = new Chart(donutCtx, {
    type: 'doughnut',
    data: {
      labels: countyNames,
      datasets: [
        {
          data: countyTotals,

          backgroundColor: [
            '#0f3a6d',
            '#f97316',
            '#94a3b8',
            '#cbd5e1'
          ],

          borderColor: '#ffffff',
          borderWidth: 2
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

function updateDemographics() {
  const demoYear = document.querySelector('#demo-year-select').value;
  const demoCounty = document.querySelector('#demo-county-select').value;
  const demoPopulation = document.querySelector('#demo-population-select').value;

  let demoRows = pitData.filter(row =>
    row.year == demoYear &&
    row.geography === demoCounty &&
    row.section === 'Age Groups'
  );

  if (demoPopulation !== 'All') {
    demoRows = demoRows.filter(row =>
      row.count_type === demoPopulation
    );
  }

  const childrenRows = demoRows.filter(row =>
  row.metric === 'Number of Children < 18'
);

const childrenTotal = childrenRows.reduce(
  (sum, row) => sum + Number(row.value || 0),
  0
);

document.querySelector('#demo-children').textContent = childrenTotal;

const youthRows = demoRows.filter(row =>
  row.metric === 'Number of Youth (18-24)'
);

const youthTotal = youthRows.reduce(
  (sum, row) => sum + Number(row.value || 0),
  0
);

document.querySelector('#demo-youth').textContent = youthTotal;

// Section 4

// Calculate older adults KPI
const seniorRows = demoRows.filter(row => 
  row.metric === 'Number of adults (65 or older)'
);

const seniorTotal = seniorRows.reduce(
  (sum, row) => sum + Number(row.value || 0),
  0
);

document.querySelector('#demo-seniors').textContent = seniorTotal;

// Calculate unaccompanied youth KPI
const unaccompaniedYouthRows = pitData.filter(row =>
  row.year == demoYear &&
  row.geography === demoCounty &&
  row.section === 'Unaccompanied Youth' &&
  row.metric === 'Youth 18-24' &&
  (
    demoPopulation === 'All' ||
    row.count_type === demoPopulation
  )
);

const unaccompaniedYouthTotal = unaccompaniedYouthRows.reduce(
  (sum, row) => sum + Number(row.value || 0),
  0
);

document.querySelector('#demo-unaccompanied').textContent =
  unaccompaniedYouthTotal;

// Calculate parenting youth KPI
const parentingYouthRows = pitData.filter(row =>
  row.year == demoYear &&
  row.geography === demoCounty &&
  row.section === 'Parenting Youth' &&
  row.metric === 'Parenting Youth 18-24' &&
  (
    demoPopulation === 'All' ||
    row.count_type === demoPopulation
  )
);

const parentingYouthTotal = parentingYouthRows.reduce(
  (sum, row) => sum + Number(row.value || 0),
  0
);

document.querySelector('#demo-parenting-youth').textContent =
parentingYouthTotal;

// Calculate children of parenting youth KPI
const parentingChildrenRows = pitData.filter(row =>
  row.year == demoYear &&
  row.geography === demoCounty &&
  row.section === 'Parenting Youth' &&
  row.metric === 'Children of Parenting youth' &&
  (
    demoPopulation === 'All' ||
    row.count_type === demoPopulation
  )
);

const parentingChildrenTotal = parentingChildrenRows.reduce(
  (sum, row) => sum + Number(row.value || 0),
  0
);

document.querySelector('#demo-parenting-children').textContent =
  parentingChildrenTotal;

// Build sex totals
const sexCategories = ['Male','Female','Unknown'];

const sexTotals = sexCategories.map(sex => {
  const rows = pitData.filter(row =>
    row.year == demoYear &&
    row.geography === demoCounty &&
    row.section === 'Sex' &&
    row.metric === sex &&
    (
      demoPopulation === 'All' ||
      row.count_type === demoPopulation
    )
  );

  return rows.reduce(
    (sum, row) => sum + Number(row.value || 0),
    0
  );
});

// Update sex chart legend totals
document.querySelector('#sex-male-total').textContent =
  `Male: ${sexTotals[0]}`;

document.querySelector('#sex-female-total').textContent =
   `Female: ${sexTotals[1]}`;

document.querySelector('#sex-unknown-total').textContent =
  `Unknown: ${sexTotals[2]}`;

// Create sex breakdown bar chart
const sexCtx = document
  .getElementById('sex-chart')
  .getContext('2d');

if (sexChart) {
  sexChart.destroy();
}

sexChart = new Chart(sexCtx, {
  type: 'doughnut',
  data: {
    labels: sexCategories,
    datasets: [
      {
        label: 'People',
        data: sexTotals,
        backgroundColor: [
          '#0f3a6d',
          '#f97316',
          '#94a3b8'
        ]
      }
    ]
  },
  options: {
    plugins: {
      legend: {
        position: 'bottom'
      },

      title: {
        display: false
      },

      tooltip: {
        enabled: true
     }
    },

  }
  
});

// Section 5
// Build race totals
let raceCategories = [
  'American Indian or Alaska Native or Indigenous',
  'American Indian or Alaska Native or Indigenous and Hispanic/Latina/e/o',
  'Asian or Asian American',
  'Black or African American or African',
  'Hispanic/Latina/o',
  'Native Hawaiian or Pacific Islander',
  'White',
  'White and Hispanic/Latina/o',
  'Multi-Racial and Hispanic/Latina/o',
  'Unknown'
];

let raceTotals = raceCategories.map(race => {

  const rows = pitData.filter(row =>
    row.year == demoYear &&
    row.geography === demoCounty &&
    row.section === 'Race' &&
    row.metric === race &&
    (
      demoPopulation === 'All' ||
      row.count_type === demoPopulation
    )
  );

  return rows.reduce(
    (sum, row) => sum + Number(row.value || 0),
    0
  );

});

const raceData = raceCategories
  .map((race, index) => ({
    race: race,
    total: raceTotals[index]
  }))
  .sort((a, b) => b.total - a.total);

raceCategories = raceData.map(item => item.race);
raceTotals = raceData.map(item => item.total);

// Create race breakdown bar chart
const raceCtx = document
  .getElementById('race-chart')
  .getContext('2d');

if (raceChart) {
  raceChart.destroy();
}

raceChart = new Chart(raceCtx, {
  type: 'bar',
  data: {

    labels: raceCategories,
    datasets: [
      {
        label: 'People',
        data: raceTotals,
        backgroundColor: '#0f3a6d',
        borderRadius: 6
      }
    ]
  },
  options: {
    indexAxis: 'y',

    plugins: {
      legend: {
        display: false
      },

      title: {
        display: false
      },

      tooltip: {
        enabled: true
      }
    },

    scales: {
      x: {
        type: 'logarithmic'
      }
    }
  }
});

// Build age totals
const ageCategories = [
  'Number of Children < 18',
  'Number of Youth (18-24)',
  'Number of adults (25-34)',
  'Number of adults (35-44)',
  'Number of adults (45-54)',
  'Number of adults (55-64)',
  'Number of adults (65 or older)'
];

const ageLabels = [
  'Children <18',
  'Youth 18-24',
  'Adults 25-34',
  'Adults 35-44',
  'Adults 45-54',
  'Adults 55-64',
  'Adults 65+'
];

const ageTotals = ageCategories.map(age => {

  const rows = pitData.filter(row =>
    row.year == demoYear &&
    row.geography === demoCounty &&
    row.section === 'Age Groups' &&
    row.metric === age &&
    (
      demoPopulation === 'All' ||
      row.count_type === demoPopulation
    )
  );

  return rows.reduce(
    (sum, row) => sum + Number(row.value || 0),
    0
  );
});

const ageCtx = document
  .getElementById('age-chart')
  .getContext('2d');

if (ageChart) {
  ageChart.destroy();
}

ageChart = new Chart(ageCtx, {
  type: 'bar',
  data: {
    labels: ageLabels,
    datasets: [
      {
        label: 'People',
        data: ageTotals,
        backgroundColor: '#0f3a6d',
        borderRadius: 10
      }
    ]
  },
  options: {
    plugins: {
      legend: {
        display: false
      },

      title: {
        display: false
      },

      tooltip: {
        enabled: true
      }
    },

    scales: {
      y: {
        beginAtZero: true
      }
    }
  }
});

// Section 6
// DEBUG: Confirm demographic rows found for selected filters
  console.log('Demo rows found:', demoRows.length);
}

//

function updateDataTables() {
    
  const tableYear =
      document.querySelector('#table-year-select').value;

  const tableCounty = 
    document.querySelector('#table-county-select').value;

  const tableRows = pitData.filter(row =>
      row.year == tableYear &&
      row.geography === tableCounty
  );

  const locationRows = tableRows.filter(row =>
    row.section === 'Location and Family Type'
  );

  const metrics = [
    'ES Houses w/children',
    'ES Adults Only',
    'TH Households w/children',
    'TH Adults Only',
    'Unsheltered w/children',
    'Unsheltered Adults Only',
    'Total'
  ];

  const tableBody = metrics.map(metric => {

    const householdRow = locationRows.find(row =>
      row.metric === metric &&
      row.count_type === 'Households'
    );

    const peopleRow = locationRows.find(row =>
      row.metric === metric &&
      row.count_type === 'People'
    );

    return `
      <tr>
        <td>${metric}</td>
        <td>${householdRow?.value ?? 0}</td>
        <td>${peopleRow?.value ?? 0}</td>
      </tr>
    `;
  }).join('');

  //

  const ageGroupRows = tableRows.filter(row =>
    row.section === 'Age Groups'
  );

  const ageMetrics = [
    'Number of Children < 18',
    'Number of Youth (18-24)',
    'Number of adults (25-34)',
    'Number of adults (35-44)',
    'Number of adults (45-54)',
    'Number of adults (55-64)',
    'Number of adults (65 or older)'
  ];

  const ageTableBody = ageMetrics.map(metric => {

    const shelteredRow = ageGroupRows.find(row =>
      row.metric === metric &&
      row.count_type === 'Sheltered'
    );

    const unshelteredRow = ageGroupRows.find(row =>
      row.metric === metric &&
      row.count_type === 'Unsheltered'
    );

    return `
      <tr>
        <td>${metric}</td>
        <td>${shelteredRow?.value ?? 0}</td>
        <td>${unshelteredRow?.value ?? 0}</td>
      </tr>
    `;
  }).join('');

  //

  const sexRows = tableRows.filter(row =>
    row.section ==='Sex'
  );

  const sexMetrics = [
    'Female',
    'Male',
    'Unknown'
  ];

  const sexTableBody = sexMetrics.map(metric => {

      const shelteredRow = sexRows.find(row =>
        row.metric === metric &&
        row.count_type === 'Sheltered'
      );

      const unshelteredRow = sexRows.find(row =>
        row.metric === metric &&
        row.count_type === 'Unsheltered'
      );

      return `
        <tr>
          <td>${metric}</td>
          <td>${shelteredRow?.value ?? 0}</td>
          <td>${unshelteredRow?.value ?? 0}</td>
        </tr>
      `;
    }).join('');

    const raceRows = tableRows.filter(row =>
      row.section === 'Race'
    );
    
    const raceMetrics = [
      'American Indian or Alaska Native or Indigenous',
      'American Indian or Alaska Native or Indigenous and Hispanic/Latina/e/o',
      'Asian or Asian American',
      'Black or African American or African',
      'Hispanic/Latina/o',
      'Native Hawaiian or Pacific Islander',
      'White',
      'White and Hispanic/Latina/o',
      'Multi-Racial and Hispanic/Latina/o',
      'Unknown'
    ];

    const raceTableBody = raceMetrics.map(metric => {

      const shelteredRow = raceRows.find(row =>
        row.metric === metric &&
        row.count_type === 'Sheltered'
      );

      const unshelteredRow = raceRows.find(row =>
        row.metric === metric &&
        row.count_type === 'Unsheltered'
      );

      return `
        <tr>
          <td>${metric}</td>
          <td>${shelteredRow?.value ?? 0}</td>
          <td>${unshelteredRow?.value ?? 0}</td>
        </tr>
      `;
    }).join('');

    const youthRows = tableRows.filter(row =>
      row.section === 'Unaccompanied Youth'
    );

    const youthMetrics = [
      'Children <18',
      'Youth 18-24'
    ];

    const youthTableBody = youthMetrics.map(metric => {

      const shelteredRow = youthRows.find(row =>
        row.metric === metric &&
        row.count_type === 'Sheltered'
      );

      const unshelteredRow = youthRows.find(row =>
        row.metric === metric &&
        row.count_type === 'Unsheltered'
      );

      return `
        <tr>
          <td>${metric}</td>
          <td>${shelteredRow?.value ?? 0}</td>
          <td>${unshelteredRow?.value ?? 0}</td>
        </tr>
      `;
    }).join('');

    const parentingRows = tableRows.filter(row =>
      row.section === 'Parenting Youth'
    );

    const parentingMetrics = [
      'Parenting Youth 18-24',
      'Children of Parenting youth'
    ];

    const parentingTableBody = parentingMetrics.map(metric => {

      const shelteredRow = parentingRows.find(row =>
        row.metric === metric &&
        row.count_type === 'Sheltered'
      );

      const unshelteredRow = parentingRows.find(row =>
        row.metric === metric &&
        row.count_type === 'Unsheltered'
      );

      return `
        <tr>
          <td>${metric}</td>
          <td>${shelteredRow?.value ?? 0}</td>
          <td>${unshelteredRow?.value ?? 0}</td>
        </tr>
      `;
    }).join('');

    const otherRows = tableRows.filter(row =>
      row.section === 'Other Categories'
    );

    const otherMetrics = [
      'Chronic Homeless Indiv.',
      'Veterans',
      'Chronic Homeless Vets.',
      'Mental Illness',
      'HIV/AIDS',
      'Fleeing Domestic Violence'
    ];

    const otherTableBody = otherMetrics.map(metric => {

      const shelteredRow = otherRows.find(row =>
        row.metric === metric &&
        row.count_type === 'Sheltered'
      );

      const unshelteredRow = otherRows.find(row =>
        row.metric === metric &&
        row.count_type === 'Unsheltered'
      );

      return `
        <tr>
          <td>${metric}</td>
          <td>${shelteredRow?.value ?? 0}</td>
          <td>${unshelteredRow?.value ?? 0}</td>
        </tr>
      `;
    }).join('');


  document.querySelector('#table-container').innerHTML = `
    <div class="chart-card">
      <h2>Location and Family Type</h2>

      <table class="data-table">
        <thead>
          <tr>
            <th>Metric</th>
            <th>Households</th>
            <th>People</th>
          </tr>
        </thead>

        <tbody>
          ${tableBody}
        </tbody>
      </table>
    </div>

    <div class="chart-card">
      <h2>Age Groups</h2>

      <table class="data-table">
        <thead>
          <tr>
            <th>Metric</th>
            <th>Sheltered</th>
            <th>Unsheltered</th>
          </tr>
        </thead>

        <tbody>
          ${ageTableBody}
        </tbody>
      </table>
    </div>

    <div class="chart-card">
      <h2>Sex</h2>

      <table class="data-table">
        <thead>
          <tr>
            <th>Metric</th>
            <th>Sheltered</th>
            <th>Unsheltered</th>
          </tr>
        </thead>

        <tbody>
          ${sexTableBody}
        </tbody>
      </table>
    </div>

    <div class="chart-card">
      <h2>Race</h2>

      <table class="data-table">
        <thead>
          <tr>
            <th>Metric</th>
            <th>Sheltered</th>
            <th>Unsheltered</th>
          </tr>
        </thead>

        <tbody>
          ${raceTableBody}
        </tbody>
      </table>
    </div>

    <div class="chart-card">
      <h2>Unaccompanied Youth</h2>

      <table class="data-table">
        <thead>
          <tr>
            <th>Metric</th>
            <th>Sheltered</th>
            <th>Unsheltered</th>
          </tr>
        </thead>

        <tbody>
          ${youthTableBody}
        </tbody>
      </table>
    </div>

        <div class="chart-card">
      <h2>Parenting Youth</h2>

      <table class="data-table">
        <thead>
          <tr>
            <th>Metric</th>
            <th>Sheltered</th>
            <th>Unsheltered</th>
          </tr>
        </thead>

        <tbody>
          ${parentingTableBody}
        </tbody>
      </table>
    </div>

    <div class="chart-card">
      <h2>Other Categories</h2>

      <table class="data-table">
        <thead>
          <tr>
            <th>Metric</th>
            <th>Sheltered</th>
            <th>Unsheltered</th>
          </tr>
        </thead>

        <tbody>
          ${otherTableBody}
        </tbody>
      </table>
    </div>
  `;
}


function updateOtherCategories() {
  
  const otherYear =
    document.querySelector('#other-year-select').value;

  const otherCounty =
    document.querySelector('#other-county-select').value;

  const otherPopulation = 
    document.querySelector('#other-population-select').value;

  let otherRows = pitData.filter(row =>
    row.year == otherYear &&
    row.geography === otherCounty &&
    row.section === 'Other Categories'
  );

  if (otherPopulation !== "All") {
    otherRows = otherRows.filter(row =>
      row.count_type === otherPopulation
    );
  }
  
  console.log('Other rows found:', otherRows.length);

  const chronicIndividualsRows = otherRows.filter(row =>
    row.metric === 'Chronic Homeless Indiv.'
  );

  const chronicIndividualsTotal = chronicIndividualsRows.reduce(
    (sum, row) => sum + Number(row.value || 0),
    0
  );

  document.querySelector('#other-chronic-individuals').textContent =
    chronicIndividualsTotal;

  //

  const veteranRows = otherRows.filter(row =>
    row.metric === 'Veterans'
  );

  const veteranTotal = veteranRows.reduce(
    (sum, row) => sum + Number(row.value || 0),
    0
  );

  document.querySelector('#other-veterans').textContent =
    veteranTotal;

  //

  const chronicVeteranRows = otherRows.filter(row =>
    row.metric === 'Chronic Homeless Vets.'
  );

  const chronicVeteranTotal = chronicVeteranRows.reduce(
    (sum, row) => sum + Number(row.value || 0),
    0
  );

  document.querySelector('#other-chronic-veterans').textContent =
    chronicVeteranTotal;
  //

  const mentalIllnessRows = otherRows.filter(row =>
    row.metric === 'Mental Illness'
  );

  const mentalIllnessTotal = mentalIllnessRows.reduce(
    (sum, row) => sum + Number(row.value || 0),
    0
  );

  document.querySelector('#other-mental-illness').textContent =
    mentalIllnessTotal;
  //

  const hivRows = otherRows.filter(row =>
    row.metric === 'HIV/AIDS'
  );

  const hivTotal = hivRows.reduce(
    (sum, row) => sum + Number(row.value || 0),
    0
  );

  document.querySelector('#other-hiv').textContent =
    hivTotal;

  //

  const dvRows = otherRows.filter(row =>
    row.metric === 'Fleeing Domestic Violence'
  );

  const dvTotal = dvRows.reduce(
    (sum, row) => sum + Number(row.value || 0),
    0
  );

  document.querySelector('#other-dv').textContent =
    dvTotal;

  //

  const otherCategoryLabels = [
    'Chronically Homeless Individuals',
    'Veterans',
    'Chronically Homeless Veterans',
    'Mental Illness',
    'HIV/AIDS',
    'Fleeing DV'
  ];

  const otherCategoryTotals = [
    chronicIndividualsTotal,
    veteranTotal,
    chronicVeteranTotal,
    mentalIllnessTotal,
    hivTotal,
    dvTotal
  ];

  const otherCtx = document
    .getElementById('other-categories-chart')
    .getContext('2d');

  if (otherCategoriesChart) {
    otherCategoriesChart.destroy();
  }

  otherCategoriesChart = new Chart(otherCtx, {
    type: 'bar',

    data: {
      labels: otherCategoryLabels,

      datasets: [
        {
          label: 'People',
          data: otherCategoryTotals,
          backgroundColor: [
            '#0f3a6d',
            '#f97316',
            '#0f3a6d',
            '#f97316',
            '#0f3a6d',
            '#f97316'
          ],
          borderRadius: 10
        }
      ]
    },

    options: {
      indexAxis: 'y',

      plugins: {
        legend: {
          display: false
        },

        title: {
          display: false,
        }
      },

      scales: {
        x: {
          beginAtZero: true
        }
      }
    }
  });
  ///
}

// =================================
// LOAD EXCEL WORKBOOK
// =================================

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

// =================================
// PAGE NAVIGATION
// =================================

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

    if (page === 'about') {
      document.querySelector('#page-content').innerHTML = `
        <h1>About</h1>

        <div class="about-card">
          <h2>What is the Point-in-Time Count?</h2>

          <p>
           The Point-in-Time (PIT) Count is an annual count of people
           experiencing homelessness conducted during the last ten days
           of January. Required by HUD, the PIT Count provides a
           snapshot of homelessness within a community on a single night.
        </p>
      </div>

      <div class="about-card">
        <h2>About CA-526</h2>

        <p>
          The Central Sierra Continuum of Care serves four rural
          counties in California:
        </p>

        <ul>
          <li>Amador County</li>
          <li>Calaveras County</li>
          <li>Mariposa County</li>
          <li>Tuolumne County</li>
        </ul>
       </div>

       <div class="about-card">
        <h2>About This Dashboard</h2>

        <p>
          This dashboard was developed to provide an interactive
          view of Point-in-Time Count results across the
          Central Sierra Continuum of Care.
        </p>

        <p>
          Users can explore county-level results, compare
          sheltered and unsheltered populations, and review
          demographic information collected during the PIT Count.
        </p>
      </div>

      <div class="about-card">
        <h2>Data Sources</h2>

        <p>
          Data presented in this dashboard is compiled from
          multiple sources used to support the annual Point-in-Time
          Count conducted by the Central Sierra Continuum of Care.
        </p>

        <ul>
          <li>Homeless Management Information System (HMIS) records</li>
          <li>Unsheltered Point-in-Time Count survey responses</li>
          <li>Sheltered Point-in-Time Count reporting submissions</li>
          <li>HUD Point-in-Time Count reporting requirements and methodologies</li>
        </ul>
      </div>

      <div class="about-card">
        <h2>Dashboard Information</h2>

        <p><strong>Version:</strong> 1.4</p>

        <p><strong>Developed By:</strong> Eric Hanaway</p>

        <p>
          HMIS Manager<br>
          Central Sierra Continuum of Care (CA-526)
        </p>
      </div>
    `;

    return;
  }

  if (page === 'demographics') {
    document.querySelector('#page-content').innerHTML = `
      <h1>Demographics</h1>

      <p class='subtitle'>
        Demographics breakdown of people counted during the PIT count.
      </p>

      <div class="controls">
        <label>
          Year
          <select id="demo-year-select">
            <option value="2026">2026</option>
            <option value="2025">2025</option>
          </select>
        </label>

        <label>
          County
          <select id="demo-county-select">
            <option value="Combined">Combined</option>
            <option value="Amador">Amador</option>
            <option value="Calaveras">Calaveras</option>
            <option value="Mariposa">Mariposa</option>
            <option value="Tuolumne">Tuolumne</option>
          </select>
        </label>

        <label> 
          Population
            <select id="demo-population-select">
              <option value="All">All</option>
              <option value="Sheltered">Sheltered</option>
              <option value="Unsheltered">Unsheltered</option>
            </select>
          </label>
      </div>

      <div class="kpi-grid">
      
        <div class="card kpi-card kpi-blue">
          <div class="kpi-content">
            <h3>Children (&lt;18)</h3>
            <p id="demo-children">--</p>
          </div>
        </div>

        <div class="card kpi-card kpi-orange">
          <div class="kpi-content">
            <h3>Youth (18-24)</h3>
            <p id="demo-youth">--</p>
          </div>
        </div>

        <div class="card kpi-card kpi-blue">
          <div class="kpi-content">
            <h3>Older Adults (65+)</h3>
            <p id="demo-seniors">--</p>
          </div>
        </div>

        <div class="card kpi-card kpi-orange">
          <div class="kpi-content">
            <h3>Unaccompanied Youth</h3>
            <p id="demo-unaccompanied">--</p>
          </div>
        </div>

        <div class="card kpi-card kpi-blue">
          <div class="kpi-content">
            <h3>Parenting Youth</h3>
            <p id="demo-parenting-youth">--</p>
          </div>
        </div>

        <div class="card kpi-card kpi-orange">
          <div class="kpi-content">
            <h3>Children of Parenting Youth</h3>
            <p id="demo-parenting-children">--</p>
          </div>
        </div>

      </div>

      <div class="demographics-charts-grid">

        <div class="chart-card">
          <h2>Sex Breakdown</h2>

          <canvas id="sex-chart"></canvas>

          <div class="chart-legend">
            <p id="sex-male-total"></p>
            <p id="sex-female-total"></p>
            <p id="sex-unknown-total"></p>
          </div>
        </div>

        <div class="chart-card">
          <h2>Race Breakdown</h2>

          <canvas id="race-chart"></canvas>
        </div>

        <div class="chart-card full-width-chart">
          <h2>Age Distribution</h2>

          <canvas id="age-chart"></canvas>
        </div>

      </div>

    `;

  document
    .querySelector('#demo-year-select')
    .addEventListener('change', updateDemographics);


  document
    .querySelector('#demo-county-select')
    .addEventListener('change', updateDemographics);

  document
    .querySelector('#demo-population-select')
    .addEventListener('change', updateDemographics);

    updateDemographics();

    return;

  }

  if (page === 'data-tables') {
  document.querySelector('#page-content').innerHTML = `
    <h1>Data Tables</h1>

    <p class="subtitle">
      Explore detailed PIT count tables.
    </p>

    <div class="controls">
      <label>
        Year
        <select id="table-year-select">
          <option value="2026">2026</option>
          <option value="2025">2025</option>
        </select>
      </label>

      <label>
        County
        <select id="table-county-select">
          <option value="Combined">Combined</option>
          <option value="Amador">Amador</option>
          <option value="Calaveras">Calaveras</option>
          <option value="Mariposa">Mariposa</option>
          <option value="Tuolumne">Tuolumne</option>
        </select>
      </label>

    </div>

    <div id="table-container">
      <div class="coming-soon-card">
        <h2>Coming Soon</h2>

        <p>
          Interactive data tables will be available in Version 1.0.
        </p>
      </div>
    </div>
  `;

  document
    .querySelector('#table-year-select')
    .addEventListener('change', updateDataTables);

  document
    .querySelector('#table-county-select')
    .addEventListener('change', updateDataTables);

  updateDataTables();

  return;
}

  if (page === 'other-categories') {
    
    document.querySelector('#page-content').innerHTML =`
      <h1>Other Categories</h1>

      <p class="subtitle">
        Other populations and special populations identified in the PIT count.
      </p>

      <div class="controls">
        <label>
          Year
          <select id="other-year-select">
            <option value="2026">2026</option>
            <option value="2025">2025</option>
          </select>
        </label>

        <label>
          County
          <select id="other-county-select">
            <option value="Combined">Combined</option>
            <option value="Amador">Amador</option>
            <option value="Calaveras">Calaveras</option>
            <option value="Mariposa">Mariposa</option>
            <option value="Tuolumne">Tuolumne</option>
          </select>
        </label>

        <label>
          Population
          <select id="other-population-select">
            <option value="All">All</option>
            <option value="Sheltered">Sheltered</option>
            <option value="Unsheltered">Unsheltered</option>
          </select>
        </label>

      </div>

      <div class="kpi-grid">

        <div class="card kpi-card kpi-blue">
          <div class="kpi-content">
            <h3>Chronic Homeless Individuals</h3>
            <p id="other-chronic-individuals">0</p>
          </div>
        </div>

        <div class="card kpi-card kpi-orange">
          <div class="kpi-content">
            <h3>Veterans</h3>
            <p id="other-veterans">0</p>
          </div>
        </div>

        <div class="card kpi-card kpi-blue">
          <div class="kpi-content">
            <h3>Chronic Homeless Veterans</h3>
            <p id="other-chronic-veterans">0</p>
          </div>
        </div>

        <div class="card kpi-card kpi-orange">
          <div class="kpi-content">
            <h3>Mental Illness</h3>
            <p id="other-mental-illness">0</p>
          </div>
        </div>

        <div class="card kpi-card kpi-blue">
          <div class="kpi-content">
            <h3>HIV/AIDS</h3>
            <p id="other-hiv">0</p>
          </div>
        </div>  

        <div class="card kpi-card kpi-orange">
          <div class="kpi-content">
            <h3>Fleeing Domestic Violence</h3>
            <p id="other-dv">0</p>
          </div>
        </div>

      </div>

      <div class="chart-card">
        <h2>Other Categories Breakdown</h2>
        <canvas id="other-categories-chart"></canvas>
      </div>
        
    `;

    document
      .querySelector('#other-year-select')
      .addEventListener('change', updateOtherCategories);

    document
      .querySelector('#other-county-select')
      .addEventListener('change', updateOtherCategories);

    document
      .querySelector('#other-population-select')
      .addEventListener('change', updateOtherCategories);

    updateOtherCategories();


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