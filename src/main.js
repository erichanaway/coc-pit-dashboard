import './style.css';
import * as XLSX from 'xlsx';
import Chart from 'chart.js/auto';

const DATA_FILE = '/data/pit_2025_2026_dashboard_upload_data.xlsx';

let pitData = [];
let compositionChart;
let countyDonutChart;
let sexChart;
let raceChart;

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
    filteredRows.filter(row =>
      row.section === 'Demographics'
    )
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

const seniorRows = demoRows.filter(row => 
  row.metric === 'Number of adults (65 or older)'
);

const seniorTotal = seniorRows.reduce(
  (sum, row) => sum + Number(row.value || 0),
  0
);

document.querySelector('#demo-seniors').textContent = seniorTotal;

// Unacc Youth

const unaccompaniedYouthRows = pitData.filter(row =>
  row.year == demoYear &&
  row.geography ===demoCounty &&
  row.section === 'Unaccompanied Youth' &&
  row.metric === 'Youth 18-24'
);

const unaccompaniedYouthTotal = unaccompaniedYouthRows.reduce(
  (sum, row) => sum + Number(row.value || 0),
  0
);

document.querySelector('#demo-unaccompanied').textContent =
unaccompaniedYouthTotal;

// Parenting Youth

const parentingYouthRows = pitData.filter(row =>
  row.year == demoYear &&
  row.geography === demoCounty &&
  row.section === 'Parenting Youth' &&
  row.metric === 'Parenting Youth 18-24'
);

const parentingYouthTotal = parentingYouthRows.reduce(
  (sum, row) => sum + Number(row.value || 0),
  0
);

document.querySelector('#demo-parenting-youth').textContent = 
parentingYouthTotal;

// Children of Parenting Youth

const parentingChildrenRows = pitData.filter(row =>
  row.year == demoYear &&
  row.geography === demoCounty &&
  row.section === 'Parenting Youth' &&
  row.metric === 'Children of Parenting youth'
);

const parentingChildrenTotal = parentingChildrenRows.reduce(
  (sum, row) => sum + Number(row.value || 0),
  0
);

document.querySelector('#demo-parenting-children').textContent =
  parentingChildrenTotal;


const sexCategories = ['Male','Female', 'Unknown'];

const sexTotals = sexCategories.map(sex => {

  const rows = pitData.filter(row =>
    row.year == demoYear &&
    row.geography === demoCounty &&
    row.section === 'Sex' &&
    row.metric === sex
  );

  return rows.reduce(
    (sum, row) => sum + Number(row.value || 0),
    0
  );
});

document.querySelector('#sex-male-total').textContent =
  `Male: ${sexTotals[0]}`;

document.querySelector('#sex-female-total').textContent =
   `Female: ${sexTotals[1]}`;

document.querySelector('#sex-unknown-total').textContent =
  `Unknown: ${sexTotals[2]}`;


const sexCtx = document
  .getElementById('sex-chart')
  .getContext('2d');

if (sexChart) {
  sexChart.destroy();
}

sexChart = new Chart(sexCtx, {
  type: 'bar',
  data: {
    labels: sexCategories,
    datasets: [
      {
        label: 'People',
        data: sexTotals,
        backgroundColor: '#4e79a7',
        borderRadius: 6
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


const raceCategories = [
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

const raceTotals = raceCategories.map(race => {

  const rows = pitData.filter(row =>
    row.year == demoYear &&
    row.geography === demoCounty &&
    row.section === 'Race' &&
    row.metric === race
  );

  return rows.reduce(
    (sum, row) => sum + Number(row.value || 0),
    0
  );

});

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
        backgroundColor: '#4e79a7',
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
        beginAtZero: true
      }
    }
  }
});


  console.log('Demo rows found:', demoRows.length);
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

            <p><strong>Version:</strong> 1.12</p>

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
        <div class="card">
          <h3>Children (&lt;18)</h3>
          <p id="demo-children">--</p>
        </div>

        <div class="card">
          <h3>Youth (18-24)</h3>
          <p id="demo-youth">--</p>
        </div>

        <div class="card">
          <h3>Older Adults (65+)</h3>
          <p id="demo-seniors">--</p>
        </div>

        <div class="card">
          <h3>Unaccompanied Youth</h3>
          <p id="demo-unaccompanied">--</p>
        </div>

        <div class="card">
          <h3>Parenting Youth</h3>
          <p id="demo-parenting-youth">--</p>
        </div>

        <div class="card">
          <h3>Children of Parenting Youth</h3>
          <p id="demo-parenting-children">--</p>
        </div>
      
      </div>

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