const fileInput = document.getElementById('fileInput');
const thresholdInput = document.getElementById('thresholdInput');
const sampleBtn = document.getElementById('sampleBtn');
const tableBody = document.querySelector('#attendanceTable tbody');

const totalCountEl = document.getElementById('totalCount');
const presentCountEl = document.getElementById('presentCount');
const absentCountEl = document.getElementById('absentCount');

let chart;
let records = [];

function parseHours(value) {
  if (typeof value === 'number') {
    return value <= 1 ? value * 24 : value;
  }

  if (typeof value !== 'string') return 0;

  const clean = value.trim();
  if (!clean) return 0;

  if (/^\d+(\.\d+)?$/.test(clean)) return Number(clean);

  const parts = clean.split(':').map(Number);
  if (parts.some(Number.isNaN)) return 0;

  const [hours = 0, mins = 0, secs = 0] = parts;
  return hours + mins / 60 + secs / 3600;
}

function normalizeRow(row) {
  const employeeId = row['employee id'] ?? row.employeeId ?? row['Employee ID'] ?? row.EmpID ?? '';
  const name = row.name ?? row.Name ?? row['employee name'] ?? '';
  const punchIn = row['punch in'] ?? row.PunchIn ?? row['Punch In'] ?? '';
  const punchOut = row['punch out'] ?? row.PunchOut ?? row['Punch Out'] ?? '';
  const total = row['total time'] ?? row.TotalTime ?? row['Total Time'] ?? row['total hours'] ?? '';

  return {
    employeeId,
    name,
    punchIn,
    punchOut,
    totalHours: parseHours(total),
  };
}

function render() {
  const threshold = Number(thresholdInput.value) || 8;
  const evaluated = records.map((r) => ({ ...r, status: r.totalHours >= threshold ? 'Present' : 'Absent' }));
  const present = evaluated.filter((r) => r.status === 'Present').length;
  const absent = evaluated.length - present;

  totalCountEl.textContent = evaluated.length;
  presentCountEl.textContent = present;
  absentCountEl.textContent = absent;

  tableBody.innerHTML = '';

  evaluated.forEach((r) => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td>${r.employeeId}</td>
      <td>${r.name}</td>
      <td>${r.punchIn}</td>
      <td>${r.punchOut}</td>
      <td>${r.totalHours.toFixed(2)}</td>
      <td class="${r.status === 'Present' ? 'status-present' : 'status-absent'}">${r.status}</td>
    `;
    tableBody.appendChild(tr);
  });

  const ctx = document.getElementById('attendanceChart');
  if (chart) chart.destroy();
  chart = new Chart(ctx, {
    type: 'doughnut',
    data: {
      labels: ['Present', 'Absent'],
      datasets: [
        {
          data: [present, absent],
          backgroundColor: ['#16a34a', '#dc2626'],
        },
      ],
    },
    options: {
      plugins: {
        legend: {
          position: 'bottom',
        },
      },
    },
  });
}

function processExcel(arrayBuffer) {
  const workbook = XLSX.read(arrayBuffer, { type: 'array' });
  const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
  const json = XLSX.utils.sheet_to_json(firstSheet, { defval: '' });
  records = json.map(normalizeRow);
  render();
}

fileInput.addEventListener('change', async (event) => {
  const [file] = event.target.files;
  if (!file) return;
  const buffer = await file.arrayBuffer();
  processExcel(buffer);
});

thresholdInput.addEventListener('input', render);

sampleBtn.addEventListener('click', () => {
  records = [
    { employeeId: 'E001', name: 'Ava', punchIn: '09:00', punchOut: '18:00', totalHours: 8.5 },
    { employeeId: 'E002', name: 'Noah', punchIn: '09:20', punchOut: '16:50', totalHours: 7.5 },
    { employeeId: 'E003', name: 'Liam', punchIn: '08:55', punchOut: '17:15', totalHours: 8.0 },
  ];
  render();
});

render();
