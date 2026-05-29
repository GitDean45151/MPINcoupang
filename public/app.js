/**
 * Coupang Ads Dashboard Client Logic
 * Interacts with Node.js backend endpoints for fetching base data, 
 * uploading single sales updates, resetting modifications, and managing UI state.
 */

// App State
let joinedRecords = [];
let filteredRecords = [];
let chartInstance = null;
let agencyGrowthRate = 0;
let agencyQ1QoqTotal = 0;
let agencyQ2QoqTotal = 0;
let expectedQ1QoqTotal = 0;
let expectedQ2QoqTotal = 0;
let q1ChartValues = {
  'group 1': 5506674,
  'group 2': 17493317,
  'group 3': 490090134,
  'group 4': 21605441
};

// Pivot table sorting state
let pivotSortColumn = 'incentive';
let pivotSortAsc = false;

// Pagination config
let currentPage = 1;
const rowsPerPage = 10;

// Initialize Dashboard
document.addEventListener('DOMContentLoaded', async () => {
  initEventListeners();
  await fetchBaseData();
});

// Fetch base joined data from backend on startup
async function fetchBaseData() {
  showLoading('서버에서 기초 데이터를 불러오는 중...');
  try {
    const res = await fetch('/api/base-data');
    if (!res.ok) {
      throw new Error('기초 데이터를 불러올 수 없습니다. coupang.xlsx 파일이 서버에 있는지 확인해 주세요.');
    }
    const data = await res.json();
    
    joinedRecords = data.records;
    filteredRecords = [...joinedRecords];
    agencyGrowthRate = data.agencyGrowthRate !== undefined ? data.agencyGrowthRate : 0;
    agencyQ1QoqTotal = data.currentQ1QoqTotal !== undefined ? data.currentQ1QoqTotal : (data.q1QoqTotal !== undefined ? data.q1QoqTotal : 0);
    agencyQ2QoqTotal = data.currentQ2QoqTotal !== undefined ? data.currentQ2QoqTotal : (data.q2QoqTotal !== undefined ? data.q2QoqTotal : 0);
    expectedQ1QoqTotal = data.expectedQ1QoqTotal !== undefined ? data.expectedQ1QoqTotal : 0;
    expectedQ2QoqTotal = data.expectedQ2QoqTotal !== undefined ? data.expectedQ2QoqTotal : 0;
    
    if (data.q1ChartValues) {
      q1ChartValues = data.q1ChartValues;
    }
    
    updateDbStatusUI(data.totalUpdatesCount, data.lastUpdated);
    updateFilterDropdowns();
    
    // Render dashboard views
    renderDashboard();
    
    // Hide empty state, show grid
    document.getElementById('dashboard-empty-state').classList.add('hidden');
    document.getElementById('dashboard-content-grid').classList.remove('hidden');
  } catch (error) {
    console.error('Error fetching base data:', error);
    const emptyState = document.getElementById('dashboard-empty-state');
    emptyState.querySelector('h2').textContent = '데이터 로딩 실패';
    emptyState.querySelector('p').textContent = error.message;
    emptyState.querySelector('.empty-state-icon').textContent = '⚠️';
  } finally {
    hideLoading();
  }
}

// Update DB status dashboard labels
function updateDbStatusUI(updatesCount, lastUpdatedTime) {
  document.getElementById('db-updates-count').textContent = `${updatesCount}건`;
  
  const lastUpdatedEl = document.getElementById('db-last-updated');
  if (lastUpdatedTime) {
    const date = new Date(lastUpdatedTime);
    lastUpdatedEl.textContent = date.toLocaleTimeString('ko-KR', {
      hour: '2-digit',
      minute: '2-digit',
      second: '2-digit'
    });
  } else {
    lastUpdatedEl.textContent = '적용 없음';
  }

  // Show/Hide reset updates button
  const resetBtn = document.getElementById('reset-updates-btn');
  if (updatesCount > 0) {
    resetBtn.classList.remove('hidden');
  } else {
    resetBtn.classList.add('hidden');
  }
}

// Set up UI Event Listeners
function initEventListeners() {
  // Theme Toggle
  const themeToggle = document.getElementById('theme-toggle');
  themeToggle.addEventListener('click', () => {
    document.body.classList.toggle('dark-mode');
    document.body.classList.toggle('light-mode');
    if (joinedRecords.length > 0) {
      renderChart();
    }
  });

  // File Upload Elements
  const uploadZone = document.getElementById('upload-zone');
  const fileInput = document.getElementById('file-input-update');
  const uploadTriggerBtn = document.getElementById('btn-upload-trigger');

  uploadTriggerBtn.addEventListener('click', (e) => {
    e.stopPropagation(); // Avoid triggering parent uploadZone click
    fileInput.click();
  });

  uploadZone.addEventListener('click', () => {
    fileInput.click();
  });

  fileInput.addEventListener('change', async (e) => {
    const file = e.target.files[0];
    if (file) {
      await handleSalesUpdateUpload(file);
    }
  });

  // Drag and drop events for upload zone
  uploadZone.addEventListener('dragover', (e) => {
    e.preventDefault();
    uploadZone.classList.add('dragover');
  });

  uploadZone.addEventListener('dragleave', () => {
    uploadZone.classList.remove('dragover');
  });

  uploadZone.addEventListener('drop', async (e) => {
    e.preventDefault();
    uploadZone.classList.remove('dragover');
    
    const file = e.dataTransfer.files[0];
    if (!file) return;

    // Check extension
    const ext = file.name.split('.').pop().toLowerCase();
    if (!['csv', 'xlsx', 'xls'].includes(ext)) {
      showUploadStatusCard(false, file.name, '지원되지 않는 형식 (.csv, .xlsx, .xls)');
      return;
    }

    fileInput.files = e.dataTransfer.files;
    await handleSalesUpdateUpload(file);
  });

  // Reset Updates Button
  document.getElementById('reset-updates-btn').addEventListener('click', handleResetUpdates);

  // Generate Mock Data Button
  document.getElementById('generate-mock-btn').addEventListener('click', loadMockData);

  // Grid controls
  document.getElementById('grid-search').addEventListener('input', handleFilterChange);
  document.getElementById('filter-team').addEventListener('change', handleFilterChange);
  document.getElementById('filter-group').addEventListener('change', handleFilterChange);

  // Pagination buttons
  document.getElementById('btn-page-prev').addEventListener('click', () => {
    if (currentPage > 1) {
      currentPage--;
      renderDataGrid();
    }
  });
  document.getElementById('btn-page-next').addEventListener('click', () => {
    const maxPage = Math.ceil(filteredRecords.length / rowsPerPage);
    if (currentPage < maxPage) {
      currentPage++;
      renderDataGrid();
    }
  });

  // Export Excel
  document.getElementById('export-excel-btn').addEventListener('click', exportToExcel);

  // Pivot Table sorting headers
  const pivotHeaders = document.querySelectorAll('#pivot-header-row th.sortable');
  pivotHeaders.forEach(th => {
    th.addEventListener('click', () => {
      const col = th.dataset.sort;
      if (pivotSortColumn === col) {
        pivotSortAsc = !pivotSortAsc;
      } else {
        pivotSortColumn = col;
        pivotSortAsc = false; // default descending on new column
      }
      renderPivotTable();
    });
  });
}

// Upload the sales update file to the backend
async function handleSalesUpdateUpload(file) {
  showLoading(`${file.name} 업로드 및 대시보드 갱신 중...`);
  
  const formData = new FormData();
  formData.append('updateFile', file);

  try {
    const res = await fetch('/api/upload-update', {
      method: 'POST',
      body: formData
    });

    if (!res.ok) {
      const errData = await res.json();
      throw new Error(errData.error || '파일 업로드 처리에 실패했습니다.');
    }

    const data = await res.json();
    
    joinedRecords = data.records;
    filteredRecords = [...joinedRecords];
    agencyGrowthRate = data.agencyGrowthRate !== undefined ? data.agencyGrowthRate : 0;
    agencyQ1QoqTotal = data.currentQ1QoqTotal !== undefined ? data.currentQ1QoqTotal : (data.q1QoqTotal !== undefined ? data.q1QoqTotal : 0);
    agencyQ2QoqTotal = data.currentQ2QoqTotal !== undefined ? data.currentQ2QoqTotal : (data.q2QoqTotal !== undefined ? data.q2QoqTotal : 0);
    expectedQ1QoqTotal = data.expectedQ1QoqTotal !== undefined ? data.expectedQ1QoqTotal : 0;
    expectedQ2QoqTotal = data.expectedQ2QoqTotal !== undefined ? data.expectedQ2QoqTotal : 0;
    if (data.q1ChartValues) {
      q1ChartValues = data.q1ChartValues;
    }
    currentPage = 1;

    // Show success status card
    showUploadStatusCard(true, file.name, `${data.updatedCount}개 행의 매출이 성공적으로 반영되었습니다.`);
    
    // Update labels
    updateDbStatusUI(data.totalUpdatesCount, new Date());
    updateFilterDropdowns();
    
    // Re-render
    renderDashboard();
    
  } catch (error) {
    console.error('Upload error:', error);
    showUploadStatusCard(false, file.name, error.message);
  } finally {
    hideLoading();
  }
}

// Reset updates on the backend
async function handleResetUpdates() {
  if (!confirm('정말 모든 매출 업데이트를 초기화하고 원본 coupang.xlsx 상태로 되돌리시겠습니까?')) {
    return;
  }

  showLoading('데이터 원본으로 복원 중...');
  try {
    const res = await fetch('/api/reset-updates', { method: 'POST' });
    if (!res.ok) {
      throw new Error('초기화 요청이 실패했습니다.');
    }
    const data = await res.json();

    joinedRecords = data.records;
    filteredRecords = [...joinedRecords];
    agencyGrowthRate = data.agencyGrowthRate !== undefined ? data.agencyGrowthRate : 0;
    agencyQ1QoqTotal = data.currentQ1QoqTotal !== undefined ? data.currentQ1QoqTotal : (data.q1QoqTotal !== undefined ? data.q1QoqTotal : 0);
    agencyQ2QoqTotal = data.currentQ2QoqTotal !== undefined ? data.currentQ2QoqTotal : (data.q2QoqTotal !== undefined ? data.q2QoqTotal : 0);
    expectedQ1QoqTotal = data.expectedQ1QoqTotal !== undefined ? data.expectedQ1QoqTotal : 0;
    expectedQ2QoqTotal = data.expectedQ2QoqTotal !== undefined ? data.expectedQ2QoqTotal : 0;
    if (data.q1ChartValues) {
      q1ChartValues = data.q1ChartValues;
    }
    currentPage = 1;

    // Reset status cards
    document.getElementById('file-status-card').classList.add('hidden');
    document.getElementById('file-input-update').value = '';
    
    updateDbStatusUI(0, null);
    updateFilterDropdowns();
    
    // Re-render
    renderDashboard();
  } catch (error) {
    alert(`초기화 실패: ${error.message}`);
  } finally {
    hideLoading();
  }
}

// Show file upload success or error card in the sidebar
function showUploadStatusCard(isSuccess, fileName, message) {
  const card = document.getElementById('file-status-card');
  const nameEl = document.getElementById('uploaded-file-name');
  const infoEl = document.getElementById('uploaded-file-info');
  
  card.classList.remove('hidden', 'error');
  if (!isSuccess) {
    card.classList.add('error');
    card.querySelector('.file-status-icon').textContent = '❌';
  } else {
    card.querySelector('.file-status-icon').textContent = '📄';
  }

  nameEl.textContent = fileName;
  infoEl.textContent = message;
}

// Refresh all components in dashboard
function renderDashboard() {
  renderKpis();
  renderChart();
  renderPivotTable();
  renderDataGrid();
}

// Generate dropdown items dynamically based on processed data
function updateFilterDropdowns() {
  const currentTeamVal = document.getElementById('filter-team').value;
  const currentGroupVal = document.getElementById('filter-group').value;

  const teams = [...new Set(joinedRecords.map(r => r.team))].sort();
  const groups = [...new Set(joinedRecords.map(r => r.commissionGroup))].sort();
  
  const filterTeam = document.getElementById('filter-team');
  const filterGroup = document.getElementById('filter-group');
  
  // Clear extra options, keep default
  filterTeam.innerHTML = '<option value="">모든 팀</option>';
  filterGroup.innerHTML = '<option value="">모든 커미션 그룹</option>';
  
  teams.forEach(t => {
    filterTeam.innerHTML += `<option value="${t}">${t}</option>`;
  });
  
  groups.forEach(g => {
    filterGroup.innerHTML += `<option value="${g}">${g}</option>`;
  });

  // Restore previous selections if they still exist
  if (teams.includes(currentTeamVal)) {
    filterTeam.value = currentTeamVal;
  }
  if (groups.includes(currentGroupVal)) {
    filterGroup.value = currentGroupVal;
  }
}

// Perform client-side filter and search
function handleFilterChange() {
  const search = document.getElementById('grid-search').value.toLowerCase().trim();
  const team = document.getElementById('filter-team').value;
  const group = document.getElementById('filter-group').value;

  filteredRecords = joinedRecords.filter(r => {
    const matchesSearch = !search || 
      String(r.vendorId).toLowerCase().includes(search) || 
      r.vendorName.toLowerCase().includes(search) || 
      r.marketer.toLowerCase().includes(search);
      
    const matchesTeam = !team || r.team === team;
    const matchesGroup = !group || r.commissionGroup === group;
    
    return matchesSearch && matchesTeam && matchesGroup;
  });

  currentPage = 1;
  renderDashboard();
}

// Render Top-line KPI Cards
function renderKpis() {
  let total15Commission = 0;
  let totalQoqCommission = 0;
  let totalIncentive = 0;

  joinedRecords.forEach(r => {
    total15Commission += r.commission15;
    totalQoqCommission += r.expectedQoqCommission;
    totalIncentive += r.expectedMarketerIncentive;
  });

  const agencyGrowthPercent = agencyGrowthRate * 100;
  const qoqEl = document.getElementById('val-qoq-growth');
  qoqEl.textContent = (agencyGrowthPercent >= 0 ? '+' : '') + agencyGrowthPercent.toFixed(2) + '%';
  
  const diff = agencyQ2QoqTotal - agencyQ1QoqTotal;
  const diffEl = document.getElementById('val-qoq-diff');
  diffEl.textContent = formatKRW(Math.abs(diff)) + (diff >= 0 ? ' 증가' : ' 감소');
  diffEl.className = 'kpi-sub ' + (diff >= 0 ? 'positive' : 'negative');
  
  const companyTotalFees = total15Commission + totalQoqCommission;
  document.getElementById('val-1q-rev').textContent = formatKRW(agencyQ1QoqTotal);
  document.getElementById('val-2q-rev').textContent = formatKRW(agencyQ2QoqTotal);
  document.getElementById('val-expected-1q-rev').textContent = formatKRW(expectedQ1QoqTotal);
  document.getElementById('val-expected-2q-rev').textContent = formatKRW(expectedQ2QoqTotal);
  document.getElementById('val-company-fees').textContent = formatKRW(companyTotalFees);
  document.getElementById('val-company-fees-detail').textContent = 
    `신규15%: ${formatKRW(total15Commission)} | QoQ: ${formatKRW(totalQoqCommission)}`;
    
  document.getElementById('val-marketer-incentive').textContent = formatKRW(totalIncentive);
}

// Render Growth Rate Distribution Chart (Bar Chart)
function renderChart() {
  const canvas = document.getElementById('growthChart');
  if (!canvas) return;
  const ctx = canvas.getContext('2d');
  
  const groups = ['group 1', 'group 2', 'group 3', 'group 4'];
  const q1Sums = [0, 0, 0, 0];
  const q2Sums = [0, 0, 0, 0];
  const q2PredictedSums = [0, 0, 0, 0];

  joinedRecords.forEach(r => {
    const groupIdx = groups.indexOf(r.commissionGroup);
    if (groupIdx !== -1) {
      q2Sums[groupIdx] += r.q2Revenue;
      q2PredictedSums[groupIdx] += r.q2PredictedRevenue || r.q2Revenue;
    }
  });

  groups.forEach((g, idx) => {
    q1Sums[idx] = q1ChartValues[g] || 0;
  });

  const subtitleEl = document.querySelector('.chart-card .card-subtitle');
  if (subtitleEl) {
    subtitleEl.textContent = `커미션 그룹별 1Q 대비 2Q 광고비 합산 비교`;
  }

  // Chart theme configurations
  const isDark = document.body.classList.contains('dark-mode');
  const textColor = isDark ? '#9ca3af' : '#475569';
  const gridColor = isDark ? 'rgba(255, 255, 255, 0.08)' : 'rgba(0, 0, 0, 0.08)';

  const datasets = [
    {
      label: '1Q 광고비',
      data: q1Sums,
      backgroundColor: 'rgba(6, 182, 212, 0.75)',
      borderColor: '#06b6d4',
      borderWidth: 1.5,
      borderRadius: 6,
      barPercentage: 0.8,
      categoryPercentage: 0.7
    },
    {
      label: '2Q 광고비 (~260513)',
      data: q2Sums,
      backgroundColor: 'rgba(217, 70, 239, 0.75)',
      borderColor: '#d946ef',
      borderWidth: 1.5,
      borderRadius: 6,
      barPercentage: 0.8,
      categoryPercentage: 0.7
    },
    {
      label: '2Q 전체기간 예측광고비',
      data: q2PredictedSums,
      backgroundColor: 'rgba(245, 158, 11, 0.75)',
      borderColor: '#f59e0b',
      borderWidth: 1.5,
      borderRadius: 6,
      barPercentage: 0.8,
      categoryPercentage: 0.7
    }
  ];

  if (chartInstance) {
    chartInstance.destroy();
  }

  chartInstance = new Chart(ctx, {
    type: 'bar',
    data: {
      labels: groups.map(g => g.toUpperCase()),
      datasets: datasets
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      scales: {
        x: {
          grid: { color: gridColor },
          ticks: { color: textColor, font: { family: 'Lexend', size: 11 } }
        },
        y: {
          grid: { color: gridColor },
          ticks: { 
            color: textColor, 
            font: { family: 'Lexend', size: 11 },
            callback: function(value) {
              // Compact KRW Formatting (e.g. ₩4.9억)
              if (value >= 100000000) {
                return '₩' + (value / 100000000).toFixed(1) + '억';
              } else if (value >= 10000) {
                return '₩' + (value / 10000).toFixed(0) + '만';
              }
              return '₩' + value;
            }
          }
        }
      },
      plugins: {
        legend: {
          position: 'top',
          labels: { color: textColor, font: { family: 'Lexend', size: 12, weight: 600 } }
        },
        tooltip: {
          backgroundColor: isDark ? 'rgba(15, 14, 46, 0.95)' : 'rgba(255, 255, 255, 0.95)',
          titleColor: isDark ? '#fff' : '#000',
          bodyColor: isDark ? '#e5e7eb' : '#374151',
          borderColor: isDark ? 'rgba(255, 255, 255, 0.1)' : 'rgba(0, 0, 0, 0.1)',
          borderWidth: 1,
          padding: 10,
          titleFont: { family: 'Outfit', size: 13, weight: 'bold' },
          bodyFont: { family: 'Inter', size: 12 },
          callbacks: {
            label: function(context) {
              let label = context.dataset.label || '';
              if (label) {
                label += ': ';
              }
              if (context.parsed.y !== null) {
                label += '₩' + new Intl.NumberFormat('ko-KR').format(context.parsed.y);
              }
              return label;
            }
          }
        }
      }
    }
  });
}

// Render Team > Marketer Pivot Table
function renderPivotTable() {
  // Update Header Sort Icons
  const pivotHeaders = document.querySelectorAll('#pivot-header-row th.sortable');
  pivotHeaders.forEach(th => {
    const icon = th.querySelector('.sort-icon');
    if (icon) {
      if (th.dataset.sort === pivotSortColumn) {
        icon.textContent = pivotSortAsc ? ' ▲' : ' ▼';
        th.classList.add('sorted');
      } else {
        icon.textContent = '';
        th.classList.remove('sorted');
      }
    }
  });

  const pivotData = {};

  // Aggregate over filteredRecords instead of joinedRecords for active filtering
  filteredRecords.forEach(r => {
    if (!pivotData[r.team]) {
      pivotData[r.team] = {
        name: r.team,
        q1: 0,
        q2: 0,
        diff: 0,
        growth: 0,
        commission15: 0,
        qoqCommission: 0,
        incentive: 0,
        marketers: {}
      };
    }

    const tData = pivotData[r.team];
    tData.q1 += r.q1Revenue;
    tData.q2 += r.q2Revenue;
    tData.commission15 += r.commission15;
    tData.qoqCommission += r.expectedQoqCommission;
    tData.incentive += r.expectedMarketerIncentive;

    if (!tData.marketers[r.marketer]) {
      tData.marketers[r.marketer] = {
        name: r.marketer,
        q1: 0,
        q2: 0,
        diff: 0,
        growth: 0,
        commission15: 0,
        qoqCommission: 0,
        incentive: 0
      };
    }

    const mData = tData.marketers[r.marketer];
    mData.q1 += r.q1Revenue;
    mData.q2 += r.q2Revenue;
    mData.commission15 += r.commission15;
    mData.qoqCommission += r.expectedQoqCommission;
    mData.incentive += r.expectedMarketerIncentive;
  });

  // Calculate derivatives for sorting
  Object.values(pivotData).forEach(team => {
    team.diff = team.q2 - team.q1;
    team.growth = team.q1 > 0 ? (team.diff / team.q1) * 100 : (team.q2 > 0 ? 100 : 0);
    
    Object.values(team.marketers).forEach(m => {
      m.diff = m.q2 - m.q1;
      m.growth = m.q1 > 0 ? (m.diff / m.q1) * 100 : (m.q2 > 0 ? 100 : 0);
    });
  });

  // Helper compare function for sorting
  const compareFn = (a, b) => {
    let valA = a[pivotSortColumn];
    let valB = b[pivotSortColumn];
    
    if (pivotSortColumn === 'name') {
      return pivotSortAsc ? valA.localeCompare(valB) : valB.localeCompare(valA);
    }
    
    return pivotSortAsc ? valA - valB : valB - valA;
  };

  // Convert to arrays and sort
  const teamsArray = Object.values(pivotData).map(team => {
    const marketersArray = Object.values(team.marketers).sort(compareFn);
    return {
      ...team,
      marketersList: marketersArray
    };
  }).sort(compareFn);

  const tbody = document.getElementById('pivot-table-body');
  tbody.innerHTML = '';

  teamsArray.forEach((team, tIdx) => {
    const teamDiff = team.diff;
    const teamGrowth = team.growth;
    const teamId = `team-row-${tIdx}`;

    const teamRow = document.createElement('tr');
    teamRow.className = 'pivot-row-team';
    teamRow.dataset.target = teamId;
    teamRow.innerHTML = `
      <td>${team.name}</td>
      <td class="text-right">${formatKRW(team.q1)}</td>
      <td class="text-right">${formatKRW(team.q2)}</td>
      <td class="text-right ${teamDiff >= 0 ? 'positive' : 'negative'}" style="color: ${teamDiff >= 0 ? 'var(--success)' : 'var(--danger)'}">
        ${teamDiff >= 0 ? '+' : ''}${formatKRW(teamDiff)}
      </td>
      <td class="text-right">${teamGrowth.toFixed(2)}%</td>
      <td class="text-right">${formatKRW(team.commission15)}</td>
      <td class="text-right">${formatKRW(team.qoqCommission)}</td>
      <td class="text-right" style="color: var(--accent-primary); font-weight: bold;">${formatKRW(team.incentive)}</td>
    `;
    
    tbody.appendChild(teamRow);

    team.marketersList.forEach(m => {
      const mDiff = m.diff;
      const mGrowth = m.growth;

      const marketerRow = document.createElement('tr');
      marketerRow.className = `pivot-row-marketer ${teamId}`;
      marketerRow.innerHTML = `
        <td>${m.name}</td>
        <td class="text-right">${formatKRW(m.q1)}</td>
        <td class="text-right">${formatKRW(m.q2)}</td>
        <td class="text-right" style="color: ${mDiff >= 0 ? 'var(--success)' : 'var(--danger)'}">
          ${mDiff >= 0 ? '+' : ''}${formatKRW(mDiff)}
        </td>
        <td class="text-right">${mGrowth.toFixed(2)}%</td>
        <td class="text-right">${formatKRW(m.commission15)}</td>
        <td class="text-right">${formatKRW(m.qoqCommission)}</td>
        <td class="text-right" style="color: var(--accent-primary); font-weight: bold;">${formatKRW(m.incentive)}</td>
      `;
      tbody.appendChild(marketerRow);
    });

    teamRow.addEventListener('click', () => {
      teamRow.classList.toggle('collapsed');
      const marketerRows = document.querySelectorAll(`.${teamId}`);
      marketerRows.forEach(row => {
        row.classList.toggle('hidden');
      });
    });
  });
}

// Render paginated detailed Data Grid
function renderDataGrid() {
  const tbody = document.getElementById('grid-table-body');
  tbody.innerHTML = '';

  const totalCount = filteredRecords.length;
  const maxPage = Math.ceil(totalCount / rowsPerPage) || 1;
  
  if (currentPage > maxPage) currentPage = maxPage;
  if (currentPage < 1) currentPage = 1;

  document.getElementById('btn-page-prev').disabled = currentPage === 1;
  document.getElementById('btn-page-next').disabled = currentPage === maxPage;
  document.getElementById('grid-page-number').textContent = `페이지 ${currentPage} / ${maxPage}`;
  document.getElementById('grid-row-count').textContent = `총 ${totalCount}건 조회됨`;

  if (totalCount === 0) {
    tbody.innerHTML = `<tr><td colspan="15" style="text-align: center; padding: 32px; color: var(--text-muted);">조건에 부합하는 데이터가 없습니다.</td></tr>`;
    return;
  }

  const startIndex = (currentPage - 1) * rowsPerPage;
  const endIndex = Math.min(startIndex + rowsPerPage, totalCount);
  const pageRecords = filteredRecords.slice(startIndex, endIndex);

  pageRecords.forEach(r => {
    const tr = document.createElement('tr');
    
    if (r.growthAmount < 0) {
      tr.className = 'row-degrowth';
    } else if (r.marketerIncentiveRate === 0.10) {
      tr.className = 'row-max-incentive';
    }
    
    // Highlight if this row was recently updated in local session
    if (r.isUpdated) {
      tr.style.borderLeft = '4px solid var(--success)';
    }

    tr.innerHTML = `
      <td>${r.vendorId}</td>
      <td style="font-weight: 500;">${r.vendorName}</td>
      <td>${r.team}</td>
      <td>${r.marketer}</td>
      <td class="text-center"><span class="badge">${r.commissionGroup}</span></td>
      <td class="text-right">${formatKRW(r.q1Revenue)}</td>
      <td class="text-right">${formatKRW(r.q2Revenue)}</td>
      <td class="text-right ${r.growthAmount >= 0 ? 'positive' : 'negative'}" style="font-weight: 600;">
        ${r.growthAmount >= 0 ? '+' : ''}${r.growthRatePercent.toFixed(2)}%
      </td>
      <td class="text-center" style="font-weight: bold; color: ${r.isNewCommission ? 'var(--success)' : 'var(--text-muted)'}">
        ${r.isNewCommission ? '지급(O)' : '미지급(X)'}
      </td>
      <td class="text-right">${formatKRW(r.commission15)}</td>
      <td class="text-center" style="font-weight: bold; color: ${r.isQoqEligible ? 'var(--success)' : 'var(--text-muted)'}">
        ${r.isQoqEligible ? '지급(O)' : '미지급(X)'}
      </td>
      <td class="text-right">${(r.qoqCommissionRate * 100).toFixed(1)}%</td>
      <td class="text-right" style="font-weight: 600;">${formatKRW(r.expectedQoqCommission)}</td>
      <td class="text-right">${(r.marketerIncentiveRate * 100).toFixed(1)}%</td>
      <td class="text-right" style="font-weight: 700; color: var(--accent-primary);">${formatKRW(r.expectedMarketerIncentive)}</td>
    `;
    
    tbody.appendChild(tr);
  });
}

// Excel Export function (SheetJS)
function exportToExcel() {
  if (joinedRecords.length === 0) return;

  showLoading('엑셀 파일 생성 중...');
  
  const exportData = filteredRecords.map(r => ({
    '벤더 ID': r.vendorId,
    '업체명': r.vendorName,
    '팀': r.team,
    '마케터': r.marketer,
    '커미션 그룹': r.commissionGroup,
    '1Q 매출': r.q1Revenue,
    '2Q 매출': r.q2Revenue,
    '매출 증감액': r.growthAmount,
    '매출 증감률(%)': Number(r.growthRatePercent.toFixed(2)),
    '신규대상': r.isNewCommission ? '지급(O)' : '미지급(X)',
    '신규 수수료': r.commission15,
    'QoQ 대상': r.isQoqEligible ? '지급(O)' : '미지급(X)',
    'QoQ 수수료율(%)': r.qoqCommissionRate * 100,
    'QoQ 수수료액': r.expectedQoqCommission,
    '인센티브율(%)': r.marketerIncentiveRate * 100,
    '마케터 인센티브액': r.expectedMarketerIncentive
  }));

  const worksheet = XLSX.utils.json_to_sheet(exportData);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, '쿠팡 커미션 리포트');

  // Adjust column widths
  const max_len = exportData.reduce((w, row) => Object.keys(row).reduce((max, key) => Math.max(max, String(row[key]).length), w), 10);
  worksheet['!cols'] = Object.keys(exportData[0]).map(() => ({ wch: Math.max(max_len, 12) }));

  XLSX.writeFile(workbook, `쿠팡_광고_커미션_정산_리포트_${new Date().toISOString().split('T')[0]}.xlsx`);
  hideLoading();
}

// Generate Mock Data locally for instant testing (Frontend Fallback)
function loadMockData() {
  showLoading('모의 샘플 데이터 생성 중...');

  const mockSellers = [
    { 'vendor id': '1001', 'vendor name': '나이키 코리아' },
    { 'vendor id': '1002', 'vendor name': '아디다스 공식몰' },
    { 'vendor id': '1003', 'vendor name': '밸런시스코리아' },
    { 'vendor id': '1004', 'vendor name': '바이퀸즈' },
    { 'vendor id': '1005', 'vendor name': '정상커뮤니케이션즈' },
    { 'vendor id': '1006', 'vendor name': '스마트스팟' },
    { 'vendor id': '1007', 'vendor name': '레고 코리아' },
    { 'vendor id': '1008', 'vendor name': '삼성전자 판매처' },
    { 'vendor id': '1009', 'vendor name': 'LG 리테일' },
    { 'vendor id': '1010', 'vendor name': '어블러썸 뷰티' },
    { 'vendor id': '1011', 'vendor name': '한샘 가구대리점' },
    { 'vendor id': '1012', 'vendor name': '다이슨 공식숍' },
    { 'vendor id': '1013', 'vendor name': '샤오미 유통망' },
    { 'vendor id': '1014', 'vendor name': '필립스 조명' },
    { 'vendor id': '1015', 'vendor name': '테팔 리빙앤뷰티' }
  ];

  const marketers = [
    { team: '마케팅1팀', name: '김용덕' },
    { team: '마케팅2팀', name: '엄도윤' },
    { team: '마케팅3팀', name: '최혜림' },
    { team: '5팀', name: '이예진' },
    { team: '9팀', name: '신승규' },
    { team: '지원팀', name: '황병동' },
    { team: '7팀', name: '남현민' }
  ];

  const mockRecords = mockSellers.map((s, idx) => {
    const vid = 'A0000' + s['vendor id'];
    const marketerInfo = marketers[idx % marketers.length];
    const group = `group ${(idx % 4) + 1}`; // group 1 to group 4
    
    let q1Revenue = 0;
    let q2Revenue = 0;
    
    if (idx % 5 === 0) {
      q1Revenue = 50000000 + idx * 1000000;
      q2Revenue = 42000000 + idx * 800000;
    } else if (idx % 5 === 1) {
      q1Revenue = 20000000 + idx * 500000;
      q2Revenue = 21000000 + idx * 550000;
    } else if (idx % 5 === 2) {
      q1Revenue = 15000000 + idx * 400000;
      q2Revenue = 17500000 + idx * 500000;
    } else if (idx % 5 === 3) {
      q1Revenue = 10000000 + idx * 300000;
      q2Revenue = 12500000 + idx * 400000;
    } else {
      q1Revenue = 8000000 + idx * 200000;
      q2Revenue = 12000000 + idx * 500000;
    }

    const growthAmount = q2Revenue - q1Revenue;
    const growthRate = q1Revenue > 0 ? (q2Revenue - q1Revenue) / q1Revenue : (q2Revenue > 0 ? 1.0 : 0.0);
    const growthRatePercent = growthRate * 100;

    // 15% new commission on Group 1 & 2
    const isNewCommission = ['group 1', 'group 2'].includes(group);
    
    // Simulate 90-day target and QoQ target revenues
    const q290DayRevenue = isNewCommission ? (group === 'group 2' ? q2Revenue * 0.9 : q2Revenue) : 0;
    const q2QoqRevenue = ['group 2', 'group 3', 'group 4'].includes(group) ? (group === 'group 2' ? q2Revenue * 0.1 : q2Revenue) : 0;
    
    // Simulate prediction for target vendors (May list targets)
    const q2PredictedRevenue = Math.round(q2Revenue * 91 / 43);

    const commission15 = isNewCommission ? q290DayRevenue * 0.15 : 0;
    const isQoqEligible = ['group 2', 'group 3', 'group 4'].includes(group);

    // Mock agency rates (since agencyGrowthRate is mocked at 25% = 0.25)
    const qoqCommissionRate = 0.175; // 17.5% for 20-30% growth
    const marketerIncentiveRate = 0.07; // 7% for 20-30% growth

    const expectedQoqCommission = isQoqEligible ? q2QoqRevenue * qoqCommissionRate : 0;
    
    // Marketer incentives logic
    const marketerNewIncentive = isNewCommission ? q290DayRevenue * 0.05 : 0;
    const marketerQoqIncentive = isQoqEligible ? q2QoqRevenue * marketerIncentiveRate : 0;
    const expectedMarketerIncentive = marketerNewIncentive + marketerQoqIncentive;

    return {
      id: `MOCK_${idx}`,
      vendorId: vid,
      vendorName: s['vendor name'],
      team: marketerInfo.team,
      marketer: marketerInfo.name,
      commissionGroup: group,
      q1Revenue,
      q2Revenue,
      q2PredictedRevenue,
      growthAmount,
      growthRatePercent,
      isNewCommission,
      commission15,
      isQoqEligible,
      qoqCommissionRate,
      expectedQoqCommission,
      marketerIncentiveRate,
      marketerNewIncentive,
      marketerGrowthIncentive: marketerQoqIncentive,
      expectedMarketerIncentive,
      isUpdated: false
    };
  });

  joinedRecords = mockRecords;
  filteredRecords = [...joinedRecords];
  currentPage = 1;

  let mockQ1QoqTotal = 0;
  let mockQ2QoqTotal = 0;
  mockRecords.forEach(r => {
    if (r.commissionGroup === 'group 3') {
      mockQ1QoqTotal += r.q1Revenue;
      mockQ2QoqTotal += r.q2Revenue;
    } else if (r.commissionGroup === 'group 2') {
      mockQ1QoqTotal += r.q1Revenue;
      mockQ2QoqTotal += r.q2Revenue * 0.1;
    }
  });

  agencyQ1QoqTotal = mockQ1QoqTotal;
  agencyQ2QoqTotal = mockQ2QoqTotal;
  expectedQ1QoqTotal = 503708883; // default mock expected Q1
  expectedQ2QoqTotal = 559924393; // default mock expected Q2
  agencyGrowthRate = agencyQ1QoqTotal > 0 ? (agencyQ2QoqTotal / agencyQ1QoqTotal - 1) : 0;
  
  q1ChartValues = {
    'group 1': 5506674,
    'group 2': 17493317,
    'group 3': 490090134,
    'group 4': 21605441
  };
  
  updateDbStatusUI('샘플', new Date());
  updateFilterDropdowns();
  renderDashboard();

  // Hide empty state, show grid
  document.getElementById('dashboard-empty-state').classList.add('hidden');
  document.getElementById('dashboard-content-grid').classList.remove('hidden');
  
  hideLoading();
}

// ==========================================================================
// Utility Helpers
// ==========================================================================

// Format numeric values to Korean Won (KRW) currency format
function formatKRW(num) {
  return '₩' + new Intl.NumberFormat('ko-KR').format(Math.round(num || 0));
}

// Show/Hide Loading Overlay
function showLoading(text) {
  document.getElementById('loading-text').textContent = text;
  document.getElementById('loading-overlay').classList.remove('hidden');
}

function hideLoading() {
  document.getElementById('loading-overlay').classList.add('hidden');
}
