let pollTimer = null;
let scrollTimer = null;
let currentScrollRow = 0;
let totalRows = 0;
let manualScrollMode = false;
let inactivityTimer = null;
const INACTIVITY_TIMEOUT = 25 * 60 * 1000; // 25 minuti in ms

// ===================== OUTPUT PAGE STATE =====================
let outputChart = null;
let outputPollTimer = null;
let rotationTimer = null;
let currentPage = 'dashboard'; // 'dashboard' or 'output'
let rotationElapsed = 0; // seconds elapsed in current rotation phase
let outputConfig = {
    daily_target: 2000,
    chart_title: "Output VDW RO",
    rotation_show_seconds: 300,   // 5 min output page
    rotation_cycle_seconds: 900,  // 15 min total cycle
    output_poll_seconds: 60
};

// ===================== DASHBOARD POLLING =====================

async function fetchStatus() {
    try {
        const resp = await fetch('/api/status');
        if (!resp.ok) throw new Error('HTTP ' + resp.status);
        const data = await resp.json();
        renderDashboard(data);
    } catch (e) {
        console.error('Fetch error:', e);
    }
}

function renderDashboard(data) {
    // Summary
    const s = data.summary || {};
    document.getElementById('count-total').textContent = s.total || 0;
    document.getElementById('count-green').textContent = s.green || 0;
    document.getElementById('count-yellow').textContent = s.yellow || 0;
    document.getElementById('count-red').textContent = s.red || 0;
    document.getElementById('count-oop').textContent = s.out_of_plan || 0;

    // Header info
    document.getElementById('excel-file').textContent = data.excel_file || '-';
    if (data.last_update) {
        const dt = new Date(data.last_update);
        document.getElementById('last-update').textContent = dt.toLocaleString();
    }

    // Error banner
    const banner = document.getElementById('error-banner');
    if (data.last_error) {
        banner.textContent = data.last_error;
        banner.style.display = 'block';
    } else {
        banner.style.display = 'none';
    }

    // Run button state
    const btn = document.getElementById('btn-run-now');
    if (data.cycle_running) {
        btn.disabled = true;
        btn.textContent = 'Running...';
    } else {
        btn.disabled = false;
        btn.textContent = 'Run Now';
    }

    // Table
    const tbody = document.getElementById('table-body');
    const rows = data.rows || [];
    totalRows = rows.length;

    if (rows.length === 0) {
        tbody.innerHTML = '<tr><td colspan="8" class="empty-msg">No data available. Waiting for cycle...</td></tr>';
        stopAutoScroll();
        return;
    }

    let html = '';
    for (const r of rows) {
        let rowClass = '';
        let dotClass = 'dot-' + r.status_color;

        if (r.is_out_of_plan) {
            rowClass = 'row-out-of-plan blink';
            dotClass = 'dot-red';
        } else if (r.status_color === 'red') {
            rowClass = 'row-red';
        } else if (r.status_color === 'yellow') {
            rowClass = 'row-yellow';
        } else {
            rowClass = 'row-green';
        }

        let planLabel;
        if (r.is_out_of_plan) {
            planLabel = '<em>NOT IN PLAN</em>';
        } else if (r.qty_adjusted) {
            planLabel = `<span class="dot dot-adjusted" title="Qty pianificata Excel: ${r.original_planned_qty}"></span>${r.planned_qty_day}`;
        } else {
            planLabel = r.planned_qty_day;
        }
        const deficitLabel = r.is_out_of_plan ? '-' : r.projected_deficit;
        const expectedLabel = r.is_out_of_plan ? '-' : r.expected_by_now;
        const projectedLabel = r.is_out_of_plan ? '-' : r.projected_end_qty;

        // Star indicator
        let starHtml = '';
        if (r.context_star === 'yellow') {
            starHtml = '<span class="star star-yellow" title="Scheduled in upcoming days">&#9733;</span>';
        } else if (r.context_star === 'blue') {
            starHtml = '<span class="star star-blue" title="Delayed from previous days">&#9733;</span>';
        }

        // Context note tooltip
        let noteAttr = '';
        if (r.context_note) {
            noteAttr = ` title="${escHtml(r.context_note)}"`;
        }

        html += `<tr class="${rowClass}"${noteAttr}>
            <td>${escHtml(r.order_number)}</td>
            <td>${escHtml(r.product_code)}</td>
            <td>${escHtml(r.phase)}</td>
            <td style="text-align:center;">${planLabel}</td>
            <td><div class="qty-cell"><span class="dot ${dotClass}"></span>${r.qty_done}${starHtml}</div></td>
            <td style="text-align:center;">${expectedLabel}</td>
            <td style="text-align:center;">${projectedLabel}</td>
            <td style="text-align:center;">${deficitLabel}</td>
        </tr>`;
    }
    tbody.innerHTML = html;

    // Auto-scroll solo se NON in modalita manuale
    if (!manualScrollMode) {
        setupAutoScroll();
    }

    // Mapping errors
    const errSection = document.getElementById('mapping-errors');
    const errList = document.getElementById('error-list');
    const errors = data.mapping_errors || [];
    if (errors.length > 0) {
        errSection.style.display = 'block';
        errList.innerHTML = errors.map(e => `<p>${escHtml(e.message || JSON.stringify(e))}</p>`).join('');
    } else {
        errSection.style.display = 'none';
    }
}

// ===================== SCROLL MODE MANAGEMENT =====================

function toggleScrollMode() {
    if (manualScrollMode) {
        // Torna ad auto-scroll
        enableAutoScroll();
    } else {
        // Passa a scroll manuale
        enableManualScroll();
    }
}

function enableManualScroll() {
    manualScrollMode = true;
    stopAutoScroll();

    const wrapper = document.getElementById('table-wrapper');
    wrapper.style.overflowY = 'auto';

    updateScrollButton();
    resetInactivityTimer();
}

function enableAutoScroll() {
    manualScrollMode = false;
    clearInactivityTimer();

    const wrapper = document.getElementById('table-wrapper');
    wrapper.style.overflowY = 'hidden';

    setupAutoScroll();
    updateScrollButton();
}

function updateScrollButton() {
    const btn = document.getElementById('btn-scroll-mode');
    if (manualScrollMode) {
        btn.textContent = 'Auto Scroll';
        btn.classList.add('manual-active');
    } else {
        btn.textContent = 'Manual Scroll';
        btn.classList.remove('manual-active');
    }
}

// ===================== INACTIVITY TIMER (25 min) =====================

function resetInactivityTimer() {
    clearInactivityTimer();
    if (manualScrollMode) {
        inactivityTimer = setTimeout(() => {
            enableAutoScroll();
        }, INACTIVITY_TIMEOUT);
    }
}

function clearInactivityTimer() {
    if (inactivityTimer) {
        clearTimeout(inactivityTimer);
        inactivityTimer = null;
    }
}

function onUserActivity() {
    if (manualScrollMode) {
        resetInactivityTimer();
    }
}

// ===================== AUTO-SCROLL LOGIC =====================

function setupAutoScroll() {
    if (manualScrollMode) return;

    const wrapper = document.getElementById('table-wrapper');
    wrapper.style.overflowY = 'hidden';

    if (wrapper.scrollHeight <= wrapper.clientHeight) {
        stopAutoScroll();
        return;
    }

    if (!scrollTimer) {
        currentScrollRow = 0;
        wrapper.scrollTop = 0;
        scrollTimer = setInterval(scrollOneRow, 5000);
    }
}

function scrollOneRow() {
    if (manualScrollMode) return;

    const wrapper = document.getElementById('table-wrapper');
    const tbody = document.getElementById('table-body');
    const rows = tbody.querySelectorAll('tr');

    if (rows.length === 0) return;

    currentScrollRow++;

    if (currentScrollRow >= rows.length) {
        currentScrollRow = 0;
        wrapper.scrollTo({ top: 0, behavior: 'smooth' });
        return;
    }

    const targetRow = rows[currentScrollRow];
    if (targetRow) {
        const rowTop = targetRow.offsetTop;
        wrapper.scrollTo({ top: rowTop, behavior: 'smooth' });
    }
}

function stopAutoScroll() {
    if (scrollTimer) {
        clearInterval(scrollTimer);
        scrollTimer = null;
    }
    currentScrollRow = 0;
}

// ===================== MANUAL VIEW SWITCH =====================

function manualSwitchView() {
    // Reset rotation timer so the full duration restarts from now
    rotationElapsed = 0;

    if (currentPage === 'dashboard') {
        switchToOutputPage();
    } else {
        switchToDashboardPage();
    }
}

// ===================== PAGE ROTATION =====================

function startRotation() {
    const dashboardDuration = outputConfig.rotation_cycle_seconds - outputConfig.rotation_show_seconds;
    rotationElapsed = 0;
    currentPage = 'dashboard';

    // Check every second
    rotationTimer = setInterval(() => {
        rotationElapsed++;

        if (currentPage === 'dashboard' && rotationElapsed >= dashboardDuration) {
            switchToOutputPage();
            rotationElapsed = 0;
        } else if (currentPage === 'output' && rotationElapsed >= outputConfig.rotation_show_seconds) {
            switchToDashboardPage();
            rotationElapsed = 0;
        }
    }, 1000);
}

function switchToOutputPage() {
    currentPage = 'output';
    document.getElementById('dashboard-page').style.display = 'none';
    document.getElementById('output-page').style.display = 'block';

    // Fetch fresh data immediately
    fetchOutputData();

    // Start output polling
    if (outputPollTimer) clearInterval(outputPollTimer);
    outputPollTimer = setInterval(fetchOutputData, outputConfig.output_poll_seconds * 1000);
}

function switchToDashboardPage() {
    currentPage = 'dashboard';
    document.getElementById('output-page').style.display = 'none';
    document.getElementById('dashboard-page').style.display = 'block';

    // Stop output polling
    if (outputPollTimer) {
        clearInterval(outputPollTimer);
        outputPollTimer = null;
    }

    // Re-fetch dashboard
    fetchStatus();
}

// ===================== OUTPUT PAGE DATA =====================

async function fetchOutputData() {
    try {
        const resp = await fetch('/api/output-summary');
        if (!resp.ok) throw new Error('HTTP ' + resp.status);
        const data = await resp.json();
        renderOutputPage(data);
    } catch (e) {
        console.error('Output fetch error:', e);
    }
}

function renderOutputPage(data) {
    // Update header
    if (data.last_update) {
        const dt = new Date(data.last_update);
        document.getElementById('output-last-update').textContent = dt.toLocaleString();
    }
    const target = data.target || 2000;
    const todayTotal = data.daily_total_produced || 0;
    const gap = data.gap || 0;
    const projected = data.projected_end || 0;

    document.getElementById('output-target').textContent = target;
    document.getElementById('output-today-total').textContent = todayTotal;

    // Gap display
    const gapEl = document.getElementById('output-gap');
    if (gap >= 0) {
        gapEl.textContent = '+' + gap;
        gapEl.className = 'output-gap-positive';
    } else {
        gapEl.textContent = String(gap);
        gapEl.className = 'output-gap-negative';
    }

    // Forecast with emoticon
    const forecastEl = document.getElementById('output-forecast');
    if (projected <= 0) {
        forecastEl.textContent = '-- ';
        forecastEl.className = 'output-forecast-neutral';
    } else if (projected >= target * 1.1) {
        forecastEl.textContent = projected + ' \u{1F929}';  // star-struck
        forecastEl.className = 'output-forecast-great';
    } else if (projected >= target) {
        forecastEl.textContent = projected + ' \u{1F60A}';  // smiling
        forecastEl.className = 'output-forecast-good';
    } else if (projected >= target * 0.9) {
        forecastEl.textContent = projected + ' \u{1F615}';  // confused/worried
        forecastEl.className = 'output-forecast-warning';
    } else {
        forecastEl.textContent = projected + ' \u{1F61F}';  // worried
        forecastEl.className = 'output-forecast-bad';
    }

    // Render phase table
    renderPhaseTable(data.phases || []);

    // Render chart
    renderOutputChart(data.history || [], data.target || 2000);
}

function renderPhaseTable(phases) {
    const headerRow = document.getElementById('output-phases-header');
    const plannedRow = document.getElementById('output-phases-planned');
    const producedRow = document.getElementById('output-phases-produced');

    if (phases.length === 0) {
        headerRow.innerHTML = '<th>No data</th>';
        plannedRow.innerHTML = '<td>-</td>';
        producedRow.innerHTML = '<td>-</td>';
        return;
    }

    // Build header
    let headerHtml = '<th class="phase-label-col">Qty</th>';
    for (const p of phases) {
        headerHtml += `<th class="phase-col">${escHtml(p.phase)}</th>`;
    }
    headerRow.innerHTML = headerHtml;

    // Build planned row
    let plannedHtml = '<td class="phase-label-cell">Planned</td>';
    for (const p of phases) {
        plannedHtml += `<td class="phase-value-cell">${p.planned}</td>`;
    }
    plannedRow.innerHTML = plannedHtml;

    // Build produced row
    let producedHtml = '<td class="phase-label-cell">Produced</td>';
    for (const p of phases) {
        const cls = p.produced >= p.planned && p.planned > 0 ? 'phase-ok' : (p.planned > 0 ? 'phase-behind' : '');
        producedHtml += `<td class="phase-value-cell ${cls}">${p.produced}</td>`;
    }
    producedRow.innerHTML = producedHtml;
}

function renderOutputChart(history, target) {
    const ctx = document.getElementById('output-chart');
    if (!ctx) return;

    // History is already filtered to working days only, with week numbers
    // Each entry: {day, produced, week}

    // Build multi-line labels: day number on top, "weekNN" below (only on first day of each week)
    const labels = [];
    let lastWeek = null;
    for (const h of history) {
        if (h.week !== lastWeek) {
            labels.push([String(h.day), 'week' + String(h.week).padStart(2, '0')]);
            lastWeek = h.week;
        } else {
            labels.push([String(h.day), '']);
        }
    }

    // Build realized data (continuous, no nulls)
    const realizedData = history.map(h => h.produced);

    // Build cumulative average
    const averageData = [];
    let cumSum = 0;
    let cumCount = 0;
    for (const h of history) {
        cumSum += h.produced;
        cumCount++;
        averageData.push(Math.round(cumSum / cumCount));
    }

    // Target line (constant)
    const targetData = history.map(() => target);

    // Destroy previous chart if exists
    if (outputChart) {
        outputChart.destroy();
    }

    outputChart = new Chart(ctx, {
        type: 'line',
        data: {
            labels: labels,
            datasets: [
                {
                    label: 'REALIZAT',
                    data: realizedData,
                    borderColor: '#2ecc71',
                    backgroundColor: 'rgba(46, 204, 113, 0.1)',
                    borderWidth: 3,
                    pointRadius: 4,
                    pointBackgroundColor: '#2ecc71',
                    tension: 0.3,
                    spanGaps: true
                },
                {
                    label: 'Target',
                    data: targetData,
                    borderColor: '#e74c3c',
                    borderWidth: 2,
                    pointRadius: 0,
                    fill: false
                },
                {
                    label: 'AVERAGE',
                    data: averageData,
                    borderColor: '#1a3a6e',
                    borderWidth: 2,
                    pointRadius: 0,
                    tension: 0.3,
                    spanGaps: true
                }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                title: {
                    display: true,
                    text: outputConfig.chart_title || 'Output VDW RO',
                    color: '#ccd6f6',
                    font: { size: 18, weight: 'bold' }
                },
                legend: {
                    position: 'bottom',
                    labels: {
                        color: '#8892b0',
                        font: { size: 13 },
                        padding: 20,
                        usePointStyle: true
                    }
                }
            },
            scales: {
                x: {
                    title: {
                        display: false
                    },
                    ticks: {
                        color: '#8892b0',
                        font: { size: 12 },
                        maxRotation: 0,
                        autoSkip: false
                    },
                    grid: {
                        color: (context) => {
                            // Stronger grid line at week boundaries
                            const label = labels[context.tick?.value];
                            if (label && label[1] && label[1].startsWith('week')) {
                                return 'rgba(255,255,255,0.15)';
                            }
                            return 'rgba(255,255,255,0.05)';
                        }
                    }
                },
                y: {
                    min: 0,
                    suggestedMax: target * 1.3,
                    ticks: {
                        color: '#8892b0',
                        font: { size: 12 },
                        stepSize: 500
                    },
                    grid: {
                        color: 'rgba(255,255,255,0.08)'
                    }
                }
            },
            interaction: {
                intersect: false,
                mode: 'index'
            }
        }
    });
}

// ===================== UTILITIES =====================

function escHtml(str) {
    if (!str) return '';
    const div = document.createElement('div');
    div.textContent = str;
    return div.innerHTML;
}

async function runNow() {
    const btn = document.getElementById('btn-run-now');
    const status = document.getElementById('run-status');
    btn.disabled = true;
    btn.textContent = 'Running...';
    status.textContent = '';

    try {
        const resp = await fetch('/api/run-now', { method: 'POST' });
        if (resp.ok) {
            status.textContent = 'Cycle started';
            setTimeout(fetchStatus, 5000);
            setTimeout(fetchStatus, 15000);
        } else {
            status.textContent = 'Error starting cycle';
        }
    } catch (e) {
        status.textContent = 'Network error';
    }

    setTimeout(() => {
        btn.disabled = false;
        btn.textContent = 'Run Now';
        status.textContent = '';
    }, 20000);
}

function toggleErrors() {
    const list = document.getElementById('error-list');
    const toggle = document.getElementById('error-toggle');
    if (list.style.display === 'none') {
        list.style.display = 'block';
        toggle.innerHTML = '&#9650;';
    } else {
        list.style.display = 'none';
        toggle.innerHTML = '&#9660;';
    }
}

// Live clock (updates both pages)
function updateClock() {
    const now = new Date();
    const h = String(now.getHours()).padStart(2, '0');
    const m = String(now.getMinutes()).padStart(2, '0');
    const s = String(now.getSeconds()).padStart(2, '0');
    const timeStr = `${h}:${m}:${s}`;
    document.getElementById('live-clock').textContent = timeStr;
    const clockOutput = document.getElementById('live-clock-output');
    if (clockOutput) clockOutput.textContent = timeStr;
}

// ===================== INIT =====================

document.addEventListener('DOMContentLoaded', async () => {
    // Load output config
    try {
        const resp = await fetch('/api/output-config');
        if (resp.ok) {
            const cfg = await resp.json();
            Object.assign(outputConfig, cfg);
        }
    } catch (e) {
        console.warn('Could not load output config, using defaults');
    }

    // Set chart title
    const titleEl = document.getElementById('output-title');
    if (titleEl && outputConfig.chart_title) {
        titleEl.textContent = outputConfig.chart_title;
    }

    fetchStatus();
    pollTimer = setInterval(fetchStatus, 15000);
    updateClock();
    setInterval(updateClock, 1000);

    // Mouse activity tracker per inactivity timeout
    document.addEventListener('mousemove', onUserActivity);
    document.addEventListener('mousedown', onUserActivity);
    document.addEventListener('wheel', onUserActivity);

    updateScrollButton();

    // Start page rotation
    startRotation();
});
