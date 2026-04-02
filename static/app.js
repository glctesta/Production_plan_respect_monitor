let pollTimer = null;
let scrollTimer = null;
let currentScrollRow = 0;
let totalRows = 0;

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

        const planLabel = r.is_out_of_plan ? '<em>NOT IN PLAN</em>' : r.planned_qty_day;
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

    // Start or restart auto-scroll
    setupAutoScroll();

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

// --- Auto-scroll logic ---
function setupAutoScroll() {
    const wrapper = document.getElementById('table-wrapper');
    const tbody = document.getElementById('table-body');

    // Check if content overflows
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
    const wrapper = document.getElementById('table-wrapper');
    const tbody = document.getElementById('table-body');
    const rows = tbody.querySelectorAll('tr');

    if (rows.length === 0) return;

    currentScrollRow++;

    // If we've scrolled past the end, reset to top
    if (currentScrollRow >= rows.length) {
        currentScrollRow = 0;
        wrapper.scrollTo({ top: 0, behavior: 'smooth' });
        return;
    }

    // Scroll so that currentScrollRow is at top of visible area (below sticky header)
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

// Pause auto-scroll on hover, resume on leave
document.addEventListener('DOMContentLoaded', () => {
    const wrapper = document.getElementById('table-wrapper');
    wrapper.addEventListener('mouseenter', () => {
        if (scrollTimer) {
            clearInterval(scrollTimer);
            scrollTimer = null;
        }
    });
    wrapper.addEventListener('mouseleave', () => {
        if (totalRows > 0 && wrapper.scrollHeight > wrapper.clientHeight) {
            scrollTimer = setInterval(scrollOneRow, 5000);
        }
    });
});

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

// Init
document.addEventListener('DOMContentLoaded', () => {
    fetchStatus();
    pollTimer = setInterval(fetchStatus, 15000);
});
