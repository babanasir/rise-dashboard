const SHEET_CSV_URL = 'https://docs.google.com/spreadsheets/d/1T0ikFqhzx61GromLJlZZj8SxBuMEtGTe6ph0t6EunpD/export?format=csv&gid=646193820';

const STATUS_STYLES = {
    'Not Started': { background: '#d9d9d9', color: '#0f172a' },
    Drafting: { background: '#fff2cc', color: '#0f172a' },
    '1st Draft Ready': { background: '#fce4d6', color: '#0f172a' },
    'Internal Review Ongoing': { background: '#ffd966', color: '#0f172a' },
    'Internal Review Completed': { background: '#c6efce', color: '#0f172a' },
    'Expert Validtion Ongoing': { background: '#bdd7ee', color: '#0f172a' },
    'Expert Validation Completed': { background: '#9bc2e6', color: '#0f172a' },
    'Approved for Testing': { background: '#00b0f0', color: '#ffffff' },
    'Field Testing Completed': { background: '#d9d2e9', color: '#0f172a' },
    'Post-Testing Review': { background: '#b4a7d6', color: '#0f172a' },
    'Final Approval': { background: '#548235', color: '#ffffff' }
};

const EXPECTED_HEADERS = [
    'Material Type',
    'Language',
    'Grade Level',
    'Unit/Week',
    'Lesson',
    'Title/Description',
    'Status',
    'Comments/Notes',
    'Current Owner',
    'Due Date',
    'Priority'
];

const metricsContainer = document.getElementById('metrics');
const progressBar = document.getElementById('progressBar');
const statusTableBody = document.querySelector('#status-table tbody');
const lastUpdatedEl = document.getElementById('lastUpdated');
const errorBanner = document.getElementById('errorBanner');
const refreshButton = document.getElementById('refreshButton');

refreshButton.addEventListener('click', () => {
    loadDashboard();
});

document.addEventListener('DOMContentLoaded', () => {
    loadDashboard();
});

async function loadDashboard() {
    toggleLoadingState(true);
    clearError();

    try {
        const rows = await fetchSheetRows();
        renderDashboard(rows);
    } catch (error) {
        console.error(error);
        showError(error.message || 'Something went wrong.');
        renderFallbackState();
    } finally {
        toggleLoadingState(false);
    }
}

function toggleLoadingState(isLoading) {
    refreshButton.disabled = isLoading;
    refreshButton.textContent = isLoading ? 'Loading…' : 'Refresh';
}

async function fetchSheetRows() {
    const response = await fetch(`${SHEET_CSV_URL}&_ts=${Date.now()}`);

    if (!response.ok) {
        throw new Error('Unable to fetch the Google Sheet. Please try again later.');
    }

    const text = await response.text();
    const parsed = parseCSV(text);

    if (!parsed.length) {
        throw new Error('The CSV did not contain any data.');
    }

    const headers = parsed[0];

    if (!arraysEqual(headers, EXPECTED_HEADERS)) {
        throw new Error('Unexpected CSV format. Please verify the sheet columns.');
    }

    return parsed.slice(1).map((row) => toObject(headers, row));
}

function renderDashboard(rows) {
    const totalLessons = rows.reduce((count, row) => {
        const lesson = Number(row['Lesson']);
        return Number.isFinite(lesson) ? count + 1 : count;
    }, 0);

    const completed = rows.filter((row) => row['Status'].trim() === 'Final Approval').length;
    const notStarted = rows.filter((row) => row['Status'].trim() === 'Not Started').length;
    const inProgress = Math.max(totalLessons - completed - notStarted, 0);
    const percentComplete = totalLessons > 0 ? completed / totalLessons : 0;

    renderMetrics({
        totalLessons,
        completed,
        notStarted,
        inProgress,
        percentComplete
    });

    renderProgress(percentComplete);
    renderStatusTable(rows);
    updateTimestamp();
}

function renderMetrics(metrics) {
    const { totalLessons, completed, notStarted, inProgress, percentComplete } = metrics;

    metricsContainer.innerHTML = '';

    const metricData = [
        { label: 'Total Lessons', value: totalLessons },
        { label: 'Completed (Final Approval)', value: completed },
        { label: 'Not Started', value: notStarted },
        { label: 'In Progress', value: inProgress },
        { label: '% Complete', value: formatPercent(percentComplete) }
    ];

    metricData.forEach((metric) => {
        const card = document.createElement('article');
        card.className = 'metric-card';

        const label = document.createElement('h3');
        label.textContent = metric.label;

        const value = document.createElement('p');
        value.className = 'metric-value';
        value.textContent = metric.value;

        card.append(label, value);
        metricsContainer.appendChild(card);
    });
}

function renderProgress(percent) {
    const blocks = 20;
    const filledBlocks = Math.round(percent * blocks);
    const emptyBlocks = Math.max(blocks - filledBlocks, 0);
    const bar = `${'█'.repeat(filledBlocks)}${'░'.repeat(emptyBlocks)}`;
    const formatted = `${bar} ${formatPercent(percent)}`;
    progressBar.textContent = formatted;
}

function renderStatusTable(rows) {
    statusTableBody.innerHTML = '';

    const counts = rows.reduce((acc, row) => {
        const status = row['Status'].trim();
        if (!status) {
            return acc;
        }
        acc[status] = (acc[status] || 0) + 1;
        return acc;
    }, {});

    const sortedStatuses = Object.entries(counts).sort((a, b) => b[1] - a[1]);

    sortedStatuses.forEach(([status, count]) => {
        const tr = document.createElement('tr');

        const statusCell = document.createElement('td');
        const pill = document.createElement('span');
        pill.className = 'status-pill';
        pill.textContent = status;
        const statusStyle = STATUS_STYLES[status];
        if (statusStyle) {
            pill.style.backgroundColor = statusStyle.background;
            pill.style.color = statusStyle.color;
        }
        statusCell.appendChild(pill);

        const countCell = document.createElement('td');
        countCell.textContent = count;

        tr.append(statusCell, countCell);
        statusTableBody.appendChild(tr);
    });
}

function updateTimestamp() {
    const now = new Date();
    lastUpdatedEl.textContent = now.toLocaleString();
}

function renderFallbackState() {
    metricsContainer.innerHTML = '<p class="empty-state">Metrics unavailable.</p>';
    progressBar.textContent = '';
    statusTableBody.innerHTML = '';
}

function showError(message) {
    errorBanner.textContent = message;
    errorBanner.hidden = false;
}

function clearError() {
    errorBanner.hidden = true;
    errorBanner.textContent = '';
}

function parseCSV(text) {
    const rows = [];
    let currentRow = [];
    let currentValue = '';
    let inQuotes = false;

    for (let i = 0; i < text.length; i++) {
        const char = text[i];
        const nextChar = text[i + 1];

        if (char === '"') {
            if (inQuotes && nextChar === '"') {
                currentValue += '"';
                i++; // Skip escaped quote
            } else {
                inQuotes = !inQuotes;
            }
        } else if (char === ',' && !inQuotes) {
            currentRow.push(currentValue);
            currentValue = '';
        } else if ((char === '\n' || char === '\r') && !inQuotes) {
            if (currentValue !== '' || currentRow.length) {
                currentRow.push(currentValue);
                rows.push(currentRow);
                currentRow = [];
                currentValue = '';
            }
            if (char === '\r' && nextChar === '\n') {
                i++;
            }
        } else {
            currentValue += char;
        }
    }

    if (currentValue !== '' || currentRow.length) {
        currentRow.push(currentValue);
        rows.push(currentRow);
    }

    return rows.map((row) => row.map((value) => value.trim()));
}

function toObject(headers, row) {
    const obj = {};
    headers.forEach((header, index) => {
        obj[header] = row[index] ?? '';
    });
    return obj;
}

function arraysEqual(a, b) {
    if (a.length !== b.length) {
        return false;
    }
    return a.every((value, index) => value === b[index]);
}

function formatPercent(value) {
    return `${(value * 100).toFixed(0)}%`;
}
