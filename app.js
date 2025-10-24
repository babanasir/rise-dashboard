const SHEET_URL = "https://docs.google.com/spreadsheets/d/1T0ikFqhzx61GromLJlZZj8SxBuMEtGTe6ph0t6EunpD/export?format=csv&gid=646193820";

const STATUS_COLORS = {
    "Not Started": { bg: "#d9d9d9", text: "#0f172a" },
    Drafting: { bg: "#fff2cc", text: "#0f172a" },
    "1st Draft Ready": { bg: "#fce4d6", text: "#0f172a" },
    "Internal Review Ongoing": { bg: "#ffd966", text: "#0f172a" },
    "Internal Review Completed": { bg: "#c6efce", text: "#0f172a" },
    "Expert Validtion Ongoing": { bg: "#bdd7ee", text: "#0f172a" },
    "Expert Validation Completed": { bg: "#9bc2e6", text: "#0f172a" },
    "Approved for Testing": { bg: "#00b0f0", text: "#ffffff" },
    "Field Testing Completed": { bg: "#d9d2e9", text: "#0f172a" },
    "Post-Testing Review": { bg: "#b4a7d6", text: "#0f172a" },
    "Final Approval": { bg: "#548235", text: "#ffffff" }
};

const STATUS_SEQUENCE = [
    "Not Started",
    "Drafting",
    "1st Draft Ready",
    "Internal Review Ongoing",
    "Internal Review Completed",
    "Expert Validtion Ongoing",
    "Expert Validation Completed",
    "Approved for Testing",
    "Field Testing Completed",
    "Post-Testing Review",
    "Final Approval"
];

const metricsContainer = document.getElementById("metrics");
const statusTableBody = document.querySelector("#status-table tbody");
const lastUpdatedEl = document.getElementById("lastUpdated");
const errorBanner = document.getElementById("error-banner");
const refreshButton = document.getElementById("refreshButton");

refreshButton.addEventListener("click", () => {
    loadDashboard();
});

document.addEventListener("DOMContentLoaded", () => {
    loadDashboard();
});

async function loadDashboard() {
    setLoadingState(true);
    clearError();

    try {
        const rows = await fetchSheetRows();
        const lessonRows = rows.filter((row) => Number.isFinite(parseLesson(row["Lesson"])));

        if (!lessonRows.length) {
            renderEmptyState();
            updateTimestamp();
            return;
        }

        renderDashboard(lessonRows);
        updateTimestamp();
    } catch (error) {
        console.error(error);
        showError(error.message || "Unable to load dashboard data.");
        renderFallbackState();
    } finally {
        setLoadingState(false);
    }
}

async function fetchSheetRows() {
    const response = await fetch(`${SHEET_URL}&_ts=${Date.now()}`);

    if (!response.ok) {
        throw new Error("Unable to fetch the Google Sheet. Please try again later.");
    }

    const text = await response.text();
    const rows = parseCSV(text);

    if (!rows.length) {
        return [];
    }

    return rows;
}

function renderDashboard(rows) {
    const totalLessons = rows.length;
    const completed = rows.filter((row) => normalizeStatus(row["Status"]) === "Final Approval").length;
    const notStarted = rows.filter((row) => normalizeStatus(row["Status"]) === "Not Started").length;
    const inProgress = Math.max(totalLessons - completed - notStarted, 0);
    const percentComplete = totalLessons > 0 ? completed / totalLessons : 0;

    renderMetrics({ totalLessons, completed, notStarted, inProgress, percentComplete });
    renderStatusTable(rows);
}

function renderMetrics({ totalLessons, completed, notStarted, inProgress, percentComplete }) {
    metricsContainer.innerHTML = "";

    const metricData = [
        { label: "Total Lessons", value: totalLessons.toLocaleString() },
        { label: "Completed (Final Approval)", value: completed.toLocaleString() },
        { label: "Not Started", value: notStarted.toLocaleString() },
        { label: "In Progress", value: inProgress.toLocaleString() }
    ];

    metricData.forEach(({ label, value }) => {
        const card = document.createElement("article");
        card.className = "metric-card";
        card.innerHTML = `
            <h3>${label}</h3>
            <p class="metric-value">${value}</p>
        `;
        metricsContainer.appendChild(card);
    });

    const progressCard = document.createElement("article");
    progressCard.className = "metric-card metric-progress";
    progressCard.innerHTML = `
        <h3>Percent Complete</h3>
        <p class="metric-value">${formatPercent(percentComplete)}</p>
        <pre class="progress-bar" aria-label="${formatPercent(percentComplete)} complete">${generateProgressBar(percentComplete)}</pre>
    `;
    metricsContainer.appendChild(progressCard);
}

function renderStatusTable(rows) {
    statusTableBody.innerHTML = "";

    const counts = rows.reduce((acc, row) => {
        const status = normalizeStatus(row["Status"]);
        if (!status) {
            return acc;
        }
        acc[status] = (acc[status] || 0) + 1;
        return acc;
    }, {});

    const dynamicStatuses = Object.keys(counts).filter((status) => !STATUS_SEQUENCE.includes(status)).sort();
    const statuses = [...STATUS_SEQUENCE, ...dynamicStatuses];

    if (!statuses.length) {
        const emptyRow = document.createElement("tr");
        emptyRow.innerHTML = '<td colspan="2">No lesson statuses available.</td>';
        statusTableBody.appendChild(emptyRow);
        return;
    }

    statuses.forEach((status) => {
        const count = counts[status] || 0;
        const colors = STATUS_COLORS[status] || { bg: "#e5e7eb", text: "#0f172a" };

        const row = document.createElement("tr");
        row.dataset.status = status;

        const statusCell = document.createElement("td");
        statusCell.className = "status-cell";
        statusCell.textContent = status;
        statusCell.style.backgroundColor = colors.bg;
        statusCell.style.color = colors.text;

        const countCell = document.createElement("td");
        countCell.className = "status-count";
        countCell.textContent = count.toLocaleString();

        row.append(statusCell, countCell);
        statusTableBody.appendChild(row);
    });
}

function renderEmptyState() {
    metricsContainer.innerHTML = '<p class="empty-state">No lessons found in the sheet yet. Once rows are added, metrics will appear here.</p>';
    statusTableBody.innerHTML = '<tr><td colspan="2">No lesson statuses available.</td></tr>';
}

function renderFallbackState() {
    metricsContainer.innerHTML = '<p class="empty-state">Metrics unavailable.</p>';
    statusTableBody.innerHTML = "";
}

function setLoadingState(isLoading) {
    refreshButton.disabled = isLoading;
    refreshButton.textContent = isLoading ? "Loading…" : "Refresh";
}

function showError(message) {
    errorBanner.textContent = message;
    errorBanner.hidden = false;
}

function clearError() {
    errorBanner.hidden = true;
    errorBanner.textContent = "";
}

function updateTimestamp() {
    const now = new Date();
    lastUpdatedEl.textContent = now.toLocaleString();
}

function generateProgressBar(percent) {
    const segments = 20;
    const filled = Math.round(percent * segments);
    const empty = segments - filled;
    return `${"█".repeat(filled)}${"░".repeat(Math.max(empty, 0))}`;
}

function parseLesson(value) {
    const parsed = Number.parseInt(value, 10);
    return Number.isFinite(parsed) ? parsed : NaN;
}

function normalizeStatus(value = "") {
    return String(value).trim();
}

function formatPercent(value) {
    return `${Math.round(value * 100)}%`;
}

function parseCSV(text) {
    if (!text || !text.trim()) {
        return [];
    }

    const rows = [];
    let current = [];
    let currentValue = "";
    let inQuotes = false;

    for (let i = 0; i < text.length; i++) {
        const char = text[i];
        const nextChar = text[i + 1];

        if (char === '"') {
            if (inQuotes && nextChar === '"') {
                currentValue += '"';
                i++;
            } else {
                inQuotes = !inQuotes;
            }
        } else if (char === "," && !inQuotes) {
            current.push(currentValue);
            currentValue = "";
        } else if ((char === "\n" || char === "\r") && !inQuotes) {
            if (char === "\r" && nextChar === "\n") {
                i++;
            }
            current.push(currentValue);
            rows.push(current);
            current = [];
            currentValue = "";
        } else {
            currentValue += char;
        }
    }

    if (currentValue !== "" || current.length) {
        current.push(currentValue);
        rows.push(current);
    }

    if (!rows.length) {
        return [];
    }

    const headers = rows[0].map((value) => value.trim());
    const dataRows = rows.slice(1);

    return dataRows
        .filter((row) => row.some((cell) => cell && cell.trim() !== ""))
        .map((row) => {
            const obj = {};
            headers.forEach((header, index) => {
                obj[header] = (row[index] || "").trim();
            });
            return obj;
        });
}
