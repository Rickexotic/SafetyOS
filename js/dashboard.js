let riskRadarChart = null;
let severityChart = null;

const SEVERITY_LEVELS = [
    "Low",
    "Medium",
    "High",
    "Critical"
];

window.addEventListener(
    "load",
    async () => {
        await msalInstance.initialize();

        const accounts =
            msalInstance.getAllAccounts();

        if (accounts.length > 0) {
            activeAccount =
                accounts[0];

            await loadProfilePhoto();
        }

        await loadDashboard();
    }
);

async function loadDashboard() {
    showLoadingState(true);
    clearErrorState();

    try {
        const token =
            await getAccessToken();

        const response =
            await fetch(
                `https://graph.microsoft.com/v1.0/sites/46y2.sharepoint.com:/sites/SaaS_OHS:/lists/${CONFIG.incidentsListId}/items?$expand=fields&$top=5000`,
                {
                    headers: {
                        Authorization:
                            "Bearer " + token
                    }
                }
            );

        if (!response.ok) {
            throw new Error(
                `Graph request failed (${response.status})`
            );
        }

        const data =
            await response.json();

        const incidents =
            Array.isArray(data.value)
                ? data.value
                : [];

        buildKPIs(incidents);
        buildRiskRadar(incidents);
        buildSeverityChart(incidents);
        buildTopSites(incidents);
    } catch (error) {
        console.error(
            "Failed to load dashboard:",
            error
        );

        buildKPIs([]);
        buildRiskRadar([]);
        buildSeverityChart([]);
        buildTopSites([]);

        showErrorState(
            "Unable to load dashboard data."
        );
    } finally {
        showLoadingState(false);
    }
}

function buildKPIs(incidents) {
    const total = incidents.length;

    const critical =
        incidents.filter(
            i =>
                getSeverity(i) === "Critical"
        ).length;

    const closed =
        incidents.filter(
            i =>
                getStatus(i) === "Closed"
        ).length;

    const open =
        incidents.filter(
            i =>
                getStatus(i) !== "Closed"
        ).length;

    setText("kpiTotal", total);
    setText("kpiCritical", critical);
    setText("kpiClosed", closed);
    setText("kpiOpen", open);
}

function buildRiskRadar(incidents) {
    const canvas =
        document.getElementById("riskRadar");

    if (!canvas || typeof Chart === "undefined") {
        return;
    }

    const total = incidents.length;

    const severeCount =
        incidents.filter(i => {
            const severity =
                getSeverity(i);
            return (
                severity === "High" ||
                severity === "Critical"
            );
        }).length;

    const openCount =
        incidents.filter(
            i =>
                getStatus(i) !== "Closed"
        ).length;

    const closedCount =
        incidents.filter(
            i =>
                getStatus(i) === "Closed"
        ).length;

    const unresolvedSevereCount =
        incidents.filter(i => {
            const severity =
                getSeverity(i);
            const status =
                getStatus(i);

            return (
                (severity === "High" ||
                    severity === "Critical") &&
                status !== "Closed"
            );
        }).length;

    const siteCounts = {};

    incidents.forEach(i => {
        const site =
            getSite(i) || "Unknown";

        siteCounts[site] =
            (siteCounts[site] || 0) + 1;
    });

    const maxSite =
        Math.max(
            0,
            ...Object.values(siteCounts)
        );

    const safePercent = value =>
        total > 0
            ? Math.round((value / total) * 100)
            : 0;

    const values = [
        safePercent(severeCount),
        safePercent(openCount),
        safePercent(maxSite),
        safePercent(unresolvedSevereCount),
        total > 0
            ? Math.round(
                (1 -
                    closedCount / total) *
                    100
            )
            : 0
    ];

    if (riskRadarChart) {
        riskRadarChart.destroy();
    }

    riskRadarChart =
        new Chart(canvas, {
            type: "radar",
            data: {
                labels: [
                    "Severe Incidents",
                    "Open Cases",
                    "Site Concentration",
                    "Unresolved Severe",
                    "Closure Gap"
                ],
                datasets: [{
                    data: values,
                    borderColor: "#60a5fa",
                    backgroundColor:
                        "rgba(96,165,250,.20)",
                    pointBackgroundColor:
                        "#60a5fa",
                    borderWidth: 2
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: true,
                plugins: {
                    legend: {
                        display: false
                    }
                },
                scales: {
                    r: {
                        suggestedMin: 0,
                        suggestedMax: 100,
                        angleLines: {
                            color:
                                "rgba(255,255,255,.08)"
                        },
                        grid: {
                            color:
                                "rgba(255,255,255,.08)"
                        },
                        pointLabels: {
                            color: "#cbd5e1"
                        },
                        ticks: {
                            color: "#94a3b8",
                            backdropColor:
                                "transparent",
                            stepSize: 20
                        }
                    }
                }
            }
        });
}

function buildSeverityChart(incidents) {
    const canvas =
        document.getElementById("severityChart");

    if (!canvas || typeof Chart === "undefined") {
        return;
    }

    const counts = {
        Low: 0,
        Medium: 0,
        High: 0,
        Critical: 0,
        Unknown: 0
    };

    incidents.forEach(i => {
        const severity =
            getSeverity(i);

        if (counts[severity] !== undefined) {
            counts[severity]++;
        } else {
            counts.Unknown++;
        }
    });

    let labels =
        [...SEVERITY_LEVELS];

    let data =
        labels.map(label => counts[label]);

    let colors = [
        "#22c55e",
        "#eab308",
        "#f97316",
        "#ef4444"
    ];

    if (counts.Unknown > 0) {
        labels.push("Unknown");
        data.push(counts.Unknown);
        colors.push("#64748b");
    }

    const total =
        data.reduce((sum, value) => sum + value, 0);

    if (total === 0) {
        labels = ["No Data"];
        data = [1];
        colors = ["#334155"];
    }

    if (severityChart) {
        severityChart.destroy();
    }

    severityChart =
        new Chart(canvas, {
            type: "doughnut",
            data: {
                labels,
                datasets: [{
                    data,
                    backgroundColor: colors,
                    borderWidth: 0
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: true,
                plugins: {
                    legend: {
                        labels: {
                            color: "#cbd5e1"
                        }
                    }
                }
            }
        });
}

function buildTopSites(incidents) {
    const container =
        document.getElementById(
            "siteSummary"
        );

    if (!container) {
        return;
    }

    const sites = {};

    incidents.forEach(i => {
        const site =
            getSite(i) || "Unknown";

        sites[site] =
            (sites[site] || 0) + 1;
    });

    const topSites =
        Object.entries(sites)
            .sort(
                (a, b) =>
                    b[1] - a[1]
            )
            .slice(0, 5);

    if (topSites.length === 0) {
        container.innerHTML =
            '<div class="chart-label">No site data available</div>';
        return;
    }

    container.innerHTML =
        topSites
            .map(
                ([site, count]) => `
                    <div class="chart-header">
                        <span class="chart-label">
                            ${escapeHtml(site)}
                        </span>
                        <span class="chart-value">
                            ${count}
                        </span>
                    </div>
                `
            )
            .join("");
}

function getFields(item) {
    return (
        (item && item.fields) ||
        {}
    );
}

function normalizeKey(value) {
    return String(value)
        .replace(/[^a-z0-9]/gi, "")
        .toLowerCase();
}

function getFieldValue(item, names) {
    const fields =
        getFields(item);

    for (const name of names) {
        const direct =
            fields[name];

        if (
            direct !== undefined &&
            direct !== null &&
            direct !== ""
        ) {
            return direct;
        }
    }

    const normalizedLookup =
        Object.keys(fields).reduce(
            (lookup, key) => {
                lookup[
                    normalizeKey(key)
                ] = fields[key];

                return lookup;
            },
            {}
        );

    for (const name of names) {
        const value =
            normalizedLookup[
                normalizeKey(name)
            ];

        if (
            value !== undefined &&
            value !== null &&
            value !== ""
        ) {
            return value;
        }
    }

    return "";
}

function toSafeText(value) {
    if (
        value === null ||
        value === undefined
    ) {
        return "";
    }

    if (Array.isArray(value)) {
        return value.join(", ");
    }

    return String(value).trim();
}

function getSeverity(item) {
    const raw =
        toSafeText(
            getFieldValue(item, [
                "Severity"
            ])
        ).toLowerCase();

    if (raw.includes("crit")) {
        return "Critical";
    }

    if (raw.includes("high")) {
        return "High";
    }

    if (raw.includes("med")) {
        return "Medium";
    }

    if (raw.includes("low")) {
        return "Low";
    }

    return "Unknown";
}

function getStatus(item) {
    const raw =
        toSafeText(
            getFieldValue(item, [
                "Status"
            ])
        ).toLowerCase();

    if (raw.includes("close")) {
        return "Closed";
    }

    if (raw.includes("open")) {
        return "Open";
    }

    return "Unknown";
}

function getSite(item) {
    return toSafeText(
        getFieldValue(item, [
            "Site"
        ])
    );
}

function setText(id, value) {
    const element =
        document.getElementById(id);

    if (element) {
        element.innerText =
            String(value);
    }
}

function escapeHtml(text) {
    const map = {
        "&": "&amp;",
        "<": "&lt;",
        ">": "&gt;",
        '"': "&quot;",
        "'": "&#39;"
    };

    return toSafeText(text)
        .replace(
            /[&<>"']/g,
            m => map[m]
        );
}

function showLoadingState(isLoading) {
    const loadingId =
        "dashboardLoading";

    const existing =
        document.getElementById(
            loadingId
        );

    if (!isLoading) {
        existing?.remove();
        return;
    }

    if (existing) {
        existing.classList.remove(
            "d-none"
        );
        return;
    }

    const cardBody =
        document.querySelector(
            ".card-body"
        );

    if (!cardBody) {
        return;
    }

    const loading =
        document.createElement("div");

    loading.id = loadingId;
    loading.className =
        "small text-muted mb-3";
    loading.innerText =
        "Loading dashboard...";

    cardBody.prepend(loading);
}

function showErrorState(message) {
    const errorId =
        "dashboardError";

    let errorEl =
        document.getElementById(
            errorId
        );

    if (!errorEl) {
        const cardBody =
            document.querySelector(
                ".card-body"
            );

        if (!cardBody) {
            return;
        }

        errorEl =
            document.createElement("div");

        errorEl.id = errorId;
        errorEl.className =
            "alert alert-warning py-2 px-3 mb-3";

        cardBody.prepend(errorEl);
    }

    errorEl.innerText =
        toSafeText(message) ||
        "Something went wrong.";
}

function clearErrorState() {
    document
        .getElementById(
            "dashboardError"
        )
        ?.remove();
}
