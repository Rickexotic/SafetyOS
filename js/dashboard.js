window.addEventListener(
    "load",
    async () => {
        try {
            await msalInstance.initialize();

            const accounts =
                msalInstance.getAllAccounts();

            if (accounts.length > 0) {
                activeAccount =
                    accounts[0];

         //       await loadProfilePhoto();
            }

            await loadDashboard();
        } catch (error) {
            console.error('Dashboard initialization failed:', error);
            showErrorMessage('Failed to load dashboard. Please refresh the page.');
        }
    }
);

async function loadDashboard() {
    try {
        showLoadingState(true);

        const token =
            await getAccessToken();

        const response =
            console.log(
    "Dashboard response status:",
    response.status
);
    await fetch(
        `https://graph.microsoft.com/v1.0/sites/46y2.sharepoint.com:/sites/SaaS_OHS:/lists/${CONFIG.incidentsListId}/items?expand=fields`,
                {
                    headers: {
                        Authorization:
                            `Bearer ${token}`
                    }
                }
            );

        if (!response.ok) {
            throw new Error(`API error: ${response.status} ${response.statusText}`);
        }

        const data =
            await response.json();
        console.log(
    "Dashboard data:",
    data
);

        if (!data.value) {
            throw new Error('Invalid API response format');
        }

        const incidents =
            data.value || [];

        buildKPIs(incidents);
        buildSeveritySummary(incidents);
        buildSiteSummary(incidents);

        showLoadingState(false);
    } catch (error) {
        console.error('Dashboard load failed:', error);
        showErrorMessage('Failed to load dashboard data. Please try again.');
        showLoadingState(false);
    }
}

function buildKPIs(
    incidents
) {
    try {
        if (!Array.isArray(incidents)) {
            throw new Error('Incidents data is not an array');
        }

        document.getElementById(
            "kpiTotal"
        ).innerText =
            incidents.length;

        document.getElementById(
            "kpiCritical"
        ).innerText =
            incidents.filter(
                i =>
                    i.fields &&
                    i.fields.Severity ===
                    "Critical"
            ).length;

        document.getElementById(
            "kpiClosed"
        ).innerText =
            incidents.filter(
                i =>
                    i.fields &&
                    i.fields.Status ===
                    "Closed"
            ).length;

        document.getElementById(
            "kpiOpen"
        ).innerText =
            incidents.filter(
                i =>
                    i.fields &&
                    i.fields.Status !==
                    "Closed"
            ).length;
    } catch (error) {
        console.error('KPI build failed:', error);
        showErrorMessage('Failed to build KPIs');
    }
}

function buildSeveritySummary(incidents) {
    try {
        if (!Array.isArray(incidents)) {
            throw new Error('Incidents data is not an array');
        }

        const severityMap = {};

        incidents.forEach(incident => {
            if (incident.fields && incident.fields.Severity) {
                const severity = incident.fields.Severity;
                severityMap[severity] = (severityMap[severity] || 0) + 1;
            }
        });

        const container = document.getElementById('severitySummary');
        if (!container) {
            console.warn('Severity summary container not found');
            return;
        }

        container.innerHTML = '';

        const severities = ['Critical', 'High', 'Medium', 'Low'];

        severities.forEach(severity => {
            const count = severityMap[severity] || 0;
            const percentage = incidents.length > 0
                ? Math.round((count / incidents.length) * 100)
                : 0;

            const severityColor = {
                'Critical': '#ef4444',
                'High': '#f97316',
                'Medium': '#eab308',
                'Low': '#22c55e'
            }[severity] || '#94a3b8';

            const severityItem = document.createElement('div');
            severityItem.style.cssText = `
                display: flex;
                justify-content: space-between;
                align-items: center;
                padding: 8px 0;
                border-bottom: 1px solid rgba(255, 255, 255, 0.05);
            `;

            severityItem.innerHTML = `
                <div style="display: flex; align-items: center; gap: 8px;">
                    <div style="width: 8px; height: 8px; border-radius: 50%; background: ${severityColor};"></div>
                    <span style="color: #cbd5e1; font-size: 0.85rem;">${severity}</span>
                </div>
                <div style="display: flex; gap: 12px; align-items: center;">
                    <span style="color: #94a3b8; font-size: 0.8rem;">${count}</span>
                    <span style="color: #94a3b8; font-size: 0.75rem; min-width: 35px; text-align: right;">${percentage}%</span>
                </div>
            `;

            container.appendChild(severityItem);
        });
    } catch (error) {
        console.error('Severity summary build failed:', error);
    }
}

function buildSiteSummary(incidents) {
    try {
        if (!Array.isArray(incidents)) {
            throw new Error('Incidents data is not an array');
        }

        const siteMap = {};

        incidents.forEach(incident => {
            if (incident.fields && incident.fields.Site) {
                const site = incident.fields.Site;
                siteMap[site] = (siteMap[site] || 0) + 1;
            }
        });

        const container = document.getElementById('siteSummary');
        if (!container) {
            console.warn('Site summary container not found');
            return;
        }

        container.innerHTML = '';

        // Sort sites by incident count (descending)
        const sortedSites = Object.entries(siteMap)
            .sort((a, b) => b[1] - a[1])
            .slice(0, 5); // Show top 5

        if (sortedSites.length === 0) {
            container.innerHTML = '<p style="color: #94a3b8; font-size: 0.85rem;">No site data available</p>';
            return;
        }

        sortedSites.forEach(([site, count]) => {
            const siteItem = document.createElement('div');
            siteItem.style.cssText = `
                display: flex;
                justify-content: space-between;
                align-items: center;
                padding: 8px 0;
                border-bottom: 1px solid rgba(255, 255, 255, 0.05);
            `;

            siteItem.innerHTML = `
                <span style="color: #cbd5e1; font-size: 0.85rem;">${escapeHtml(site)}</span>
                <span style="color: #60a5fa; font-weight: 600; font-size: 0.85rem;">${count}</span>
            `;

            container.appendChild(siteItem);
        });
    } catch (error) {
        console.error('Site summary build failed:', error);
    }
}

function showLoadingState(isLoading) {
    const container = document.querySelector('.card-body');
    if (!container) return;

    if (isLoading) {
        const loadingDiv = document.createElement('div');
        loadingDiv.id = 'loadingIndicator';
        loadingDiv.style.cssText = `
            display: flex;
            justify-content: center;
            align-items: center;
            min-height: 200px;
            color: #94a3b8;
            font-size: 0.9rem;
        `;
        loadingDiv.innerHTML = '<span>Loading dashboard data...</span>';
        container.appendChild(loadingDiv);
    } else {
        const loadingDiv = document.getElementById('loadingIndicator');
        if (loadingDiv) {
            loadingDiv.remove();
        }
    }
}

function showErrorMessage(message) {
    const container = document.querySelector('.card-body');
    if (!container) return;

    const existingError = document.getElementById('errorMessage');
    if (existingError) {
        existingError.remove();
    }

    const errorDiv = document.createElement('div');
    errorDiv.id = 'errorMessage';
    errorDiv.className = 'alert alert-danger';
    errorDiv.style.cssText = `
        margin-bottom: 12px;
    `;
    errorDiv.innerText = message;
    container.insertBefore(errorDiv, container.firstChild);
}

function escapeHtml(text) {
    const map = {
        '&': '&amp;',
        '<': '&lt;',
        '>': '&gt;',
        '"': '&quot;',
        "'": '&#039;'
    };
    return text.replace(/[&<>"']/g, m => map[m]);
}
