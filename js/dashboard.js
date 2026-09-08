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

    const token =
        await getAccessToken();

    const response =
        await fetch(
            `https://graph.microsoft.com/v1.0/sites/46y2.sharepoint.com:/sites/SaaS_OHS:/lists/${CONFIG.incidentsListId}/items?expand=fields`,
            {
                headers:{
                    Authorization:
                        `Bearer ${token}`
                }
            }
        );

    const data =
        await response.json();

    const incidents =
        data.value || [];

    buildKPIs(incidents);
}

function buildKPIs(
    incidents
) {

    document.getElementById(
        "kpiTotal"
    ).innerText =
        incidents.length;

    document.getElementById(
        "kpiCritical"
    ).innerText =
        incidents.filter(
            i =>
                i.fields.Severity ===
                "Critical"
        ).length;

    document.getElementById(
        "kpiClosed"
    ).innerText =
        incidents.filter(
            i =>
                i.fields.Status ===
                "Closed"
        ).length;

    document.getElementById(
        "kpiOpen"
    ).innerText =
        incidents.filter(
            i =>
                i.fields.Status !==
                "Closed"
        ).length;
}
