let allIncidents = [];

window.addEventListener(
    "load",
    async () => {

        await msalInstance.initialize();

        const accounts =
            msalInstance.getAllAccounts();

        if (accounts.length > 0) {

            activeAccount =
                accounts[0];

            const userInfo =
                document.getElementById(
                    "userInfo"
                );

            if (userInfo) {

                userInfo.innerText =
                    activeAccount.username;
            }
        }

        await loadIncidents();

    }
);

async function loadIncidents() {

    try {

        const token =
            await getAccessToken();

        const response =
            await fetch(
                `https://graph.microsoft.com/v1.0/sites/46y2.sharepoint.com:/sites/SaaS_OHS:/lists/${CONFIG.incidentsListId}/items?expand=fields`,
                {
                    headers: {
                        Authorization:
                            `Bearer ${token}`
                    }
                }
            );

        const data =
            await response.json();

        console.log(data);

        allIncidents =
            data.value || [];

        renderIncidents(
            allIncidents
        );

    } catch (ex) {

        console.error(ex);
    }
}

function renderIncidents(
    incidents
) {

    const container =
        document.getElementById(
            "incidentHistory"
        );

    container.innerHTML = "";

    incidents
        .slice()
        .reverse()
        .forEach(item => {

            const f =
                item.fields;

            container.innerHTML += `
                <div class="incident-card">

                    <div class="incident-id">
                        ${f.IncidentID || ""}
                    </div>

                    <div class="incident-type">
                        ${f.IncidentType || ""}
                    </div>

                    <div class="incident-grid">

                        <span>
                            📍 ${f.Site || ""}
                        </span>

                        <span>
                            ⚠ ${f.Severity || ""}
                        </span>

                    </div>

                    <div class="incident-status">
                        ${f.Status || ""}
                    </div>

                </div>
            `;
        });
}

document.addEventListener(
    "input",
    function (e) {

        if (
            e.target.id !==
            "searchBox"
        ) {
            return;
        }

        const text =
            e.target.value
                .toLowerCase();

        const filtered =
            allIncidents.filter(i => {

                const f =
                    i.fields;

                return JSON.stringify(f)
                    .toLowerCase()
                    .includes(text);
            });

        renderIncidents(
            filtered
        );
    }
);
