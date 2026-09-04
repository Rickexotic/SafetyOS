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

    incidents.forEach(item => {

        const f = item.fields;

        container.innerHTML += `

            <div class="incident-row">

                <div class="row-id">
                    ${f.IncidentID || ""}
                </div>

                <div class="row-type">
                    ${f.IncidentType || ""}
                </div>

                <div class="row-site">
                    ${f.Site || ""}
                </div>

                <div class="row-severity severity-${(f.Severity || "").toLowerCase()}">
                    ${f.Severity || ""}
                </div>

                <div class="row-status">
                    ${f.Status || ""}
                </div>

            </div>

        `;
    });
}


document
    .getElementById("sortBy")
    ?.addEventListener(
        "change",
        function(){

            const value =
                this.value;

            const sorted =
                [...allIncidents];

            if(value === "site"){

                sorted.sort(
                    (a,b)=>
                        (a.fields.Site || "")
                        .localeCompare(
                            b.fields.Site || ""
                        )
                );
            }

            if(value === "status"){

                sorted.sort(
                    (a,b)=>
                        (a.fields.Status || "")
                        .localeCompare(
                            b.fields.Status || ""
                        )
                );
            }

            if(value === "severity"){

                const rank = {
                    Low:1,
                    Medium:2,
                    High:3,
                    Critical:4
                };

                sorted.sort(
                    (a,b)=>
                        (rank[b.fields.Severity] || 0)
                        -
                        (rank[a.fields.Severity] || 0)
                );
            }

            renderIncidents(
                sorted
            );
        }
    );

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
