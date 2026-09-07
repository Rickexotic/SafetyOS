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

        const f =
            item.fields;

        const photoLink =
            f.PhotoLink || "";

       const linkIcon = photoLink
    ? `
        <a href="${photoLink}"
           class="incident-link"
           target="_blank"
           rel="noopener noreferrer"
           aria-label="Open incident photo">
            <svg xmlns="http://www.w3.org/2000/svg"
                 width="14"
                 height="14"
                 fill="currentColor"
                 viewBox="0 0 16 16">
                <path d="M6.354 5.5H4a3 3 0 0 0 0 6h3a3 3 0 0 0 2.83-4H9q-.13 0-.25.031A2 2 0 0 1 7 10.5H4a2 2 0 1 1 0-4h1.535c.218-.376.495-.714.82-1z"/>
                <path d="M9 5.5a3 3 0 0 0-2.83 4h1.098A2 2 0 0 1 9 6.5h3a2 2 0 1 1 0 4h-1.535a4 4 0 0 1-.82 1H12a3 3 0 1 0 0-6z"/>
            </svg>
        </a>
      `
    : "";

        container.innerHTML += `

            <div class="incident-row">

                <div class="incident-top">

                    <span class="incident-status">
                        ${f.Status || ""}
                    </span>

                    <span class="incident-id">
                        ${f.IncidentID || ""}
                    </span>
                   
                   ${linkIcon}
                   
                </div>

                <div class="incident-bottom">

                     <span class="incident-severity">
                        ${f.Severity || ""}
                    </span>
                    
                    <span class="incident-type">
                        ${f.IncidentType || ""}
                    </span>

                    <span class="incident-site">
                        ${f.Site || ""}
                    </span>
               
                </div>

            </div>

        `;
    });
}

/* SEARCH */

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

/* SORT */

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

/* ROW SELECTION */

document.addEventListener(
    "click",
    function (e) {

        const row =
            e.target.closest(
                ".incident-row"
            );

        if (!row) {
            return;
        }

        document
            .querySelectorAll(
                ".incident-row"
            )
            .forEach(r =>
                r.classList.remove(
                    "active"
                )
            );

        row.classList.add(
            "active"
        );
    }
);
