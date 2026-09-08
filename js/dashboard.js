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

                await loadProfilePhoto();
            }

            await loadDashboard();

        } catch (error) {

            console.error(error);
        }
    }
);

async function loadDashboard() {

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

    const incidents =
        data.value || [];

    buildKPIs(incidents);

    buildRiskRadar(incidents);

    buildSeverityChart(incidents);

    buildTopSites(incidents);
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

function buildRiskRadar(
    incidents
){

    const categories = {

        Injury:0,
        Environmental:0,
        Vehicle:0,
        Property:0,
        Unsafe:0

    };

    incidents.forEach(i => {

        const type =
            i.fields.IncidentType || "";

        if(type.includes("Injury"))
            categories.Injury++;

        if(type.includes("Environmental"))
            categories.Environmental++;

        if(type.includes("Vehicle"))
            categories.Vehicle++;

        if(type.includes("Property"))
            categories.Property++;

        if(type.includes("Unsafe"))
            categories.Unsafe++;
    });

    new Chart(

        document.getElementById(
            "riskRadar"
        ),

        {

            type:"radar",

            data:{

                labels:
                    Object.keys(
                        categories
                    ),

                datasets:[{

                    data:
                        Object.values(
                            categories
                        ),

                    borderColor:"#60a5fa",

                    backgroundColor:
                        "rgba(96,165,250,.20)",

                    pointBackgroundColor:
                        "#60a5fa"

                }]
            },

            options:{

                plugins:{
                    legend:{
                        display:false
                    }
                },

                scales:{
                    r:{

                        grid:{
                            color:
                                "rgba(255,255,255,.08)"
                        },

                        pointLabels:{
                            color:"#cbd5e1"
                        },

                        ticks:{
                            display:false
                        }
                    }
                }
            }
        }
    );
}

function buildSeverityChart(
    incidents
){

    const counts = {

        Low:0,
        Medium:0,
        High:0,
        Critical:0

    };

    incidents.forEach(i => {

        const sev =
            i.fields.Severity;

        if(counts[sev] !== undefined){

            counts[sev]++;
        }
    });

    new Chart(

        document.getElementById(
            "severityChart"
        ),

        {

            type:"doughnut",

            data:{

                labels:
                    Object.keys(counts),

                datasets:[{

                    data:
                        Object.values(counts),

                    backgroundColor:[

                        "#22c55e",
                        "#eab308",
                        "#f97316",
                        "#ef4444"

                    ],

                    borderWidth:0
                }]
            },

            options:{

                plugins:{

                    legend:{

                        labels:{
                            color:"#cbd5e1"
                        }
                    }
                }
            }
        }
    );
}

function buildTopSites(
    incidents
){

    const sites = {};

    incidents.forEach(i => {

        const site =
            i.fields.Site || "Unknown";

        sites[site] =
            (sites[site] || 0) + 1;
    });

    const topSites =
        Object.entries(sites)
            .sort(
                (a,b)=> b[1]-a[1]
            )
            .slice(0,5);

    const container =
        document.getElementById(
            "siteSummary"
        );

    container.innerHTML = "";

    topSites.forEach(
        ([site,count]) => {

            container.innerHTML += `

                <div class="chart-header">

                    <span class="chart-label">
                        ${site}
                    </span>

                    <span class="chart-value">
                        ${count}
                    </span>

                </div>

            `;
        }
    );
}
