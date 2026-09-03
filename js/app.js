function showSuccess(message){

    const success =
        document.getElementById("successMessage");

    success.innerText =
        message;

    success.classList.remove("d-none");

    document
        .getElementById("errorMessage")
        .classList.add("d-none");
}

function showError(message){

    const error =
        document.getElementById("errorMessage");

    error.innerText =
        message;

    error.classList.remove("d-none");

    document
        .getElementById("successMessage")
        .classList.add("d-none");
}

function validateForm(){

    const site =
        document.getElementById("site").value;

    const type =
        document.getElementById("incidentType").value;

    const severity =
        document.getElementById("severity").value;

    const description =
        document.getElementById("description").value;

    console.log("Site:", site);
    console.log("Type:", type);
    console.log("Severity:", severity);
    console.log("Description:", description);

    if(
        !site ||
        !type ||
        !severity ||
        !description
    ){
        showError(
            "Please complete all required fields."
        );

        return false;
    }

    return true;
}
async function submitIncident(){

    try{

        if(!activeAccount){

            showError(
                "Please sign in first."
            );

            return;
        }

        if(!validateForm()){
            return;
        }

        const token =
            await getAccessToken();

        const body = {

            fields:{

                Title: "Incident",

                IncidentID:
                    "INC-" +
                    new Date().getFullYear() +
                    "-" +
                    Math.floor(Math.random() * 100000),

                DateReported:
                    new Date().toISOString(),

                Site:
                    document.getElementById("site").value,

                IncidentType:
                    document.getElementById("incidentType").value,

                Severity:
                    document.getElementById("severity").value,

                Description:
                    document.getElementById("description").value,

                Status:
                    "Reported"
            }
        };

        const response =
            await fetch(
                `https://graph.microsoft.com/v1.0/sites/${CONFIG.siteId}/lists/${CONFIG.incidentsListId}/items`,
                {
                    method: "POST",

                    headers: {
                        Authorization: `Bearer ${token}`,
                        "Content-Type": "application/json"
                    },

                    body: JSON.stringify(body)
                }
            );

        if(response.ok){

            document
                .getElementById("incidentForm")
                .reset();

            showSuccess(
                "Incident submitted successfully."
            );
        }
        else{

            const error =
                await response.text();

            console.error(error);

            showError(
                "Unable to create incident."
            );
        }

    }
    catch(ex){

        console.error(ex);

        showError(
            "Unexpected error occurred." );
    }
}
        
async function getSiteInfo() {

    const token = await getAccessToken();

    const response = await fetch(
        "https://graph.microsoft.com/v1.0/sites/46y2.sharepoint.com:/sites/SaaS_OHS",
        {
            headers: {
                Authorization: `Bearer ${token}`
            }
        }
    );

    const data = await response.json();

    console.log(data);

        
    }
