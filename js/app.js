function showSuccess(message) {

    const success =
        document.getElementById("successMessage");

    success.innerText = message;

    success.classList.remove("d-none");

    document
        .getElementById("errorMessage")
        .classList.add("d-none");
}

function showError(message) {

    const error =
        document.getElementById("errorMessage");

    error.innerText = message;

    error.classList.remove("d-none");

    document
        .getElementById("successMessage")
        .classList.add("d-none");
}

function validateForm() {

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

    if (
        !site ||
        !type ||
        !severity ||
        !description
    ) {

        showError(
            "Please complete all required fields."
        );

        return false;
    }

    return true;
}

async function submitIncident() {

    try {

        if (!activeAccount) {

            showError(
                "Please sign in first."
            );

            return;
        }

        if (!validateForm()) {
            return;
        }

        const token =
            await getAccessToken();

        const photoFile =
    document.getElementById("incidentPhoto")
        .files[0];

        const body = {
            fields: {

                Title: "Incident",

                IncidentID:
                    "INC-" +
                    new Date().getFullYear() +
                    "-" +
                    Math.floor(
                        Math.random() * 100000
                    ),

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

        console.log(
            "Payload:",
            JSON.stringify(body, null, 2)
        );

       const url =
    `https://graph.microsoft.com/v1.0/sites/46y2.sharepoint.com:/sites/SaaS_OHS:/lists/${CONFIG.incidentsListId}/items`;

console.log("POST URL:", url);

const response =
    await fetch(
        url,
        {
            method: "POST",

            headers: {
                Authorization: `Bearer ${token}`,
                "Content-Type": "application/json"
            },

            body: JSON.stringify(body)
        }
    );


        if (response.ok) {

            document
                .getElementById("incidentForm")
                .reset();

            showSuccess(
                "Incident submitted successfully."
            );

        } else {

            const error =
                await response.json();

            console.error(
                "Graph Error:",
                error
            );

            showError(
                JSON.stringify(error)
            );
        }

    } catch (ex) {

        console.error(
            "Unexpected Error:",
            ex
        );

        showError(
            ex.message
        );
    }
}

async function uploadPhoto(file) {

    if (!file) {
        return "";
    }

    const token =
        await getAccessToken();

    const fileName =
        `${Date.now()}_${file.name}`;

    const response =
        await fetch(
            `https://graph.microsoft.com/v1.0/sites/46y2.sharepoint.com:/sites/SaaS_OHS:/drive/root:/IncidentPhotos/${fileName}:/content`,
            {
                method: "PUT",

                headers: {
                    Authorization: `Bearer ${token}`,
                    "Content-Type": file.type
                },

                body: file
            }
        );

    const data =
        await response.json();

    return data.webUrl;
}

async function testSiteInfo() {

    try {

        const token =
            await getAccessToken();

        const response =
            await fetch(
                "https://graph.microsoft.com/v1.0/sites/root",
                {
                    headers: {
                        Authorization:
                            `Bearer ${token}`
                    }
                }
            );

        const data =
            await response.json();

        console.log(
            "SITE INFO:",
            data
        );

    } catch (error) {

        console.error(error);
    }
}
