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

async function uploadPhoto(file) {

    if (!file) {
        return "";
    }

    const token =
        await getAccessToken();

    const fileName =
        `${Date.now()}_${file.name}`;

    const driveId =
        "b!cy-mWJulekmNfv_GhI-cxBK1NLW_-7NKiS9coSLvnBf9iMkSyWHMTbiRI93oerkZ";

    const uploadUrl =
        `https://graph.microsoft.com/v1.0/drives/${driveId}/root:/${fileName}:/content`;

    console.log("UPLOAD URL:", uploadUrl);

    const response =
        await fetch(
            uploadUrl,
            {
                method: "PUT",

                headers: {
                    Authorization: `Bearer ${token}`,
                    "Content-Type": file.type
                },

                body: file
            }
        );

    const result =
        await response.json();

    console.log("UPLOAD RESPONSE:", result);

    if (!response.ok) {

        throw new Error(
            JSON.stringify(result)
        );
    }

    return result.webUrl;
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
            document.getElementById(
                "incidentPhoto"
            ).files[0];
        
console.log("PHOTO FILE:", photoFile);
        
        const photoUrl =
            await uploadPhoto(photoFile);

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

                PhotoUrl:
                    photoUrl,

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

        console.log(
            "POST URL:",
            url
        );

        console.log("UPLOAD URL:", uploadUrl);

        const response =
            await fetch(
                url,
                {
                    method: "POST",

                    headers: {
                        Authorization:
                            `Bearer ${token}`,
                        "Content-Type":
                            "application/json"
                    },

                    body:
                        JSON.stringify(body)
                }
            );

        if (response.ok) {

            document
                .getElementById("incidentForm")
                .reset();

            const preview =
                document.getElementById(
                    "previewImage"
                );

            if (preview) {
                preview.classList.add(
                    "d-none"
                );
            }

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

async function getDrives() {

    const token =
        await getAccessToken();

    const response =
        await fetch(
            "https://graph.microsoft.com/v1.0/sites/46y2.sharepoint.com:/sites/SaaS_OHS:/drives",
            {
                headers: {
                    Authorization: `Bearer ${token}`
                }
            }
        );

    const data =
        await response.json();

    console.log(data);
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

document
    .getElementById("incidentPhoto")
    ?.addEventListener(
        "change",
        function () {

            const file =
                this.files[0];

            if (!file) {
                return;
            }

            const reader =
                new FileReader();

            reader.onload =
                function (e) {

                    const img =
                        document.getElementById(
                            "previewImage"
                        );

                    if (!img) {
                        return;
                    }

                    img.src =
                        e.target.result;

                    img.classList.remove(
                        "d-none"
                    );
                };

            reader.readAsDataURL(file);
        }
    );
