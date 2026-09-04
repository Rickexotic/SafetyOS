const msalConfig = {
    auth: {
        clientId: CONFIG.clientId,
        authority: `https://login.microsoftonline.com/${CONFIG.tenantId}`,
        redirectUri: CONFIG.redirectUri
    }
};

const msalInstance =
    new msal.PublicClientApplication(msalConfig);

let activeAccount = null;

async function signIn() {

    try {

        await msalInstance.initialize();

        const loginResponse =
            await msalInstance.loginPopup({
                scopes: [
                    "User.Read",
                    "Sites.ReadWrite.All"
                ]
            });

        activeAccount =
            loginResponse.account;

        document.getElementById("userInfo").innerText =
            activeAccount.username;

        document.getElementById("loginButton").style.display =
            "none";

        console.log(
            "Logged in:",
            activeAccount
        );

    }
    catch(error) {

        console.error(
            "MSAL ERROR:",
            error
        );

        alert(
            "Login failed: " +
            error.message
        );
    }
}

async function getAccessToken() {

    const tokenResponse =
        await msalInstance.acquireTokenSilent({

            account:
                activeAccount,

            scopes: [
                "Sites.ReadWrite.All"
            ]
        });

    return tokenResponse.accessToken;
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
