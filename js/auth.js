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

window.addEventListener("load", async () => {

    await msalInstance.initialize();

    const accounts =
        msalInstance.getAllAccounts();

    if (accounts.length > 0) {

        activeAccount =
            accounts[0];

        const userInfo =
            document.getElementById("userInfo");

        if (userInfo) {

            userInfo.innerText =
                activeAccount.username;
        }

        const loginButton =
            document.getElementById("loginButton");

        if (loginButton) {

            loginButton.style.display =
                "none";
        }
    }
});

async function signIn() {

    try {

        const loginResponse =
            await msalInstance.loginPopup({
                scopes: [
                    "User.Read",
                    "Sites.ReadWrite.All"
                ]
            });

        activeAccount =
            loginResponse.account;

        const userInfo =
            document.getElementById("userInfo");

        if (userInfo) {

            userInfo.innerText =
                activeAccount.username;
        }

        const loginButton =
            document.getElementById("loginButton");

        if (loginButton) {

            loginButton.style.display =
                "none";
        }

    }
    catch (error) {

        console.error(error);

        alert(
            "Login failed: " +
            error.message
        );
    }
}

async function getAccessToken() {

    const response =
        await msalInstance.acquireTokenSilent({
            account: activeAccount,
            scopes: [
                "Sites.ReadWrite.All"
            ]
        });

    return response.accessToken;
}
