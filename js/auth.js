const msalConfig = {
    auth: {
        clientId: CONFIG.clientId,
        authority: `https://login.microsoftonline.com/${CONFIG.tenantId}`,
        redirectUri: CONFIG.redirectUri
    }
};

const msalInstance = new msal.PublicClientApplication(msalConfig);

let activeAccount = null;

async function signIn() {
    try {

        const loginResponse =
            await msalInstance.loginPopup({
                scopes: [
                    "User.Read",
                    "Sites.ReadWrite.All"
                ]
            });

        activeAccount = loginResponse.account;

        document.getElementById("userInfo").innerText =
            activeAccount.username;

    } catch (error) {

        console.error(error);

        alert("Login failed");
    }
}
