const msalConfig = {
    auth: {
        clientId: CONFIG.clientId,
        authority: 
            `https://login.microsoftonline.com/${CONFIG.tenantId}`,
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

      activeAccount = loginResponse.account;

document.getElementById("userInfo").innerText =
    activeAccount.username;

document.getElementById("loginButton").style.display =
    "none";

document.getElementById("loginButton").style.display =
    "none";

    } catch (error) {

        console.error("MSAL ERROR:", error);

        alert(
            "Login failed: " +
            error.message
        );
    }
}
