const msalConfig = {
    auth: {
        clientId: CONFIG.clientId,

        authority:
            `https://login.microsoftonline.com/${CONFIG.tenantId}`,

        redirectUri:
            CONFIG.redirectUri
    }
};

const msalInstance =
    new msal.PublicClientApplication(
        msalConfig
    );

let activeAccount = null;

/* -------------------------
   Restore Account
-------------------------- */

window.addEventListener(
     => {

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

            const loginButton =
                document.getElementById(
                    "loginButton"
                );

            if (loginButton) {

                loginButton.style.display =
                    "none";
            }
        }
    }
);

/* -------------------------
   Login
-------------------------- */

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

        document
            .getElementById("userInfo")
            .innerText =
            activeAccount.username;

        const loginButton =
            document.getElementById(
                "loginButton"
            );

        if (loginButton) {

            loginButton.style.display =
                "none";
        }

    } catch (error) {

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

/* -------------------------
   Access Token
-------------------------- */

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
