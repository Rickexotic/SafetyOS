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

        sessionStorage.setItem(
    "SafetyOSUser",
    activeAccount.username
);

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

async function loadProfilePhoto() {

    try {

        const token =
            await getAccessToken();

        const response =
            await fetch(
                "https://graph.microsoft.com/v1.0/me/photo/$value",
                {
                    headers: {
                        Authorization:
                            `Bearer ${token}`
                    }
                }
            );

        if (!response.ok) {
            return;
        }

        const blob =
            await response.blob();

        const imageUrl =
            URL.createObjectURL(blob);

        const profilePhoto =
            document.getElementById(
                "profilePhoto"
            );

        const profileInitials =
            document.getElementById(
                "profileInitials"
            );

        if (profilePhoto) {

            profilePhoto.src =
                imageUrl;

            profilePhoto.classList.remove(
                "d-none"
            );
        }

        if (profileInitials) {

            profileInitials.style.display =
                "none";
        }

    } catch (error) {

        console.log(
            "No profile photo found:",
            error
        );
    }
}

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

                const loginStatus =
                    document.getElementById(
                        "loginStatus"
                    );

                if (loginStatus) {

                    loginStatus.classList.remove(
                        "d-none"
                    );
                }
            }

        } catch (error) {

            console.error(
                "Session restore error:",
                error
            );
        }
    }
);
