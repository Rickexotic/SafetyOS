async function getAccessToken(){

    const tokenResponse =
        await msalInstance.acquireTokenSilent({

            account:
                activeAccount,

            scopes:[
                "Sites.ReadWrite.All"
            ]
        });

    return tokenResponse.accessToken;
}
