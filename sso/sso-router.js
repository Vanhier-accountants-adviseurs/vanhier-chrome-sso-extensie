const TENANT_ID =
    "790e3646-e472-40af-b3ee-1ce89d1472c3";

const SITE_BASE =
    "/vanhier-chrome-sso-extensie";

const MAPPING_URL =
    `${SITE_BASE}/mapping.json?v=${Date.now()}`;

const statusElement =
    document.getElementById("status");


function setStatus(text) {

    console.log("[SSO]", text);

    if (statusElement) {
        statusElement.textContent = text;
    }

}


function normalizePath(path) {

    if (!path) {
        return "/";
    }

    path =
        decodeURIComponent(path);

    return (
        path.replace(/\/+$/, "") || "/"
    );

}


function getRoutePath() {

    let path =
        normalizePath(
            window.location.pathname
        );

    console.log(
        "[SSO] Volledige browser path:",
        path
    );

    if (path.startsWith(SITE_BASE)) {

        path =
            path.substring(
                SITE_BASE.length
            );

    }

    path =
        normalizePath(path);

    console.log(
        "[SSO] Route path:",
        path
    );

    return path;

}


async function loadMapping() {

    console.log(
        "[SSO] Mapping laden:",
        MAPPING_URL
    );

    const response =
        await fetch(
            MAPPING_URL,
            {
                cache: "no-store"
            }
        );

    if (!response.ok) {

        throw new Error(
            `mapping.json kon niet worden geladen (${response.status})`
        );

    }

    const mapping =
        await response.json();

    console.log(
        "[SSO] Mapping:",
        mapping
    );

    if (!Array.isArray(mapping)) {

        throw new Error(
            "mapping.json moet een array zijn"
        );

    }

    return mapping;

}


function findConfig(mapping, path) {

    const wanted =
        normalizePath(path);

    console.log(
        "[SSO] Zoek configuratie voor:",
        wanted
    );

    const config =
        mapping.find(item => {

            const itemPath =
                normalizePath(
                    item.path
                );

            console.log(
                "[SSO] Vergelijk:",
                itemPath,
                "met",
                wanted
            );

            return itemPath === wanted;

        });

    console.log(
        "[SSO] Configuratie:",
        config
    );

    return config;

}


function buildMsal(config) {

    if (
        !config.applicationId ||
        config.applicationId.startsWith("HIER-DE-")
    ) {

        throw new Error(
            `Geen geldige applicationId voor ${config.path}`
        );

    }

    /*
        BELANGRIJK:
        iedere toepassing gebruikt zichzelf
        als redirect URI.
    */

    const redirectUri =
        `${window.location.origin}${SITE_BASE}${config.path}`;

    console.log(
        "[SSO] Redirect URI:",
        redirectUri
    );

    return new msal.PublicClientApplication({

        auth: {

            clientId:
                config.applicationId,

            authority:
                `https://login.microsoftonline.com/${TENANT_ID}`,

            redirectUri:
                redirectUri

        },

        cache: {

            cacheLocation:
                "sessionStorage"

        }

    });

}


async function authenticate(config) {

    const msalInstance =
        buildMsal(config);

    await msalInstance.initialize();

    /*
        HEEL BELANGRIJK:

        Ook op de eerste pagina-load roepen we
        handleRedirectPromise() aan.

        Als er nog een lopende redirect-interactie
        is, wordt die hier afgehandeld voordat
        loginRedirect() opnieuw wordt aangeroepen.
    */

    console.log(
        "[SSO] Controleren op bestaande Entra callback..."
    );

    const response =
        await msalInstance.handleRedirectPromise();

    if (response) {

        console.log(
            "[SSO] Entra login succesvol:",
            response.account
        );

        setStatus(
            "Toegang gecontroleerd. Doorsturen..."
        );

        execute(config);

        return;

    }

    /*
        Bestaande login?
    */

    const accounts =
        msalInstance.getAllAccounts();

    console.log(
        "[SSO] Bestaande accounts:",
        accounts
    );

    if (accounts.length > 0) {

        setStatus(
            "Toegang gecontroleerd. Doorsturen..."
        );

        execute(config);

        return;

    }

    /*
        Nog niet ingelogd.
    */

    setStatus(
        "Doorsturen naar Microsoft..."
    );

    console.log(
        "[SSO] Start loginRedirect()"
    );

    await msalInstance.loginRedirect({

        scopes: [
            "openid",
            "profile",
            "email"
        ]

    });

}


function execute(config) {

    console.log(
        "[SSO] Actie uitvoeren:",
        config
    );

    switch (config.type) {

        case "redirect":

        case "form":

            window.location.replace(
                config.outputUrl
            );

            break;

        default:

            throw new Error(
                `Onbekend type: ${config.type}`
            );

    }

}


async function main() {

    try {

        const mapping =
            await loadMapping();

        const path =
            getRoutePath();

        const config =
            findConfig(
                mapping,
                path
            );

        if (!config) {

            setStatus(
                `Deze SSO-route bestaat niet: ${path}`
            );

            return;

        }

        setStatus(
            "Aanmelden..."
        );

        await authenticate(
            config
        );

    }
    catch (error) {

        console.error(
            "[SSO] FOUT:",
            error
        );

        setStatus(
            `Aanmelden mislukt: ${error.message}`
        );

    }

}


main();
