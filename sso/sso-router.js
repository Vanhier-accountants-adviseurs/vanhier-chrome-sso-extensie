const TENANT_ID =
    "790e3646-e472-40af-b3ee-1ce89d1472c3";

const SITE_BASE =
    "/vanhier-chrome-sso-extensie";

const MAPPING_URL =
    `${SITE_BASE}/mapping.json?v=${Date.now()}`;

const CALLBACK_URL =
    `${window.location.origin}${SITE_BASE}/sso/`;

const PENDING_KEY =
    "vanhier_sso_pending";

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

    path = decodeURIComponent(path);

    return path.replace(/\/+$/, "") || "/";

}


function getRoutePath() {

    let path =
        normalizePath(
            window.location.pathname
        );

    console.log(
        "[SSO] Browser path:",
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
        "[SSO] Zoek route:",
        wanted
    );

    const config =
        mapping.find(item => {

            const itemPath =
                normalizePath(item.path);

            console.log(
                "[SSO] Vergelijk:",
                itemPath,
                "==",
                wanted
            );

            return itemPath === wanted;

        });

    console.log(
        "[SSO] Resultaat:",
        config
    );

    return config;

}


function getPending() {

    try {

        const value =
            sessionStorage.getItem(
                PENDING_KEY
            );

        if (!value) {
            return null;
        }

        return JSON.parse(value);

    }
    catch (error) {

        console.error(
            "[SSO] Ongeldige pending data:",
            error
        );

        sessionStorage.removeItem(
            PENDING_KEY
        );

        return null;

    }

}


function setPending(config) {

    const pending = {

        path:
            config.path,

        applicationId:
            config.applicationId

    };

    sessionStorage.setItem(
        PENDING_KEY,
        JSON.stringify(pending)
    );

    console.log(
        "[SSO] Pending opgeslagen:",
        pending
    );

}


function clearPending() {

    sessionStorage.removeItem(
        PENDING_KEY
    );

}


function buildMsal(applicationId) {

    if (
        !applicationId ||
        applicationId.startsWith("HIER-DE-")
    ) {

        throw new Error(
            "Geen geldige applicationId ingesteld"
        );

    }

    console.log(
        "[SSO] MSAL Application ID:",
        applicationId
    );

    console.log(
        "[SSO] Redirect URI:",
        CALLBACK_URL
    );

    return new msal.PublicClientApplication({

        auth: {

            clientId:
                applicationId,

            authority:
                `https://login.microsoftonline.com/${TENANT_ID}`,

            redirectUri:
                CALLBACK_URL

        },

        cache: {

            cacheLocation:
                "sessionStorage",

            storeAuthStateInCookie:
                false

        }

    });

}


async function startLogin(config) {

    setPending(config);

    const msalInstance =
        buildMsal(
            config.applicationId
        );

    await msalInstance.initialize();

    /*
        Eerst controleren of MSAL al een callback
        moet verwerken.
    */

    const response =
        await msalInstance.handleRedirectPromise();

    if (response) {

        console.log(
            "[SSO] Bestaande callback verwerkt:",
            response
        );

        execute(config);

        return;

    }

    /*
        Geen callback, dus normale login starten.
    */

    setStatus(
        "Doorsturen naar Microsoft..."
    );

    console.log(
        "[SSO] loginRedirect starten"
    );

    await msalInstance.loginRedirect({

        scopes: [
            "openid",
            "profile",
            "email"
        ]

    });

}


async function handleCallback(mapping) {

    console.log(
        "[SSO] Callbackpagina geopend"
    );

    const pending =
        getPending();

    console.log(
        "[SSO] Pending:",
        pending
    );

    if (!pending) {

        throw new Error(
            "Geen openstaande SSO-aanmelding gevonden"
        );

    }

    const config =
        findConfig(
            mapping,
            pending.path
        );

    if (!config) {

        throw new Error(
            `Geen configuratie gevonden voor ${pending.path}`
        );

    }

    const msalInstance =
        buildMsal(
            pending.applicationId
        );

    await msalInstance.initialize();

    console.log(
        "[SSO] handleRedirectPromise uitvoeren"
    );

    const response =
        await msalInstance.handleRedirectPromise();

    console.log(
        "[SSO] Entra response:",
        response
    );

    if (!response) {

        /*
            Dit kan voorkomen als de callbackpagina
            opnieuw geladen wordt nadat de callback
            al verwerkt is.

            Controleer daarom ook bestaande accounts.
        */

        const accounts =
            msalInstance.getAllAccounts();

        if (accounts.length === 0) {

            throw new Error(
                "Geen Entra callback en geen ingelogde gebruiker gevonden"
            );

        }

        console.log(
            "[SSO] Bestaand account gevonden:",
            accounts[0]
        );

    }

    clearPending();

    setStatus(
        "Toegang gecontroleerd. Doorsturen..."
    );

    execute(config);

}


function execute(config) {

    console.log(
        "[SSO] Uitvoeren:",
        config
    );

    switch (config.type) {

        case "redirect":

            window.location.replace(
                config.outputUrl
            );

            return;

        case "form":

            /*
                Voorlopig hetzelfde als redirect.
                Later kan hier de specifieke Twinfield-flow
                worden toegevoegd.
            */

            window.location.replace(
                config.outputUrl
            );

            return;

        default:

            throw new Error(
                `Onbekend type: ${config.type}`
            );

    }

}


async function startRoute(mapping) {

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

    await startLogin(
        config
    );

}


async function main() {

    try {

        const mapping =
            await loadMapping();

        const path =
            getRoutePath();

        /*
            /sso/ is de centrale Entra callback.
        */

        if (path === "/sso") {

            await handleCallback(
                mapping
            );

            return;

        }

        await startRoute(
            mapping
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
