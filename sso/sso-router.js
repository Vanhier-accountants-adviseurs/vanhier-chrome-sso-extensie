const TENANT_ID = "790e3646-e472-40af-b3ee-1ce89d1472c3";

const SITE_BASE = "/vanhier-chrome-sso-extensie";

const MAPPING_URL =
    `${SITE_BASE}/mapping.json?v=${Date.now()}`;

const CALLBACK_URL =
    `${window.location.origin}${SITE_BASE}/sso/`;

const PENDING_KEY = "vanhier_sso_pending";

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

    if (path.startsWith("http")) {
        path = new URL(path).pathname;
    }

    return path.replace(/\/+$/, "") || "/";

}


function getRoutePath() {

    let path = normalizePath(
        window.location.pathname
    );

    console.log(
        "[SSO] Browser path:",
        path
    );

    /*
        Verwijder GitHub Pages repository-pad.
    */

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

    const routePath =
        normalizePath(path);

    console.log(
        "[SSO] Zoek route:",
        routePath
    );

    const config =
        mapping.find(item => {

            const mappingPath =
                normalizePath(item.path);

            console.log(
                "[SSO] Vergelijk:",
                mappingPath,
                "==",
                routePath
            );

            return (
                mappingPath === routePath ||
                mappingPath ===
                    `${SITE_BASE}${routePath}`
            );

        });

    console.log(
        "[SSO] Resultaat:",
        config
    );

    return config;

}


function getPending() {

    try {

        return JSON.parse(
            sessionStorage.getItem(
                PENDING_KEY
            ) || "null"
        );

    }
    catch {

        sessionStorage.removeItem(
            PENDING_KEY
        );

        return null;

    }

}


function setPending(config) {

    sessionStorage.setItem(
        PENDING_KEY,
        JSON.stringify({
            path: config.path,
            applicationId:
                config.applicationId
        })
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
                "sessionStorage"

        }

    });

}


async function authenticate(config) {

    console.log(
        "[SSO] Authenticatie voor:",
        config
    );

    setPending(config);

    const msalInstance =
        buildMsal(
            config.applicationId
        );

    await msalInstance.initialize();

    await msalInstance.loginRedirect({

        scopes: [
            "openid",
            "profile",
            "email"
        ],

        state:
            btoa(
                JSON.stringify({
                    path: config.path
                })
            )

    });

}


async function handleCallback(mapping) {

    console.log(
        "[SSO] Entra callback"
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

    const response =
        await msalInstance.handleRedirectPromise();

    console.log(
        "[SSO] Entra response:",
        response
    );

    if (!response) {

        throw new Error(
            "Geen Entra callback ontvangen"
        );

    }

    if (!response.account) {

        throw new Error(
            "Geen gebruikersaccount ontvangen van Entra"
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

            break;


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

    await authenticate(
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
            Centrale callback:
            /sso/
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
