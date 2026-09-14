const TENANT_ID = "790e3646-e472-40af-b3ee-1ce89d1472c3";

const SITE_BASE = "/vanhier-chrome-sso-extensie";

const MAPPING_URL =
    `${SITE_BASE}/mapping.json`;

const CALLBACK_URL =
    `${window.location.origin}${SITE_BASE}/sso/`;

const PENDING_KEY = "vanhier_sso_pending";

const statusElement =
    document.getElementById("status");


function setStatus(text) {

    if (statusElement) {
        statusElement.textContent = text;
    }

}


function normalizePath(path) {

    if (!path) {
        return "/";
    }

    return path.replace(/\/+$/, "") || "/";

}


function currentPath() {

    /*
        Browser:
        /vanhier-chrome-sso-extensie/sso/start-auditcase

        Mapping:
        /sso/start-auditcase

        Daarom verwijderen we SITE_BASE.
    */

    let path = window.location.pathname;

    if (path.startsWith(SITE_BASE)) {
        path = path.substring(SITE_BASE.length);
    }

    return normalizePath(path);

}


async function loadMapping() {

    const response = await fetch(
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

    return mapping.find(item =>
        normalizePath(item.path) === wanted
    );

}


function getPending() {

    try {

        return JSON.parse(
            sessionStorage.getItem(PENDING_KEY) || "null"
        );

    }
    catch {

        sessionStorage.removeItem(PENDING_KEY);

        return null;

    }

}


function setPending(config) {

    sessionStorage.setItem(
        PENDING_KEY,
        JSON.stringify({
            path: config.path,
            applicationId: config.applicationId
        })
    );

}


function clearPending() {

    sessionStorage.removeItem(PENDING_KEY);

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

            clientId: applicationId,

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

    setPending(config);

    const msalInstance =
        buildMsal(config.applicationId);

    await msalInstance.initialize();

    await msalInstance.loginRedirect({

        scopes: [
            "openid",
            "profile",
            "email"
        ]

    });

}


async function handleCallback(mapping) {

    const pending =
        getPending();

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

    switch (config.type) {

        case "redirect":

            window.location.replace(
                config.outputUrl
            );

            break;


        case "form":

            /*
                Voor nu alleen redirect.

                Later kunnen we hier bijvoorbeeld
                Twinfield-specifieke logica
                aan koppelen.
            */

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
        currentPath();

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

    await authenticate(config);

}


async function main() {

    try {

        const mapping =
            await loadMapping();

        const path =
            currentPath();

        /*
            /sso is onze centrale
            Entra callback.
        */

        if (path === "/sso") {

            await handleCallback(mapping);

            return;

        }

        await startRoute(mapping);

    }
    catch (error) {

        console.error(
            "SSO router fout:",
            error
        );

        setStatus(
            `Aanmelden mislukt: ${error.message}`
        );

    }

}


main();
