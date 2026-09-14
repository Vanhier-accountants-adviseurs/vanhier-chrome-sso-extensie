const TENANT_ID = "790e3646-e472-40af-b3ee-1ce89d1472c3";
const MAPPING_URL = "../mapping.json";
const ROUTER_URL = "../sso/";

const statusElement = document.getElementById("status");

function setStatus(text) {
    if (statusElement) {
        statusElement.textContent = text;
    }
}

function currentPath() {
    const path = window.location.pathname.replace(/\/+$/, "");
    return path || "/";
}

async function loadMapping() {
    const response = await fetch(MAPPING_URL, { cache: "no-store" });
    if (!response.ok) {
        throw new Error(`mapping.json kon niet worden geladen (${response.status})`);
    }

    const mapping = await response.json();

    if (!Array.isArray(mapping)) {
        throw new Error("mapping.json moet een array bevatten");
    }

    return mapping;
}

function findConfig(mapping, path) {
    return mapping.find(item => item.path.replace(/\/+$/, "") === path);
}

function buildMsal(config) {
    if (!config.applicationId || config.applicationId.startsWith("HIER-DE-")) {
        throw new Error(`Geen geldige applicationId ingesteld voor ${config.path}`);
    }

    return new msal.PublicClientApplication({
        auth: {
            clientId: config.applicationId,
            authority: `https://login.microsoftonline.com/${TENANT_ID}`,
            redirectUri: `${window.location.origin}${ROUTER_URL}`
        }
    });
}

async function authenticate(config) {
    const msalInstance = buildMsal(config);

    await msalInstance.initialize();

    const response = await msalInstance.handleRedirectPromise();

    if (response) {
        return response.account;
    }

    const accounts = msalInstance.getAllAccounts();

    if (accounts.length > 0) {
        return accounts[0];
    }

    await msalInstance.loginRedirect({
        scopes: ["openid", "profile", "email"],
        state: btoa(JSON.stringify({ path: config.path }))
    });

    return null;
}

function execute(config) {
    if (config.type === "redirect") {
        window.location.href = config.outputUrl;
        return;
    }

    if (config.type === "form") {
        // De daadwerkelijke formulier-invulling gebeurt door de Chrome-extensie.
        // We sturen alleen naar de doelapplicatie. De extensie kan op basis van
        // hostname/configuratie de velden invullen.
        window.location.href = config.outputUrl;
        return;
    }

    throw new Error(`Onbekend type: ${config.type}`);
}

async function main() {
    try {
        const mapping = await loadMapping();
        const path = currentPath();
        const config = findConfig(mapping, path);

        if (!config) {
            setStatus("Deze SSO-route bestaat niet.");
            return;
        }

        setStatus(`Aanmelden bij ${config.path}...`);
        const account = await authenticate(config);

        if (!account) {
            return;
        }

        setStatus("Toegang gecontroleerd. Doorsturen...");
        execute(config);
    } catch (error) {
        console.error("SSO router fout:", error);
        setStatus(`Aanmelden mislukt: ${error.message}`);
    }
}

main();
