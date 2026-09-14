const TENANT_ID = "790e3646-e472-40af-b3ee-1ce89d1472c3";
const SITE_BASE = "/vanhier-chrome-sso-extensie";
const MAPPING_URL = `${SITE_BASE}/mapping.json`;
const CALLBACK_URL = `${window.location.origin}${SITE_BASE}/sso/`;

const statusElement = document.getElementById("status");

function setStatus(text) {
    if (statusElement) {
        statusElement.textContent = text;
    }
}

function normalizePath(path) {
    if (!path) return "/";
    const withoutTrailingSlash = path.replace(/\/+$/, "");
    return withoutTrailingSlash || "/";
}

function currentPath() {
    return normalizePath(window.location.pathname);
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
    const normalized = normalizePath(path);

    return mapping.find(item => {
        return normalizePath(item.path) === normalized;
    });
}

function decodeState(state) {
    if (!state) return null;

    try {
        return JSON.parse(atob(state));
    } catch {
        return null;
    }
}

function buildMsal(config) {
    if (!config.applicationId || config.applicationId.startsWith("HIER-DE-")) {
        throw new Error(`Geen geldige applicationId ingesteld voor ${config.path}`);
    }

    return new msal.PublicClientApplication({
        auth: {
            clientId: config.applicationId,
            authority: `https://login.microsoftonline.com/${TENANT_ID}`,
            redirectUri: CALLBACK_URL
        },
        cache: {
            cacheLocation: "sessionStorage"
        }
    });
}

async function authenticate(config) {
    const msalInstance = buildMsal(config);

    await msalInstance.initialize();

    const response = await msalInstance.handleRedirectPromise();

    if (response) {
        return {
            account: response.account,
            state: decodeState(response.state)
        };
    }

    const accounts = msalInstance.getAllAccounts();

    if (accounts.length > 0) {
        return {
            account: accounts[0],
            state: null
        };
    }

    await msalInstance.loginRedirect({
        scopes: ["openid", "profile", "email"],
        state: btoa(JSON.stringify({
            path: config.path
        }))
    });

    return null;
}

function execute(config) {
    if (config.type === "redirect") {
        window.location.replace(config.outputUrl);
        return;
    }

    if (config.type === "form") {
        // De formulierhandeling gebeurt later via de Chrome-extensie.
        window.location.replace(config.outputUrl);
        return;
    }

    throw new Error(`Onbekend type: ${config.type}`);
}

async function main() {
    try {
        const mapping = await loadMapping();

        const initialPath = currentPath();

        // Een Entra callback komt terug op /sso/.
        // Het oorspronkelijke pad staat in state.
        let path = initialPath;

        setStatus("Aanmelden...");

        // Gebruik eerst een tijdelijke MSAL-instantie om de redirect te verwerken.
        // De clientId staat in de mapping van het oorspronkelijke pad.
        if (initialPath === normalizePath(`${SITE_BASE}/sso`)) {
            // Op de callbackpagina kunnen we nog niet weten welke app gebruikt werd.
            // Probeer het oorspronkelijke pad uit de URL/state niet te raden;
            // daarom gebruiken we de pending route uit sessionStorage.
            const pendingPath = sessionStorage.getItem("vanhier_sso_pending_path");

            if (pendingPath) {
                path = normalizePath(pendingPath);
            }
        }

        let config = findConfig(mapping, path);

        if (!config) {
            setStatus("Deze SSO-route bestaat niet.");
            return;
        }

        // Onthoud de route voordat we naar Entra gaan.
        sessionStorage.setItem("vanhier_sso_pending_path", config.path);

        const authResult = await authenticate(config);

        if (!authResult) {
            return;
        }

        // Bij een callback mag state de route definitief bepalen.
        const statePath = authResult.state?.path;
        if (statePath) {
            config = findConfig(mapping, normalizePath(statePath));

            if (!config) {
                throw new Error("De SSO-route uit de Entra state bestaat niet meer.");
            }
        }

        sessionStorage.removeItem("vanhier_sso_pending_path");

        setStatus("Toegang gecontroleerd. Doorsturen...");
        execute(config);

    } catch (error) {
        console.error("SSO router fout:", error);
        setStatus(`Aanmelden mislukt: ${error.message}`);
    }
}

main();
