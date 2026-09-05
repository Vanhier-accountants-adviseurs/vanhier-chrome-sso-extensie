// ================================
// Existing Vanhier SSO
// ================================

// Visma Connect - Nmbrs Accounting
if (location.hostname === "connect.visma.com") {
    if (location.href.includes("accounting.nmbrs.nl")) {
        console.log("Visma: Nmbrs-login gedetecteerd");

        function startMicrosoftLogin() {
            const form = document.getElementById("form-provider-microsoft");

            if (form) {
                console.log("Visma: Microsoft-formulier gevonden");
                form.requestSubmit();
                return true;
            }

            return false;
        }

        if (!startMicrosoftLogin()) {
            console.log("Visma: Microsoft-formulier nog niet gevonden");

            const observer = new MutationObserver(() => {
                if (startMicrosoftLogin()) {
                    observer.disconnect();
                }
            });

            observer.observe(document.documentElement, {
                childList: true,
                subtree: true
            });
        }
    }
}

// Visma Connect - Nmbrs Payroll
if (location.hostname === "connect.visma.com") {
    if (location.href.includes("payroll.nmbrs.com")) {
        console.log("Visma: Nmbrs-login gedetecteerd");

        function startMicrosoftLogin() {
            const form = document.getElementById("form-provider-microsoft");

            if (form) {
                console.log("Visma: Microsoft-formulier gevonden");
                form.requestSubmit();
                return true;
            }

            return false;
        }

        if (!startMicrosoftLogin()) {
            console.log("Visma: Microsoft-formulier nog niet gevonden");

            const observer = new MutationObserver(() => {
                if (startMicrosoftLogin()) {
                    observer.disconnect();
                }
            });

            observer.observe(document.documentElement, {
                childList: true,
                subtree: true
            });
        }
    }
}

// Payt
if (location.hostname === "app.paytsoftware.com") {
    function klikMicrosoftPayt() {
        const button = [...document.querySelectorAll("button")]
            .find(b => b.textContent.trim() === "Microsoft");

        if (button) {
            console.log("Payt: Microsoft-knop gevonden");
            button.click();
            return true;
        }

        return false;
    }

    if (!klikMicrosoftPayt()) {
        const observer = new MutationObserver(() => {
            if (klikMicrosoftPayt()) {
                observer.disconnect();
            }
        });

        observer.observe(document.documentElement, {
            childList: true,
            subtree: true
        });
    }
}

// ================================
// ChatGPT Auto Login trigger page
// ================================
if (
    location.hostname === "vanhier-accountants-adviseurs.github.io" &&
    location.pathname === "/vanhier-chrome-sso-extensie/start-chatgpt/"
) {
    console.log("ChatGPT SSO: trigger endpoint geopend");
    chrome.runtime.sendMessage({ type: "start-chatgpt-login" });
}

// ================================
// ChatGPT login flow
// ================================

let handledRoute = "";

const sleep = (ms) =>
    new Promise(resolve => setTimeout(resolve, ms));

function setInputValue(input, value) {
    input.focus();

    const setter = Object.getOwnPropertyDescriptor(
        HTMLInputElement.prototype,
        "value"
    ).set;

    setter.call(input, value);

    input.dispatchEvent(
        new InputEvent("input", {
            bubbles: true,
            inputType: "insertText",
            data: value
        })
    );

    input.dispatchEvent(
        new Event("change", { bubbles: true })
    );
}

async function waitFor(selector, timeout = 10000) {
    const start = Date.now();

    while (Date.now() - start < timeout) {
        const element = document.querySelector(selector);

        if (element) {
            return element;
        }

        await sleep(50);
    }

    return null;
}

async function getLoginData() {
    const result = await chrome.storage.session.get("chatgptLoginData");
    return result.chatgptLoginData || null;
}

async function handleRoute() {
    const path = location.pathname;

    if (
        location.hostname === "chatgpt.com" &&
        !path.startsWith("/auth/")
    ) {
        if (handledRoute !== "logged-in") {
            handledRoute = "logged-in";
            await focusChatGPT();
            await chrome.storage.session.remove("chatgptLoginData");
        }
        return;
    }

    if (path === "/auth/login") {
        if (handledRoute === path) return;
        handledRoute = path;

        const data = await getLoginData();
        if (!data) return;

        const emailInput = await waitFor(
            'input[type="email"], input[name="email"]'
        );

        if (!emailInput) {
            handledRoute = "";
            return;
        }

        setInputValue(emailInput, data.email);

        const button = await waitFor('button[type="submit"]');

        if (button) {
            await sleep(150);
            button.click();
        }

        return;
    }

    if (path === "/log-in/password") {
        if (handledRoute === path) return;
        handledRoute = path;

        const data = await getLoginData();
        if (!data) return;

        const passwordInput = await waitFor(
            'input[type="password"], input[name="password"]'
        );

        if (!passwordInput) {
            handledRoute = "";
            return;
        }

        setInputValue(passwordInput, data.password);

        const button = await waitFor(
            'button[name="intent"][value="validate"]'
        );

        if (button) {
            await sleep(150);
            button.click();
        }

        return;
    }

    if (path.startsWith("/mfa-challenge/")) {
        if (handledRoute === path) return;
        handledRoute = path;

        const data = await getLoginData();
        if (!data) return;

        const codeInput = await waitFor('input[name="code"]');

        if (!codeInput) {
            handledRoute = "";
            return;
        }

        setInputValue(codeInput, data.totpCode);

        const verifyButton = await waitFor(
            'button[name="intent"][value="verify"]'
        );

        if (verifyButton) {
            await sleep(200);
            verifyButton.click();
        }

        return;
    }
}

async function focusChatGPT() {
    const editor = await waitFor(
        '#prompt-textarea[contenteditable="true"]',
        15000
    );

    if (!editor) return;

    editor.focus();

    const selection = window.getSelection();
    const range = document.createRange();

    range.selectNodeContents(editor);
    range.collapse(false);

    selection.removeAllRanges();
    selection.addRange(range);
}

if (
    location.hostname === "chatgpt.com" ||
    location.hostname === "auth.openai.com"
) {
    setInterval(handleRoute, 100);
    handleRoute();
}
