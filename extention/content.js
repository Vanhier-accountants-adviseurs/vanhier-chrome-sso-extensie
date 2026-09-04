// ================================
// Visma Connect
// ================================
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

// ================================
// Payt
// ================================
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

    // Direct proberen
    if (!klikMicrosoftPayt()) {

        // Wachten als Payt de knop later inlaadt
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
