importScripts("config.js");

const sleep = (ms) => new Promise(resolve => setTimeout(resolve, ms));

chrome.storage.session.setAccessLevel({
  accessLevel: "TRUSTED_AND_UNTRUSTED_CONTEXTS"
}).catch(error => {
  console.error("ChatGPT Auto Login: storage access level fout:", error);
});

chrome.action.onClicked.addListener(async (tab) => {
  try {
    await startChatGPTLogin(tab.id);
  } catch (error) {
    console.error("ChatGPT Auto Login fout:", error);
  }
});

chrome.runtime.onMessage.addListener((message, sender) => {
  if (message?.type !== "start-chatgpt-login") return;

  const tabId = sender.tab?.id;
  if (!tabId) return;

  startChatGPTLogin(tabId).catch(error => {
    console.error("ChatGPT Auto Login fout:", error);
  });
});

async function startChatGPTLogin(tabId) {
  const loginData = await getLoginData();

  await chrome.storage.session.set({
    chatgptLoginData: loginData
  });

  await chrome.tabs.update(tabId, {
    url: "https://chatgpt.com/auth/logout"
  });

  await sleep(2000);

  await chrome.tabs.update(tabId, {
    url: "https://chatgpt.com/auth/login"
  });
}

async function getLoginData() {
  const token = await getAccessToken();

  const response = await fetch(CONFIG.functionUrl, {
    method: "GET",
    headers: {
      "Authorization": `Bearer ${token}`,
      "Accept": "application/json"
    }
  });

  const text = await response.text();

  if (!response.ok) {
    throw new Error(`Function HTTP ${response.status}: ${text}`);
  }

  const data = JSON.parse(text);

  if (!data.email || !data.password || !data.totpCode) {
    throw new Error("Function response mist email/password/totpCode");
  }

  return data;
}

async function getAccessToken() {
  const redirectUri = chrome.identity.getRedirectURL("entra");
  const state = crypto.randomUUID();
  const verifier = randomString(64);
  const challenge = await pkceChallenge(verifier);

  const authorizeUrl = new URL(
    `https://login.microsoftonline.com/${CONFIG.tenantId}/oauth2/v2.0/authorize`
  );

  authorizeUrl.searchParams.set("client_id", CONFIG.clientId);
  authorizeUrl.searchParams.set("response_type", "code");
  authorizeUrl.searchParams.set("redirect_uri", redirectUri);
  authorizeUrl.searchParams.set("response_mode", "query");
  authorizeUrl.searchParams.set("scope", `openid profile ${CONFIG.apiScope}`);
  authorizeUrl.searchParams.set("state", state);
  authorizeUrl.searchParams.set("code_challenge", challenge);
  authorizeUrl.searchParams.set("code_challenge_method", "S256");

  const resultUrl = await chrome.identity.launchWebAuthFlow({
    url: authorizeUrl.toString(),
    interactive: true
  });

  const result = new URL(resultUrl);

  if (result.searchParams.get("state") !== state) {
    throw new Error("OAuth state mismatch");
  }

  const error = result.searchParams.get("error");
  if (error) {
    throw new Error(result.searchParams.get("error_description") || error);
  }

  const code = result.searchParams.get("code");
  if (!code) {
    throw new Error("Geen authorization code ontvangen");
  }

  const tokenResponse = await fetch(
    `https://login.microsoftonline.com/${CONFIG.tenantId}/oauth2/v2.0/token`,
    {
      method: "POST",
      headers: {
        "Content-Type": "application/x-www-form-urlencoded"
      },
      body: new URLSearchParams({
        client_id: CONFIG.clientId,
        grant_type: "authorization_code",
        code,
        redirect_uri: redirectUri,
        code_verifier: verifier,
        scope: CONFIG.apiScope
      })
    }
  );

  const tokenText = await tokenResponse.text();

  if (!tokenResponse.ok) {
    throw new Error(`Token HTTP ${tokenResponse.status}: ${tokenText}`);
  }

  const tokenData = JSON.parse(tokenText);

  if (!tokenData.access_token) {
    throw new Error("Geen access token ontvangen");
  }

  return tokenData.access_token;
}

function randomString(length) {
  const chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-._~";
  const bytes = new Uint8Array(length);
  crypto.getRandomValues(bytes);

  return [...bytes]
    .map(b => chars[b % chars.length])
    .join("");
}

async function pkceChallenge(verifier) {
  const bytes = new TextEncoder().encode(verifier);
  const digest = await crypto.subtle.digest("SHA-256", bytes);

  let binary = "";
  for (const byte of new Uint8Array(digest)) {
    binary += String.fromCharCode(byte);
  }

  return btoa(binary)
    .replace(/\+/g, "-")
    .replace(/\//g, "_")
    .replace(/=+$/, "");
}
