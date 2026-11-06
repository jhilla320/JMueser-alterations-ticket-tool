import { GOOGLE_CLIENT_ID, DRIVE_SCOPE, ALLOWED_EMAIL_DOMAIN, ALLOWED_EMAILS } from "./config.js";

const TOKEN_KEY = "driveAccessToken";
const TOKEN_EXP_KEY = "driveAccessTokenExp";
const SCOPE_KEY = "driveAccessTokenScope";
const EMAIL_KEY = "driveAccountEmail";

let tokenClient = null;

function ensureClient() {
  if (tokenClient) return tokenClient;
  if (!window.google?.accounts?.oauth2) return null;
  tokenClient = window.google.accounts.oauth2.initTokenClient({
    client_id: GOOGLE_CLIENT_ID,
    scope: DRIVE_SCOPE,
    callback: () => {},
  });
  return tokenClient;
}

// One token request at a time: if a background refresh is already open and
// the person also clicks Sign in, both share the same request instead of
// the second one orphaning the first.
let pendingRequest = null;

function requestToken(prompt) {
  if (pendingRequest) return pendingRequest;
  pendingRequest = requestTokenOnce(prompt).finally(() => {
    pendingRequest = null;
  });
  return pendingRequest;
}

function requestTokenOnce(prompt) {
  return new Promise((resolve, reject) => {
    const client = ensureClient();
    if (!client) {
      reject(new Error("Google Identity Services not loaded"));
      return;
    }
    client.callback = (response) => {
      if (response?.access_token) {
        resolve(response);
      } else {
        reject(new Error(response?.error_description || "No access token returned"));
      }
    };
    client.error_callback = (err) => {
      reject(new Error(err?.message || "Google sign-in was cancelled"));
    };
    // Passing the remembered email as a hint lets Google skip the account
    // chooser, so re-signing in is a single quick step.
    const hint = getSessionEmail();
    client.requestAccessToken(hint ? { prompt, hint } : { prompt });
  });
}

async function fetchAccountEmail(accessToken) {
  const response = await fetch("https://www.googleapis.com/oauth2/v3/userinfo", {
    headers: { Authorization: `Bearer ${accessToken}` },
  });
  if (!response.ok) throw new Error("Could not verify Google account");
  const data = await response.json();
  return (data.email || "").toLowerCase();
}

function isEmailAllowed(email) {
  if (!ALLOWED_EMAIL_DOMAIN && ALLOWED_EMAILS.length === 0) return true;
  if (!email) return false;
  const normalizedAllowList = ALLOWED_EMAILS.map((item) => item.toLowerCase());
  if (normalizedAllowList.includes(email)) return true;
  if (ALLOWED_EMAIL_DOMAIN && email.endsWith(`@${ALLOWED_EMAIL_DOMAIN.toLowerCase()}`)) return true;
  return false;
}

function storeSession(response, email) {
  const expiry = Date.now() + Number(response.expires_in || 0) * 1000;
  localStorage.setItem(TOKEN_KEY, response.access_token);
  localStorage.setItem(TOKEN_EXP_KEY, String(expiry));
  localStorage.setItem(SCOPE_KEY, DRIVE_SCOPE);
  localStorage.setItem(EMAIL_KEY, email);
}

export function clearSession() {
  localStorage.removeItem(TOKEN_KEY);
  localStorage.removeItem(TOKEN_EXP_KEY);
  localStorage.removeItem(SCOPE_KEY);
  localStorage.removeItem(EMAIL_KEY);
}

export function getSessionEmail() {
  return localStorage.getItem(EMAIL_KEY) || "";
}

export function hasValidSession() {
  const token = localStorage.getItem(TOKEN_KEY);
  const exp = Number(localStorage.getItem(TOKEN_EXP_KEY) || 0);
  const scope = localStorage.getItem(SCOPE_KEY) || "";
  const email = getSessionEmail();
  return Boolean(token && exp && scope === DRIVE_SCOPE && Date.now() < exp - 30_000 && isEmailAllowed(email));
}

export function getToken() {
  return hasValidSession() ? localStorage.getItem(TOKEN_KEY) : "";
}

// Runs the full sign-in flow: request a token, verify the account is
// allowed, and store the session. Throws a friendly error (and clears any
// partial session) if the account is not on the allow list.
export async function signIn(prompt = "") {
  const response = await requestToken(prompt);
  const email = await fetchAccountEmail(response.access_token);
  if (!isEmailAllowed(email)) {
    clearSession();
    const scope = ALLOWED_EMAIL_DOMAIN ? `an @${ALLOWED_EMAIL_DOMAIN} account` : "an approved account";
    throw new Error(`${email} isn't authorized for this tool. Sign in with ${scope}.`);
  }
  storeSession(response, email);
  return email;
}

export async function getValidToken() {
  if (hasValidSession()) return getToken();
  try {
    await signIn(""); // silent refresh — no visible prompt if still signed into Google and already granted access
  } catch {
    await signIn("consent"); // fall back to the full sign-in screen if silent refresh isn't possible
  }
  return getToken();
}

// Browsers only let a page open Google's small sign-in popup in response to
// a click or keypress, so a refresh at a random moment gets blocked. Calling
// this from normal user input means a token that's about to expire (or just
// did) is renewed during ordinary use, instead of surprising someone later
// with a sign-in screen. Resolves true if a refresh succeeded.
let backgroundRefresh = null;
let lastRefreshFailure = 0;

export function refreshIfExpiringSoon(windowMs = 10 * 60 * 1000) {
  if (backgroundRefresh) return backgroundRefresh;
  const exp = Number(localStorage.getItem(TOKEN_EXP_KEY) || 0);
  if (!exp || !getSessionEmail()) return null; // never signed in on this browser
  if (Date.now() < exp - windowMs) return null; // plenty of time left
  if (Date.now() - lastRefreshFailure < 60_000) return null; // don't keep retrying right after a failure

  backgroundRefresh = signIn("")
    .then(() => true)
    .catch(() => {
      lastRefreshFailure = Date.now();
      return false;
    })
    .finally(() => {
      backgroundRefresh = null;
    });
  return backgroundRefresh;
}

export function signOut() {
  clearSession();
}
