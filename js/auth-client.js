/* =====================================================================
   Google Sign-In (Google Identity Services)
   ---------------------------------------------------------------------
   Gates the whole app behind a sign-in step. The ID token this produces
   is attached as an Authorization: Bearer header on every request to
   the CORS proxy — the proxy itself verifies it and checks it against
   the Sheet's "Users" tab (see worker.js's authorize()), so this is a
   real access boundary, not just a UI gate.

   The token is held only in memory (a module-level variable) — never
   written to localStorage — and expires after roughly an hour, at which
   point the next request fails with a clear "please sign in again"
   rather than silently doing nothing.
   ===================================================================== */

const AuthClient = (() => {
  let idToken = null;
  let user = null; // { email, name, picture }
  let initialized = false;

  function loadClientId() { return localStorage.getItem("slc.signInClientId") || ""; }
  function saveClientId(id) { localStorage.setItem("slc.signInClientId", id || ""); }

  function init(onSignedIn) {
    const clientId = loadClientId();
    if (!clientId || typeof google === "undefined") return false;
    google.accounts.id.initialize({
      client_id: clientId,
      callback: (resp) => {
        idToken = resp.credential;
        user = decodeJwtPayload(resp.credential);
        onSignedIn();
      },
      auto_select: true,
    });
    initialized = true;
    return true;
  }

  function renderButton(el) {
    if (!initialized) return;
    google.accounts.id.renderButton(el, { theme: "outline", size: "large", text: "signin_with" });
    google.accounts.id.prompt(); // offers One Tap / auto-select if a session is already remembered
  }

  function decodeJwtPayload(jwt) {
    try {
      const payload = jwt.split(".")[1];
      const json = atob(payload.replace(/-/g, "+").replace(/_/g, "/"));
      return JSON.parse(decodeURIComponent(escape(json)));
    } catch { return null; }
  }

  function getToken() { return idToken; }
  function currentUser() { return user; }
  function isSignedIn() { return !!idToken; }

  function signOut() {
    idToken = null;
    user = null;
    if (typeof google !== "undefined") google.accounts.id.disableAutoSelect();
  }

  return { loadClientId, saveClientId, init, renderButton, getToken, currentUser, isSignedIn, signOut };
})();
