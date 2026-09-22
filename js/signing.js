/* =====================================================================
   Signing helpers — requires CryptoJS (loaded via CDN in index.html)
   ===================================================================== */

// SolisCloud request signing: Authorization: API {KeyId}:{sign}
// sign = Base64(HMAC-SHA1(KeySecret, "POST\n{Content-MD5}\napplication/json\n{Date}\n{CanonicalizedResource}"))
// Verify field order/casing against your SolisCloud Open Platform API PDF —
// this mirrors the pattern SolisCloud publishes but has not been tested
// against a live account (see BRANDS.solis.docsNote in js/brands.js).
function solisSign(apiSecret, method, resourcePath, bodyString) {
  const md5 = CryptoJS.MD5(bodyString).toString(CryptoJS.enc.Base64);
  const date = new Date().toUTCString();
  const contentType = "application/json;charset=UTF-8";
  const stringToSign = `${method}\n${md5}\n${contentType}\n${date}\n${resourcePath}`;
  const sign = CryptoJS.HmacSHA1(stringToSign, apiSecret).toString(CryptoJS.enc.Base64);
  return { md5, date, sign };
}

// Growatt legacy portal login expects an MD5 hash of the plaintext password.
function md5Hex(plain) {
  return CryptoJS.MD5(plain).toString(CryptoJS.enc.Hex);
}

if (typeof module !== "undefined") module.exports = { solisSign, md5Hex };
