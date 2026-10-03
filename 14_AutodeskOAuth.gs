// Per-Google-user Autodesk Platform Services 3-legged OAuth.
// Store APS_CLIENT_ID and APS_CLIENT_SECRET in Script Properties. Register
// https://script.google.com/macros/d/<script-id>/usercallback in the APS app.

const APS_USER_OAUTH_AUTHORIZE_ = 'https://developer.api.autodesk.com/authentication/v2/authorize';
const APS_USER_OAUTH_TOKEN_ = 'https://developer.api.autodesk.com/authentication/v2/token';
const APS_USER_OAUTH_SCOPE_ = 'data:read account:read viewables:read user-profile:read';
const APS_USER_TOKEN_PROP_ = 'APS_USER_TOKEN_V1';
const APS_USER_PENDING_PROP_ = 'APS_USER_OAUTH_PENDING_V1';
const APS_USER_EPOCH_PROP_ = 'APS_CONNECTION_EPOCH';
const APS_USER_ACCESS_CACHE_KEY_ = 'APS_USER_ACCESS_V1';
const APS_USER_INFO_URL_ = 'https://api.userprofile.autodesk.com/userinfo';

function apsUserIdentity_() {
  const context = getAuthenticatedUserContext_();
  if (!context || !context.success || !context.email) throw new Error('המשתמש אינו מורשה או שלא ניתן לזהות את חשבון Google.');
  const activeEmail = String(context.email).toLowerCase().trim();
  let effectiveEmail = '';
  try { effectiveEmail = String(Session.getEffectiveUser().getEmail() || '').toLowerCase().trim(); } catch (e) {}
  if (!effectiveEmail || effectiveEmail !== activeEmail) throw new Error('החיבור זמין רק כאשר אפליקציית Google פועלת בשם המשתמש המחובר.');
  return activeEmail;
}

function apsUserOAuthCredentials_() {
  const props = PropertiesService.getScriptProperties();
  const clientId = String(props.getProperty('APS_CLIENT_ID') || '').trim();
  const clientSecret = String(props.getProperty('APS_CLIENT_SECRET') || '').trim();
  return clientId && clientSecret ? { clientId: clientId, clientSecret: clientSecret } : null;
}

function apsUserOAuthCallbackUrl_() {
  return 'https://script.google.com/macros/d/' + ScriptApp.getScriptId() + '/usercallback';
}

function getAutodeskConnectionStatus() {
  try {
    const email = apsUserIdentity_();
    const props = PropertiesService.getUserProperties();
    const token = apsUserParseStoredToken_(props.getProperty(APS_USER_TOKEN_PROP_));
    const configured = !!apsUserOAuthCredentials_();
    if (!configured) return { success: true, configured: false, connected: false, email: email, error: 'חסרה הגדרת APS_CLIENT_ID או APS_CLIENT_SECRET ב-Script Properties.' };
    if (!token || !token.refresh_token || token.email !== email) return { success: true, configured: true, connected: false, email: email, epoch: apsUserCacheEpoch_() };
    const accessToken = apsUserToken_();
    const profile = apsUserProfile_(accessToken);
    return { success: true, configured: true, connected: true, email: email, autodeskName: String(profile.name || ''), autodeskUserId: String(profile.sub || profile.userId || ''), epoch: apsUserCacheEpoch_() };
  } catch (e) { return { success: false, error: String(e.message || 'לא ניתן לבדוק את מצב החיבור ל-Autodesk.') }; }
}

function getAutodeskAuthorizationUrl() {
  const email = apsUserIdentity_();
  const credentials = apsUserOAuthCredentials_();
  if (!credentials) throw new Error('החיבור ל-Autodesk טרם הוגדר על ידי מנהל המערכת.');
  const nonce = Utilities.getUuid();
  const state = ScriptApp.newStateToken()
    .withMethod('apsOAuthCallback')
    .withArgument('apsUserEmail', email)
    .withArgument('apsNonce', nonce)
    .withTimeout(600)
    .createToken();
  PropertiesService.getUserProperties().setProperty(APS_USER_PENDING_PROP_, JSON.stringify({ email: email, nonce: nonce, createdAt: Date.now() }));
  const params = {
    response_type: 'code', client_id: credentials.clientId,
    redirect_uri: apsUserOAuthCallbackUrl_(), scope: APS_USER_OAUTH_SCOPE_, state: state
  };
  return APS_USER_OAUTH_AUTHORIZE_ + '?' + Object.keys(params).map(function (key) {
    return encodeURIComponent(key) + '=' + encodeURIComponent(params[key]);
  }).join('&');
}

// Invoked by Google's /usercallback endpoint after it validates the state token.
function apsOAuthCallback(request) {
  const lock = LockService.getUserLock();
  let locked = false;
  try {
    const email = apsUserIdentity_();
    lock.waitLock(30000);
    locked = true;
    const params = request && request.parameter || {};
    const props = PropertiesService.getUserProperties();
    const pending = apsUserParseStoredToken_(props.getProperty(APS_USER_PENDING_PROP_));
    const pendingAge = pending ? Date.now() - Number(pending.createdAt) : NaN;
    if (!pending || !params.apsUserEmail || !params.apsNonce || !Number.isFinite(pendingAge) || pendingAge < 0 || pendingAge > 600000 ||
        String(params.apsUserEmail).toLowerCase() !== email || pending.email !== email ||
        params.apsNonce !== pending.nonce) {
      throw new Error('הבקשה אינה תקפה או שתוקפה פג.');
    }
    // Consume before exchanging so a callback cannot be replayed, including after failure.
    props.deleteProperty(APS_USER_PENDING_PROP_);
    if (params.error) throw new Error('Autodesk לא אישר את החיבור.');
    if (!params.code) throw new Error('Autodesk לא החזיר קוד הרשאה.');
    const token = apsUserTokenRequest_({
      grant_type: 'authorization_code', code: params.code,
      redirect_uri: apsUserOAuthCallbackUrl_()
    });
    if (!token.access_token || !token.refresh_token) throw new Error('Autodesk לא החזיר אסימוני חיבור תקינים.');
    const storedToken = apsUserNormalizeToken_(token);
    storedToken.email = email;
    props.setProperty(APS_USER_TOKEN_PROP_, JSON.stringify(storedToken));
    CacheService.getUserCache().remove(APS_USER_ACCESS_CACHE_KEY_);
    apsUserBumpCacheEpoch_();
    return HtmlService.createHtmlOutput('<p>החיבור ל-Autodesk הושלם. אפשר לסגור חלון זה.</p>');
  } catch (e) {
    return HtmlService.createHtmlOutput('<p>החיבור ל-Autodesk לא הושלם. חזור ל-BIMBOT ונסה שוב.</p>');
  } finally { if (locked) lock.releaseLock(); }
}

function disconnectAutodesk() {
  const lock = LockService.getUserLock();
  let locked = false;
  try {
    apsUserIdentity_();
    lock.waitLock(30000);
    locked = true;
    const props = PropertiesService.getUserProperties();
    props.deleteProperty(APS_USER_TOKEN_PROP_);
    props.deleteProperty(APS_USER_PENDING_PROP_);
    CacheService.getUserCache().remove(APS_USER_ACCESS_CACHE_KEY_);
    apsUserBumpCacheEpoch_();
    return { success: true, connected: false };
  } catch (e) { return { success: false, error: 'לא ניתן לנתק את Autodesk עבור המשתמש הנוכחי.' }; }
  finally { if (locked) lock.releaseLock(); }
}

function apsUserCacheEpoch_() {
  apsUserIdentity_();
  const props = PropertiesService.getUserProperties();
  return String(props.getProperty(APS_USER_EPOCH_PROP_) || '0');
}

function apsUserBumpCacheEpoch_() {
  const props = PropertiesService.getUserProperties();
  const current = Number(props.getProperty(APS_USER_EPOCH_PROP_) || 0);
  props.setProperty(APS_USER_EPOCH_PROP_, String((Number.isFinite(current) ? current : 0) + 1));
}

// Server-side only: returns an APS access token bound to the current Google user.
function apsUserToken_() {
  const email = apsUserIdentity_();
  const props = PropertiesService.getUserProperties();
  const cache = CacheService.getUserCache();
  const cached = cache.get(APS_USER_ACCESS_CACHE_KEY_);
  if (cached) return cached;
  const lock = LockService.getUserLock();
  lock.waitLock(30000);
  try {
    const afterLock = cache.get(APS_USER_ACCESS_CACHE_KEY_);
    if (afterLock) return afterLock;
    let token = apsUserParseStoredToken_(props.getProperty(APS_USER_TOKEN_PROP_));
    if (!token || !token.refresh_token || token.email !== email) throw new Error('יש לחבר תחילה את חשבון Autodesk האישי.');
    if (token.access_token && token.expires_at > Date.now() + 60000) {
      cache.put(APS_USER_ACCESS_CACHE_KEY_, token.access_token, Math.max(1, Math.min(21600, Math.floor((token.expires_at - Date.now() - 30000) / 1000))));
      return token.access_token;
    }
    let refreshed;
    try { refreshed = apsUserTokenRequest_({ grant_type: 'refresh_token', refresh_token: token.refresh_token, scope: APS_USER_OAUTH_SCOPE_ }); }
    catch (e) {
      if (e.httpCode === 400 || e.httpCode === 401) {
        props.deleteProperty(APS_USER_TOKEN_PROP_);
        cache.remove(APS_USER_ACCESS_CACHE_KEY_);
        apsUserBumpCacheEpoch_();
        throw new Error('פג תוקף חיבור Autodesk. יש להתחבר מחדש.');
      }
      throw e;
    }
    if (!refreshed.access_token) throw new Error('Autodesk לא הצליח לרענן את החיבור. יש להתחבר מחדש.');
    if (!refreshed.refresh_token) refreshed.refresh_token = token.refresh_token;
    token = apsUserNormalizeToken_(refreshed);
    token.email = email;
    props.setProperty(APS_USER_TOKEN_PROP_, JSON.stringify(token));
    cache.put(APS_USER_ACCESS_CACHE_KEY_, token.access_token, Math.max(1, Math.min(21600, token.expires_in - 120)));
    return token.access_token;
  } finally { lock.releaseLock(); }
}

function apsUserClearTokenCache_() {
  CacheService.getUserCache().remove(APS_USER_ACCESS_CACHE_KEY_);
}

function apsUserTokenRequest_(payload) {
  const credentials = apsUserOAuthCredentials_();
  if (!credentials) throw new Error('החיבור ל-Autodesk טרם הוגדר.');
  const body = Object.assign({}, payload);
  const response = UrlFetchApp.fetch(APS_USER_OAUTH_TOKEN_, {
    method: 'post', contentType: 'application/x-www-form-urlencoded', payload: body,
    headers: { Authorization: 'Basic ' + Utilities.base64Encode(credentials.clientId + ':' + credentials.clientSecret), Accept: 'application/json' },
    muteHttpExceptions: true, followRedirects: false
  });
  if (response.getResponseCode() < 200 || response.getResponseCode() >= 300) {
    const error = new Error('Autodesk דחה את בקשת החיבור (' + response.getResponseCode() + ').');
    error.httpCode = response.getResponseCode();
    throw error;
  }
  try { return JSON.parse(response.getContentText()); }
  catch (e) { throw new Error('Autodesk החזיר תשובת חיבור לא תקינה.'); }
}

function apsUserProfile_(accessToken) {
  const response = UrlFetchApp.fetch(APS_USER_INFO_URL_, {
    method: 'get', headers: { Authorization: 'Bearer ' + accessToken, Accept: 'application/json' },
    muteHttpExceptions: true, followRedirects: false
  });
  const code = response.getResponseCode();
  if (code < 200 || code >= 300) {
    if (code === 401) apsUserExpireToken_(accessToken);
    const error = new Error(code === 401 ? 'פג תוקף חיבור Autodesk. יש להתחבר מחדש.' : 'לא ניתן לקרוא את פרטי חשבון Autodesk (' + code + ').');
    error.httpCode = code;
    throw error;
  }
  try { return JSON.parse(response.getContentText()); }
  catch (e) { throw new Error('Autodesk החזיר פרטי משתמש לא תקינים.'); }
}

function apsUserExpireToken_(accessToken) {
  const lock = LockService.getUserLock();
  lock.waitLock(30000);
  try {
    const props = PropertiesService.getUserProperties();
    const current = apsUserParseStoredToken_(props.getProperty(APS_USER_TOKEN_PROP_));
    if (current && current.access_token === accessToken) {
      props.deleteProperty(APS_USER_TOKEN_PROP_);
      CacheService.getUserCache().remove(APS_USER_ACCESS_CACHE_KEY_);
      apsUserBumpCacheEpoch_();
    }
  } finally { lock.releaseLock(); }
}

function apsUserNormalizeToken_(token) {
  const normalized = Object.assign({}, token);
  const expiresIn = Number(token.expires_in);
  normalized.expires_in = Number.isFinite(expiresIn) && expiresIn > 0 ? expiresIn : 3600;
  normalized.expires_at = Date.now() + normalized.expires_in * 1000;
  return normalized;
}

function apsUserParseStoredToken_(raw) {
  if (!raw) return null;
  try { return JSON.parse(raw); } catch (e) { return null; }
}
