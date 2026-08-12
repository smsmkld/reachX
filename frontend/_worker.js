/**
 * Reach X — Cloudflare Pages Worker
 *
 * Sits between the browser and the Apps Script Master Controls.
 *
 *   Browser  ──/api/*──▶  this Worker  ──POST──▶  Master Control /exec
 *
 * Responsibilities:
 *   1. Serve the static frontend (index.html, app.html, assets/)
 *   2. Authenticate users against KV
 *   3. Map each user to their assigned Master Control
 *   4. Proxy allow-listed actions, so the browser never learns the /exec URL
 *
 * Required bindings (set in Cloudflare dashboard):
 *   KV namespace binding : REACHX_KV
 *   Secret               : SESSION_PEPPER      (any long random string)
 *   Secret               : MASTER_1_URL        (…/exec of master control 1)
 *   Secret               : MASTER_2_URL  … MASTER_5_URL   (as many as you run)
 */

const SESSION_TTL_SECONDS = 60 * 60 * 12;  // 12 hours
const COOKIE_NAME         = 'rx_session';
const PBKDF2_ITERATIONS   = 100000;

/* Admin is a separate credential with its own cookie and its own KV namespace
 * prefix. Sharing either with the user session would mean one stolen cookie
 * reaches every tenant, and it would make "am I admin?" a field on a record the
 * admin screen itself can edit. Shorter TTL because the blast radius is bigger. */
const ADMIN_COOKIE_NAME   = 'rx_admin';
const ADMIN_TTL_SECONDS   = 60 * 60 * 8;   // 8 hours

/* ── Actions the browser is allowed to invoke ───────────────────────────────
 * Anything not on this list is rejected by the Worker before it reaches Apps
 * Script. Sender-internal actions (writeColdInbox, updateColdSent, syncConfig,
 * getSenderStatus, …) are deliberately absent — they are part of the
 * sender→master callback protocol and must never be callable from a browser.
 */
const ALLOWED_ACTIONS = new Set([
  // reads
  'getCampaigns', 'getCampaignStats', 'getLeads', 'getSenders', 'getDailySummary',
  'getSystemStatus', 'getFullSystemReport', 'getSettings', 'getAllTags',
  'getSendersByTag', 'getBlocklist', 'getMasterTriggers', 'healthCheck',
  'validateSystem',
  // writes
  'createCampaign', 'updateCampaign', 'deleteCampaign', 'uploadLeads',
  'updateSender', 'updateSenderConfig', 'updateSendersByTag', 'assignTag',
  'removeTag', 'updateSettings', 'addToBlocklist', 'removeFromBlocklist',
  'sendTestEmail',
  // operations
  'distributeCold', 'distributeFollowups', 'kickoffAllSenders',
  'kickoffAllFollowups', 'emergencyStop', 'resumeAll', 'setupAllTriggers',
  'cleanupAllTriggers', 'purgeAllInboxes',
  'runWarmupIncrease', 'systemBoot', 'sendDailyDigest', 'resetOrphanedLeads',
  'archiveCampaignLeads', 'syncStats', 'checkSenderChains', 'restartSenderChain',
  // resetSentToday is deliberately absent: it lets senders exceed their daily
  // limit, so it belongs to the admin app rather than to every user.
]);

/* ── helpers ───────────────────────────────────────────────────────────── */

const json = (data, status = 200, headers = {}) =>
  new Response(JSON.stringify(data), {
    status,
    headers: { 'Content-Type': 'application/json', 'Cache-Control': 'no-store', ...headers },
  });

const bufToHex = (buf) =>
  [...new Uint8Array(buf)].map((b) => b.toString(16).padStart(2, '0')).join('');

function randomHex(bytes = 32) {
  return bufToHex(crypto.getRandomValues(new Uint8Array(bytes)));
}

/** Constant-time string compare — avoids leaking hash bytes via timing. */
function safeEqual(a, b) {
  if (typeof a !== 'string' || typeof b !== 'string' || a.length !== b.length) return false;
  let diff = 0;
  for (let i = 0; i < a.length; i++) diff |= a.charCodeAt(i) ^ b.charCodeAt(i);
  return diff === 0;
}

/**
 * PBKDF2-SHA256. `pepper` is a Worker secret, so a leaked KV dump alone is not
 * enough to mount an offline attack against the stored hashes.
 */
async function hashPassword(password, saltHex, pepper) {
  const key = await crypto.subtle.importKey(
    'raw', new TextEncoder().encode(password + (pepper || '')), 'PBKDF2', false, ['deriveBits'],
  );
  const bits = await crypto.subtle.deriveBits(
    {
      name: 'PBKDF2',
      salt: new Uint8Array(saltHex.match(/.{2}/g).map((h) => parseInt(h, 16))),
      iterations: PBKDF2_ITERATIONS,
      hash: 'SHA-256',
    },
    key, 256,
  );
  return bufToHex(bits);
}

function getCookie(request, name) {
  const header = request.headers.get('Cookie') || '';
  for (const part of header.split(';')) {
    const [k, ...v] = part.trim().split('=');
    if (k === name) return v.join('=');
  }
  return null;
}

function sessionCookie(token, maxAge) {
  return `${COOKIE_NAME}=${token}; HttpOnly; Secure; SameSite=Strict; Path=/; Max-Age=${maxAge}`;
}

/**
 * Resolves the /exec URL for a user's master control.
 *
 * Two sources, checked in order:
 *
 *   1. `masterUrl` on the user's own KV record — the URL travels with the user,
 *      so adding one is a single KV entry with no redeploy and no ceiling on
 *      how many Master Controls exist.
 *   2. `MASTER_<n>_URL` in the environment, keyed by `masterId` — the original
 *      scheme, kept so existing users keep working.
 *
 * Both are server-side only; neither ever reaches the browser.
 */
function masterUrlFor(env, user) {
  const own = user && typeof user.masterUrl === 'string' ? user.masterUrl.trim() : '';
  if (own) return own;

  const id = user && user.masterId;
  if (id === undefined || id === null || id === '') return null;

  const url = env[`MASTER_${id}_URL`];
  return typeof url === 'string' && url.trim() ? url.trim() : null;
}

async function readSession(request, env) {
  const token = getCookie(request, COOKIE_NAME);
  if (!token) return null;
  const raw = await env.REACHX_KV.get(`session:${token}`);
  if (!raw) return null;
  try {
    return { token, ...JSON.parse(raw) };
  } catch {
    return null;
  }
}

/**
 * Resolves the caller from either an API key or a browser session.
 *
 * A cookie is the wrong credential for a machine: it is HttpOnly, it expires in
 * 12 hours, and an automation would have to re-login and juggle Set-Cookie on a
 * schedule. A key in KV under `apikey:<key>` holds the same record a session
 * does, so downstream code cannot tell the two apart.
 *
 * The key IS the lookup, so there is no secret to compare and nothing to leak
 * through timing. It also never expires — revoking means deleting the KV entry.
 */
async function readAuth(request, env) {
  const key = request.headers.get('X-API-Key');
  if (key) {
    const raw = await env.REACHX_KV.get(`apikey:${key}`);
    if (!raw) return null;
    try {
      return JSON.parse(raw);
    } catch {
      return null;
    }
  }
  return readSession(request, env);
}

/* ── admin ─────────────────────────────────────────────────────────────── */

function adminCookie(token, maxAge) {
  return `${ADMIN_COOKIE_NAME}=${token}; HttpOnly; Secure; SameSite=Strict; Path=/; Max-Age=${maxAge}`;
}

/** Resolves the admin from its own cookie. Never falls back to a user session. */
async function readAdmin(request, env) {
  const token = getCookie(request, ADMIN_COOKIE_NAME);
  if (!token) return null;
  const raw = await env.REACHX_KV.get(`adminsession:${token}`);
  if (!raw) return null;
  try {
    return { token, ...JSON.parse(raw) };
  } catch {
    return null;
  }
}

/** Every /api/admin/* route goes through this. No route may skip it. */
const requireAdmin = (admin) =>
  admin ? null : json({ success: false, error: 'Admin sign-in required' }, 401);

async function listKeys(env, prefix) {
  const out = [];
  let cursor;
  do {
    const page = await env.REACHX_KV.list({ prefix, cursor });
    out.push(...page.keys.map((k) => k.name));
    cursor = page.list_complete ? null : page.cursor;
  } while (cursor);
  return out;
}

/** Strips the secrets before any user record is sent to a browser. */
const publicUser = (u) => ({
  email:      u.email,
  username:   u.username || u.name || u.email,
  masterUrl:  u.masterUrl || '',
  masterId:   u.masterId ?? '',
  role:       u.role || 'user',
  disabled:   !!u.disabled,
  createdAt:  u.createdAt || '',
  lastLogin:  u.lastLogin || '',
});

async function readUser(env, email) {
  const raw = await env.REACHX_KV.get(`user:${email}`);
  if (!raw) return null;
  try {
    return JSON.parse(raw);
  } catch {
    return null;
  }
}

/**
 * Drops every live session for one user.
 *
 * Disabling an account has to bite immediately. Without this, an already
 * signed-in user keeps working for the rest of their 12-hour session — the one
 * moment you most need the switch to be instant is the one it would miss.
 */
async function revokeUserSessions(env, email) {
  const names = await listKeys(env, 'session:');
  let dropped = 0;
  for (const name of names) {
    const raw = await env.REACHX_KV.get(name);
    if (!raw) continue;
    try {
      if (JSON.parse(raw).email === email) {
        await env.REACHX_KV.delete(name);
        dropped++;
      }
    } catch { /* a corrupt record is not a session worth keeping */ }
  }
  return dropped;
}

/**
 * Creates the first admin, once.
 *
 * Refuses when an admin already exists, and requires ADMIN_SETUP_KEY so the
 * window between deploying and claiming the account is not open to whoever
 * finds the URL first. Leave the secret unset and this route stays shut.
 */
async function handleAdminBootstrap(request, env) {
  const setupKey = env.ADMIN_SETUP_KEY;
  if (!setupKey) {
    return json({ success: false, error: 'ADMIN_SETUP_KEY is not set, so bootstrap is disabled.' }, 403);
  }
  if (!safeEqual(request.headers.get('X-Setup-Key') || '', setupKey)) {
    return json({ success: false, error: 'Bad setup key' }, 403);
  }

  const existing = await listKeys(env, 'admin:');
  if (existing.length > 0) {
    return json({ success: false, error: 'An admin already exists. Bootstrap is closed.' }, 403);
  }

  let body;
  try {
    body = await request.json();
  } catch {
    return json({ success: false, error: 'Invalid request' }, 400);
  }

  const email    = String(body.email || '').trim().toLowerCase();
  const password = String(body.password || '');
  if (!email || password.length < 12) {
    return json({ success: false, error: 'Email and a password of at least 12 characters are required' }, 400);
  }

  const salt = randomHex(16);
  await env.REACHX_KV.put(`admin:${email}`, JSON.stringify({
    email,
    salt,
    passwordHash: await hashPassword(password, salt, env.SESSION_PEPPER),
    createdAt:    new Date().toISOString(),
  }));

  return json({ success: true, email });
}

async function handleAdminLogin(request, env) {
  const problem = configError(env);
  if (problem) return json({ success: false, error: problem }, 500);

  let body;
  try {
    body = await request.json();
  } catch {
    return json({ success: false, error: 'Invalid request' }, 400);
  }

  const email    = String(body.email || '').trim().toLowerCase();
  const password = String(body.password || '');
  const fail = () => json({ success: false, error: 'Invalid email or password' }, 401);
  if (!email || !password) return fail();

  const raw = await env.REACHX_KV.get(`admin:${email}`);
  if (!raw) {
    await hashPassword(password, '00'.repeat(16), env.SESSION_PEPPER);  // same work on both paths
    return fail();
  }

  let admin;
  try {
    admin = JSON.parse(raw);
  } catch {
    return fail();
  }

  const computed = await hashPassword(password, admin.salt, env.SESSION_PEPPER);
  if (!safeEqual(computed, admin.passwordHash)) return fail();

  const token = randomHex(32);
  await env.REACHX_KV.put(
    `adminsession:${token}`,
    JSON.stringify({ email, at: new Date().toISOString() }),
    { expirationTtl: ADMIN_TTL_SECONDS },
  );

  return json({ success: true, admin: { email } }, 200, {
    'Set-Cookie': adminCookie(token, ADMIN_TTL_SECONDS),
  });
}

async function handleAdminLogout(request, env) {
  const token = getCookie(request, ADMIN_COOKIE_NAME);
  if (token) await env.REACHX_KV.delete(`adminsession:${token}`);
  return json({ success: true }, 200, { 'Set-Cookie': adminCookie('', 0) });
}

async function handleAdminChangePassword(request, env, admin) {
  let body;
  try {
    body = await request.json();
  } catch {
    return json({ success: false, error: 'Invalid request' }, 400);
  }

  const current = String(body.currentPassword || '');
  const next    = String(body.newPassword || '');
  if (next.length < 12) {
    return json({ success: false, error: 'New password must be at least 12 characters' }, 400);
  }

  const raw = await env.REACHX_KV.get(`admin:${admin.email}`);
  if (!raw) return json({ success: false, error: 'Admin record missing' }, 500);
  const rec = JSON.parse(raw);

  const computed = await hashPassword(current, rec.salt, env.SESSION_PEPPER);
  if (!safeEqual(computed, rec.passwordHash)) {
    return json({ success: false, error: 'Current password is wrong' }, 403);
  }

  const salt = randomHex(16);
  await env.REACHX_KV.put(`admin:${admin.email}`, JSON.stringify({
    ...rec,
    salt,
    passwordHash: await hashPassword(next, salt, env.SESSION_PEPPER),
    passwordChangedAt: new Date().toISOString(),
  }));

  return json({ success: true });
}

async function handleAdminListUsers(env) {
  const names = await listKeys(env, 'user:');
  const users = [];
  for (const name of names) {
    const raw = await env.REACHX_KV.get(name);
    if (!raw) continue;
    try {
      users.push(publicUser(JSON.parse(raw)));
    } catch { /* skip a record that is not JSON rather than failing the list */ }
  }
  users.sort((a, b) => a.email.localeCompare(b.email));
  return json({ success: true, users });
}

async function handleAdminSaveUser(request, env) {
  let body;
  try {
    body = await request.json();
  } catch {
    return json({ success: false, error: 'Invalid request' }, 400);
  }

  const email = String(body.email || '').trim().toLowerCase();
  if (!email || !email.includes('@')) {
    return json({ success: false, error: 'A valid email is required' }, 400);
  }

  const existing = await readUser(env, email);
  const creating = !existing;

  if (creating && String(body.password || '').length < 8) {
    return json({ success: false, error: 'New users need a password of at least 8 characters' }, 400);
  }

  const rec = existing || { email, createdAt: new Date().toISOString() };

  if (body.username  !== undefined) rec.username  = String(body.username).trim();
  if (body.masterUrl !== undefined) rec.masterUrl = String(body.masterUrl).trim();
  if (body.masterId  !== undefined) rec.masterId  = body.masterId;
  if (body.role      !== undefined) rec.role      = String(body.role).trim() || 'user';
  if (body.disabled  !== undefined) rec.disabled  = !!body.disabled;

  // Only rewrite the hash when a password was actually supplied, or every edit
  // to an unrelated field would silently reset it.
  if (body.password) {
    rec.salt         = randomHex(16);
    rec.passwordHash = await hashPassword(String(body.password), rec.salt, env.SESSION_PEPPER);
  }

  await env.REACHX_KV.put(`user:${email}`, JSON.stringify(rec));

  // Disabling, or moving the account to a different backend, must not leave a
  // live session running against the old state.
  let revoked = 0;
  if (rec.disabled || (existing && body.masterUrl !== undefined && body.masterUrl !== existing.masterUrl)) {
    revoked = await revokeUserSessions(env, email);
  }

  return json({ success: true, created: creating, user: publicUser(rec), sessionsRevoked: revoked });
}

async function handleAdminDeleteUser(request, env) {
  let body;
  try {
    body = await request.json();
  } catch {
    return json({ success: false, error: 'Invalid request' }, 400);
  }

  const email = String(body.email || '').trim().toLowerCase();
  if (!email) return json({ success: false, error: 'Missing email' }, 400);

  const revoked = await revokeUserSessions(env, email);

  // Keys outlive the account they were minted for unless they go too.
  const keyNames = await listKeys(env, 'apikey:');
  let keysDropped = 0;
  for (const name of keyNames) {
    const raw = await env.REACHX_KV.get(name);
    if (!raw) continue;
    try {
      if (JSON.parse(raw).email === email) {
        await env.REACHX_KV.delete(name);
        keysDropped++;
      }
    } catch { /* ignore */ }
  }

  await env.REACHX_KV.delete(`user:${email}`);
  return json({ success: true, sessionsRevoked: revoked, keysRevoked: keysDropped });
}

async function handleAdminListKeys(env) {
  const names = await listKeys(env, 'apikey:');
  const keys = [];
  for (const name of names) {
    const raw = await env.REACHX_KV.get(name);
    if (!raw) continue;
    try {
      const rec = JSON.parse(raw);
      keys.push({
        // Returned in full. An admin session can already impersonate any user
        // and reach every action on every backend, so withholding the key it
        // issued protects nothing — it only forced the operator to keep a copy
        // somewhere less safe than this page.
        key:       name.slice('apikey:'.length),
        email:     rec.email || '',
        label:     rec.label || '',
        createdAt: rec.createdAt || '',
      });
    } catch { /* ignore */ }
  }
  keys.sort((a, b) => (a.email || '').localeCompare(b.email || ''));
  return json({ success: true, keys });
}

async function handleAdminCreateKey(request, env) {
  let body;
  try {
    body = await request.json();
  } catch {
    return json({ success: false, error: 'Invalid request' }, 400);
  }

  const email = String(body.email || '').trim().toLowerCase();
  const user  = await readUser(env, email);
  if (!user) return json({ success: false, error: 'No such user: ' + email }, 404);

  const key = randomHex(32);
  await env.REACHX_KV.put(`apikey:${key}`, JSON.stringify({
    email:     user.email,
    username:  user.username || user.email,
    masterUrl: user.masterUrl || '',
    masterId:  user.masterId,
    role:      user.role || 'user',
    label:     String(body.label || '').trim(),
    createdAt: new Date().toISOString(),
  }));

  // The only time the full key is ever returned.
  return json({ success: true, key });
}

async function handleAdminRevokeKey(request, env) {
  let body;
  try {
    body = await request.json();
  } catch {
    return json({ success: false, error: 'Invalid request' }, 400);
  }
  const key = String(body.key || '').trim();
  if (!key) return json({ success: false, error: 'Missing key' }, 400);
  await env.REACHX_KV.delete(`apikey:${key}`);
  return json({ success: true });
}

/**
 * Starts an impersonation session: signs the admin into the app AS a user.
 *
 * This is why the admin console does not reimplement the dashboard, campaigns,
 * leads and settings screens. A second copy of all that would be a second copy
 * to keep correct, and the twin that drifts is the one nobody notices — the
 * same shape of bug that once silently stopped every follow-up. The admin gets
 * the real app instead, so there is exactly one of everything.
 *
 * The session is handed over through a single-use token rather than a cookie
 * set here, because the console may be on admin.example.com while the app is on
 * example.com, and a Strict cookie set on one is not sent to the other.
 *
 * The resulting session records impersonatedBy, so the app can say plainly
 * whose account is on screen and the record shows it was not the user acting.
 */
async function handleAdminImpersonate(request, env, admin, url) {
  let body;
  try {
    body = await request.json();
  } catch {
    return json({ success: false, error: 'Invalid request' }, 400);
  }

  const email = String(body.email || '').trim().toLowerCase();
  const user  = await readUser(env, email);
  if (!user) return json({ success: false, error: 'No such user: ' + email }, 404);
  if (!masterUrlFor(env, user)) {
    return json({ success: false, error: 'No backend configured for ' + email }, 400);
  }

  const token = randomHex(32);
  await env.REACHX_KV.put(
    `handoff:${token}`,
    JSON.stringify({ email, by: admin.email }),
    { expirationTtl: 60 },        // seconds: long enough to redirect, not to sit in a log
  );

  // admin.example.com → example.com. Same host when the console is at /admin.
  const appHost = url.hostname.startsWith('admin.') ? url.hostname.slice(6) : url.hostname;
  return json({ success: true, url: `https://${appHost}/api/impersonate?t=${token}` });
}

/**
 * Consumes a handoff token and lands the browser in the app as that user.
 *
 * Single use: the token is deleted before the session exists, so a replayed
 * link — from history, a proxy log, a shared screenshot — is inert.
 */
async function handleImpersonateHandoff(request, env, url) {
  const token = url.searchParams.get('t') || '';
  if (!token) return json({ success: false, error: 'Missing token' }, 400);

  const raw = await env.REACHX_KV.get(`handoff:${token}`);
  await env.REACHX_KV.delete(`handoff:${token}`);
  if (!raw) {
    return json({ success: false, error: 'This link has expired or was already used.' }, 403);
  }

  let handoff;
  try {
    handoff = JSON.parse(raw);
  } catch {
    return json({ success: false, error: 'Bad handoff record' }, 500);
  }

  const user = await readUser(env, handoff.email);
  if (!user) return json({ success: false, error: 'User no longer exists' }, 404);

  const sessionToken = randomHex(32);
  await env.REACHX_KV.put(
    `session:${sessionToken}`,
    JSON.stringify({
      email:          user.email,
      masterId:       user.masterId,
      masterUrl:      user.masterUrl || '',
      name:           user.username || user.name || user.email,
      role:           user.role || 'user',
      impersonatedBy: handoff.by,
    }),
    { expirationTtl: SESSION_TTL_SECONDS },
  );

  return new Response(null, {
    status: 303,
    headers: {
      Location: '/app',
      'Cache-Control': 'no-store',
      'Set-Cookie': sessionCookie(sessionToken, SESSION_TTL_SECONDS),
    },
  });
}

/**
 * Runs one action against any user's Master Control.
 *
 * Deliberately NOT allow-listed: admin is the operator, and the actions the
 * browser is barred from — adding a sender, rewriting a Config sheet — are
 * exactly the ones this exists for. The gate is the admin session, not the
 * action name.
 */
async function handleAdminCall(request, env) {
  let body;
  try {
    body = await request.json();
  } catch {
    return json({ success: false, error: 'Invalid request' }, 400);
  }

  const email  = String(body.email || '').trim().toLowerCase();
  const action = String(body.action || '').trim();
  if (!email || !action) return json({ success: false, error: 'email and action are required' }, 400);

  const user = await readUser(env, email);
  if (!user) return json({ success: false, error: 'No such user: ' + email }, 404);

  const target = masterUrlFor(env, user);
  if (!target) return json({ success: false, error: 'No backend configured for ' + email }, 400);

  try {
    const upstream = await fetch(target, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify({ ...(body.params || {}), action }),
      redirect: 'follow',
    });
    const text = await upstream.text();
    if (!upstream.ok) {
      return json({ success: false, error: `Backend returned HTTP ${upstream.status}` }, 502);
    }
    try {
      return json(JSON.parse(text));
    } catch {
      return json({ success: false, error: 'Backend did not return JSON.' }, 502);
    }
  } catch (err) {
    return json({ success: false, error: `Could not reach the backend: ${err.message}` }, 504);
  }
}

/* ── route handlers ────────────────────────────────────────────────────── */

/**
 * Reports what is and isn't wired up, without exposing any secret values.
 * Booleans and counts only — never a URL, never the pepper.
 */
function configReport(env) {
  const masters = [1, 2, 3, 4, 5].filter((n) => masterUrlFor(env, { masterId: n }));
  return {
    kvBound:          !!env.REACHX_KV,
    pepperSet:        !!env.SESSION_PEPPER,
    mastersConfigured: masters,
    // Booleans only. Reports whether the running deployment can see the admin
    // setup secret — a Pages secret added after a deploy is invisible until the
    // next one, and "is it set?" is otherwise unanswerable without guessing.
    adminSetupKeySet: !!env.ADMIN_SETUP_KEY,
  };
}

/** Blocks every auth path with one clear message when setup is incomplete. */
function configError(env) {
  if (!env.REACHX_KV) {
    return 'The KV namespace is not connected. In Cloudflare: Settings → Bindings → add a KV namespace binding named REACHX_KV, then redeploy.';
  }
  if (!env.SESSION_PEPPER) {
    return 'SESSION_PEPPER is not set. In Cloudflare: Settings → Variables and Secrets → add it as a Secret, then redeploy.';
  }
  return null;
}

async function handleLogin(request, env) {
  const problem = configError(env);
  if (problem) return json({ success: false, error: problem }, 500);

  let body;
  try {
    body = await request.json();
  } catch {
    return json({ success: false, error: 'Invalid request' }, 400);
  }

  const email    = String(body.email || '').trim().toLowerCase();
  const password = String(body.password || '');
  if (!email || !password) {
    return json({ success: false, error: 'Email and password are required' }, 400);
  }

  const raw = await env.REACHX_KV.get(`user:${email}`);

  // Same generic message and a real hash computation on both paths, so a
  // missing account is not distinguishable from a wrong password.
  const fail = () => json({ success: false, error: 'Invalid email or password' }, 401);

  if (!raw) {
    await hashPassword(password, '00'.repeat(16), env.SESSION_PEPPER);
    return fail();
  }

  let user;
  try {
    user = JSON.parse(raw);
  } catch {
    return fail();
  }

  if (user.disabled) return json({ success: false, error: 'This account is disabled' }, 403);

  const computed = await hashPassword(password, user.salt, env.SESSION_PEPPER);
  if (!safeEqual(computed, user.passwordHash)) return fail();

  if (!masterUrlFor(env, user)) {
    return json(
      { success: false,
        error: 'No backend is configured for your account. It needs either a masterUrl on its record, or a matching MASTER_<n>_URL setting.' },
      500,
    );
  }

  const token = randomHex(32);
  await env.REACHX_KV.put(
    `session:${token}`,
    JSON.stringify({
      email,
      masterId:  user.masterId,
      masterUrl: user.masterUrl || '',
      name:      user.username || user.name || email,
      role:      user.role || 'user',
    }),
    { expirationTtl: SESSION_TTL_SECONDS },
  );

  // Stamped on the record the admin screen reads, so "has this account ever
  // been used?" is answerable without digging through logs.
  await env.REACHX_KV.put(`user:${email}`, JSON.stringify({
    ...user,
    lastLogin: new Date().toISOString(),
  }));

  return json(
    { success: true, user: { email, name: user.name || email, masterId: user.masterId, role: user.role || 'user' } },
    200,
    { 'Set-Cookie': sessionCookie(token, SESSION_TTL_SECONDS) },
  );
}

async function handleLogout(request, env) {
  const token = getCookie(request, COOKIE_NAME);
  if (token) await env.REACHX_KV.delete(`session:${token}`);
  return json({ success: true }, 200, { 'Set-Cookie': sessionCookie('', 0) });
}

async function handleMe(session) {
  if (!session) return json({ success: false, error: 'Not signed in' }, 401);
  return json({
    success: true,
    user: {
      email: session.email,
      name: session.name,
      masterId: session.masterId,
      role: session.role,
      // Present only during impersonation. The app shows a banner from this —
      // an operator acting on a customer's account should never be able to
      // forget whose account it is.
      impersonatedBy: session.impersonatedBy || null,
    },
  });
}

/**
 * Proxies one action to the signed-in user's master control.
 *
 * Content-Type must be text/plain — Apps Script cannot set CORS headers, and
 * anything other than a simple request would need a preflight it can't answer.
 * (Server-to-server this is not strictly required, but keeping it identical to
 * the documented contract means the same payload works if ever called direct.)
 */
async function handleCall(request, env, session) {
  if (!session) return json({ success: false, error: 'Not signed in' }, 401);

  let body;
  try {
    body = await request.json();
  } catch {
    return json({ success: false, error: 'Invalid request' }, 400);
  }

  const action = String(body.action || '').trim();
  if (!action) return json({ success: false, error: 'Missing action' }, 400);
  if (!ALLOWED_ACTIONS.has(action)) {
    return json({ success: false, error: `Action not permitted: ${action}` }, 403);
  }

  const target = masterUrlFor(env, session);
  if (!target) {
    return json({ success: false, error: 'Backend not configured for your account' }, 500);
  }

  const payload = { ...(body.params || {}), action };

  try {
    const upstream = await fetch(target, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify(payload),
      redirect: 'follow',
    });

    const text = await upstream.text();

    // Apps Script answers 200 with a JSON body on every path. Anything else
    // means the deployment is broken, unshared, or has thrown outside doPost.
    if (!upstream.ok) {
      return json(
        { success: false, error: `Backend returned HTTP ${upstream.status}. Check the Apps Script deployment.` },
        502,
      );
    }

    try {
      return json(JSON.parse(text));
    } catch {
      // Apps Script served HTML rather than JSON. Report the possibilities
      // instead of asserting one — the earlier message blamed access settings
      // for what is usually a missing doGet or a script error.
      const looksLikeLogin = /accounts\.google\.com|ServiceLogin/i.test(text);
      const looksLikeScriptError = /Script function not found|doGet/i.test(text);

      let hint = 'The deployment may need redeploying, or the script threw before it could reply. Check its Executions log.';
      if (looksLikeLogin) {
        hint = 'It returned a Google sign-in page, so the deployment is not set to "Anyone" access.';
      } else if (looksLikeScriptError) {
        hint = 'It has no doGet handler and something issued a GET. Apply the doGet patch in BACKEND-PATCH.md.';
      }

      return json({ success: false, error: `Backend did not return JSON. ${hint}` }, 502);
    }
  } catch (err) {
    return json({ success: false, error: `Could not reach the backend: ${err.message}` }, 504);
  }
}

/* ── entry point ───────────────────────────────────────────────────────── */

export default {
  async fetch(request, env) {
    try {
      return await route(request, env);
    } catch (err) {
      // Anything uncaught would otherwise surface as a Cloudflare error page,
      // which reads to the user as "the site is down".
      const wantsJson = new URL(request.url).pathname.startsWith('/api/');
      if (wantsJson) return json({ success: false, error: `Server error: ${err.message}` }, 500);
      return new Response(
        `Reach X could not render this page.\n\n${err.message}\n\n` +
        `Check /api/health, then re-deploy if a binding is missing.`,
        { status: 500, headers: { 'Content-Type': 'text/plain; charset=utf-8', 'Cache-Control': 'no-store' } },
      );
    }
  },
};

async function route(request, env) {
  const url = new URL(request.url);

  if (url.pathname.startsWith('/api/')) {
    // Without this, any thrown error becomes a Cloudflare HTML error page,
    // which the browser cannot parse as JSON — the frontend then reports a
    // misleading "could not reach the server" instead of the real cause.
    try {
      // Setup check. Safe to leave enabled: booleans and counts only.
      if (url.pathname === '/api/health') {
        return json({ success: true, ...configReport(env) });
      }

      /* ── admin API ──────────────────────────────────────────────────────
       * Every route below the first three requires an admin session. The check
       * lives here, in one place, rather than inside each handler — a handler
       * added later cannot forget a gate it never had to write. */
      if (url.pathname.startsWith('/api/admin/')) {
        if (url.pathname === '/api/admin/bootstrap' && request.method === 'POST') {
          return handleAdminBootstrap(request, env);
        }
        if (url.pathname === '/api/admin/login' && request.method === 'POST') {
          return handleAdminLogin(request, env);
        }
        if (url.pathname === '/api/admin/logout' && request.method === 'POST') {
          return handleAdminLogout(request, env);
        }

        const admin = await readAdmin(request, env);
        const denied = requireAdmin(admin);
        if (denied) return denied;

        if (url.pathname === '/api/admin/me')                                     return json({ success: true, admin: { email: admin.email } });
        if (url.pathname === '/api/admin/password'   && request.method === 'POST') return handleAdminChangePassword(request, env, admin);
        if (url.pathname === '/api/admin/users'      && request.method === 'GET')  return handleAdminListUsers(env);
        if (url.pathname === '/api/admin/users'      && request.method === 'POST') return handleAdminSaveUser(request, env);
        if (url.pathname === '/api/admin/users/delete' && request.method === 'POST') return handleAdminDeleteUser(request, env);
        if (url.pathname === '/api/admin/keys'       && request.method === 'GET')  return handleAdminListKeys(env);
        if (url.pathname === '/api/admin/keys'       && request.method === 'POST') return handleAdminCreateKey(request, env);
        if (url.pathname === '/api/admin/keys/revoke' && request.method === 'POST') return handleAdminRevokeKey(request, env);
        if (url.pathname === '/api/admin/call'       && request.method === 'POST') return handleAdminCall(request, env);
        if (url.pathname === '/api/admin/impersonate' && request.method === 'POST') return handleAdminImpersonate(request, env, admin, url);

        return json({ success: false, error: 'Not found' }, 404);
      }

      // Consumes an admin-issued handoff token. Not admin-gated by cookie — the
      // single-use token IS the credential, and it was minted behind the gate.
      if (url.pathname === '/api/impersonate' && request.method === 'GET') {
        return handleImpersonateHandoff(request, env, url);
      }

      // Accepts an API key as well as a cookie, so automations (n8n, cron,
      // scripts) can call the same allow-listed actions the browser does.
      const session = await readAuth(request, env);

      if (url.pathname === '/api/login' && request.method === 'POST')  return handleLogin(request, env);
      if (url.pathname === '/api/logout' && request.method === 'POST') return handleLogout(request, env);
      if (url.pathname === '/api/me' && request.method === 'GET')      return handleMe(session);
      if (url.pathname === '/api/call' && request.method === 'POST')   return handleCall(request, env, session);

      return json({ success: false, error: 'Not found' }, 404);
    } catch (err) {
      return json({ success: false, error: `Server error: ${err.message}` }, 500);
    }
  }

  /* The admin console answers on its own hostname (admin.example.com) and at
     /admin on any hostname, so it works whether or not the subdomain is set up
     yet. It is NOT gated here: the page itself is a login form, and every byte
     of data behind it comes from /api/admin/*, which is. */
  const isAdminHost = url.hostname.startsWith('admin.');
  if (url.pathname === '/admin' || url.pathname === '/admin.html' ||
      (isAdminHost && (url.pathname === '/' || url.pathname === '/index.html'))) {
    const page = await env.ASSETS.fetch(
      new Request(new URL('/admin.html', url).toString(), { method: 'GET' }),
    );
    if (page.status !== 404) return page;
  }

  // On the admin hostname the tenant app is not reachable at all — one less
  // surface, and no chance of a half-signed-in state across two cookies.
  if (isAdminHost) {
    return new Response('Not found', { status: 404, headers: { 'Content-Type': 'text/plain' } });
  }

  // Gate the app shell so an unauthenticated visitor never sees it.
  if (url.pathname === '/app' || url.pathname === '/app.html') {
    let session = null;
    try {
      session = await readSession(request, env);
    } catch {
      session = null;              // a KV hiccup must not brick the page
    }

    if (!session) {
      // 303 so the browser always follows with GET, and no-store so this
      // redirect is never cached — a cached one would survive signing in.
      return new Response(null, {
        status: 303,
        headers: { Location: '/', 'Cache-Control': 'no-store' },
      });
    }

    // Return the asset response as-is. Rebuilding it risks re-applying a
    // content-encoding header to a body that is already being streamed.
    const shell = await env.ASSETS.fetch(
      new Request(new URL('/app.html', url).toString(), { method: 'GET' }),
    );
    if (shell.status === 404) return env.ASSETS.fetch(request);
    return shell;
  }

  return env.ASSETS.fetch(request);
}
