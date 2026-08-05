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
      name:      user.name || email,
      role:      user.role || 'user',
    }),
    { expirationTtl: SESSION_TTL_SECONDS },
  );

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

      const session = await readSession(request, env);

      if (url.pathname === '/api/login' && request.method === 'POST')  return handleLogin(request, env);
      if (url.pathname === '/api/logout' && request.method === 'POST') return handleLogout(request, env);
      if (url.pathname === '/api/me' && request.method === 'GET')      return handleMe(session);
      if (url.pathname === '/api/call' && request.method === 'POST')   return handleCall(request, env, session);

      return json({ success: false, error: 'Not found' }, 404);
    } catch (err) {
      return json({ success: false, error: `Server error: ${err.message}` }, 500);
    }
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
