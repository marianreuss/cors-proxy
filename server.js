'use strict';

// Minimal CORS proxy so the Claude for Microsoft 365 Office add-ins (task pane
// served from https://pivot.claude.ai) can call Databricks Unity Gateway from
// the browser. The proxy holds no credentials: the caller's Databricks token is
// forwarded unchanged and Unity Gateway still does authN/authZ.
//
// Zero dependencies on purpose (Node >= 20).

const http = require('node:http');
const https = require('node:https');

const PORT = Number(process.env.PORT) || 8080;
const UPSTREAM = new URL(requiredEnv('UPSTREAM_URL')); // e.g. https://adb-123.11.azuredatabricks.net/ai-gateway/anthropic
const UPSTREAM_BASE_PATH = UPSTREAM.pathname.replace(/\/+$/, '');
const ALLOWED_ORIGINS = (process.env.ALLOWED_ORIGINS || 'https://pivot.claude.ai')
  .split(',')
  .map((s) => s.trim())
  .filter(Boolean);
const UPSTREAM_TIMEOUT_MS = Number(process.env.UPSTREAM_TIMEOUT_MS) || 10 * 60 * 1000;
const REQUIRE_ORIGIN = process.env.REQUIRE_ORIGIN === 'true';
// Models the gateway serves but doesn't list in GET /v1/models (e.g. models
// registered in a Unity Catalog schema). Appended to the gateway's list.
const EXTRA_MODELS = parseExtraModels(process.env.EXTRA_MODELS);

// Only the Anthropic Messages API surface the add-in uses. Everything else is
// rejected so the proxy can't be used to reach other workspace APIs.
const ALLOWED_ROUTES = [
  { method: 'POST', path: /^\/v1\/messages$/ },
  { method: 'POST', path: /^\/v1\/messages\/count_tokens$/ },
  { method: 'GET', path: /^\/v1\/models(\/[^/]+)?$/ },
];

const HOP_BY_HOP = new Set([
  'connection', 'keep-alive', 'proxy-authenticate', 'proxy-authorization',
  'te', 'trailer', 'transfer-encoding', 'upgrade', 'host',
]);
// Browser context headers Databricks doesn't need (and may treat as CSRF signals).
const STRIP_FROM_REQUEST = new Set([...HOP_BY_HOP, 'origin', 'referer', 'cookie', 'x-api-key']);
const STRIP_FROM_RESPONSE = new Set([...HOP_BY_HOP, 'set-cookie']);

const upstreamClient = UPSTREAM.protocol === 'http:' ? http : https;
const upstreamAgent = new upstreamClient.Agent({ keepAlive: true });

function requiredEnv(name) {
  const value = process.env[name];
  if (!value) {
    console.error(`Missing required environment variable ${name}`);
    process.exit(1);
  }
  return value;
}

// Accepts "id1,id2" or a JSON array of ids and/or {id, display_name} objects.
function parseExtraModels(value) {
  if (!value || !value.trim()) return [];
  let entries;
  if (value.trim().startsWith('[')) {
    try {
      entries = JSON.parse(value);
    } catch (err) {
      console.error(`EXTRA_MODELS is not valid JSON: ${err.message}`);
      process.exit(1);
    }
  } else {
    entries = value.split(',').map((s) => s.trim()).filter(Boolean);
  }
  return entries.map((entry) => {
    const { id, display_name: displayName } = typeof entry === 'string' ? { id: entry } : entry;
    if (!id) {
      console.error(`EXTRA_MODELS entry without id: ${JSON.stringify(entry)}`);
      process.exit(1);
    }
    // Same shape Unity Gateway returns for its own models.
    return { type: 'model', id, display_name: displayName || id, description: id, created_at: '' };
  });
}

function corsHeaders(req) {
  const origin = req.headers.origin;
  if (!origin || !ALLOWED_ORIGINS.includes(origin)) return {};
  return {
    'access-control-allow-origin': origin,
    'access-control-expose-headers': '*',
    vary: 'Origin',
  };
}

function sendError(res, status, cors, message) {
  // Anthropic-style error body so the add-in can surface the message.
  const body = JSON.stringify({ type: 'error', error: { type: 'proxy_error', message } });
  res.writeHead(status, { ...cors, 'content-type': 'application/json', 'content-length': Buffer.byteLength(body) });
  res.end(body);
}

function handlePreflight(req, res, cors) {
  if (!cors['access-control-allow-origin']) return sendError(res, 403, {}, 'Origin not allowed');
  res.writeHead(204, {
    ...cors,
    // Reflect requested headers: the add-in may send anthropic-* and SDK headers.
    'access-control-allow-headers':
      req.headers['access-control-request-headers'] ||
      'authorization, x-api-key, content-type, anthropic-version, anthropic-beta',
    'access-control-allow-methods': 'GET, POST, OPTIONS',
    'access-control-max-age': '600',
    vary: 'Origin, Access-Control-Request-Headers',
  });
  res.end();
}

function forward(req, res, cors, pathname, search) {
  const mergeModels = EXTRA_MODELS.length > 0 && req.method === 'GET' && pathname === '/v1/models';
  const headers = {};
  for (const [name, value] of Object.entries(req.headers)) {
    if (!STRIP_FROM_REQUEST.has(name)) headers[name] = value;
  }
  // The add-in sends the token as x-api-key by default; Databricks expects a Bearer token.
  if (!headers.authorization && req.headers['x-api-key']) {
    headers.authorization = `Bearer ${req.headers['x-api-key']}`;
  }
  // The model list is rewritten, so ask for it uncompressed.
  if (mergeModels) delete headers['accept-encoding'];

  const upstreamReq = upstreamClient.request(
    {
      protocol: UPSTREAM.protocol,
      hostname: UPSTREAM.hostname,
      port: UPSTREAM.port || undefined,
      method: req.method,
      path: `${UPSTREAM_BASE_PATH}${pathname}${search}`,
      headers,
      agent: upstreamAgent,
      timeout: UPSTREAM_TIMEOUT_MS,
    },
    (upstreamRes) => {
      const out = {};
      for (const [name, value] of Object.entries(upstreamRes.headers)) {
        if (!STRIP_FROM_RESPONSE.has(name) && !name.startsWith('access-control-')) out[name] = value;
      }
      if (mergeModels && upstreamRes.statusCode === 200) return sendMergedModels(upstreamRes, res, out, cors);
      res.writeHead(upstreamRes.statusCode, { ...out, ...cors });
      res.flushHeaders();
      upstreamRes.pipe(res); // streams SSE chunk by chunk
    },
  );

  upstreamReq.on('timeout', () => upstreamReq.destroy(new Error('upstream timeout')));
  upstreamReq.on('error', (err) => {
    if (!res.headersSent) sendError(res, 502, cors, `Upstream request failed: ${err.message}`);
    else res.destroy(err);
  });
  // Client went away (e.g. user pressed stop): cancel the upstream call too.
  res.on('close', () => {
    if (!res.writableFinished) upstreamReq.destroy();
  });

  req.pipe(upstreamReq);
}

function sendMergedModels(upstreamRes, res, headers, cors) {
  const chunks = [];
  upstreamRes.on('data', (chunk) => chunks.push(chunk));
  upstreamRes.on('end', () => {
    const raw = Buffer.concat(chunks);
    let body = raw;
    try {
      const list = JSON.parse(raw.toString('utf8'));
      // Only append on the last page so paginated callers see each model once.
      if (Array.isArray(list.data) && !list.has_more) {
        const known = new Set(list.data.map((m) => m.id));
        list.data.push(...EXTRA_MODELS.filter((m) => !known.has(m.id)));
        body = Buffer.from(JSON.stringify(list));
      }
    } catch {
      // Not JSON: pass the gateway's response through unchanged.
    }
    delete headers['content-encoding'];
    res.writeHead(200, { ...headers, ...cors, 'content-length': body.length });
    res.end(body);
  });
  upstreamRes.on('error', (err) => res.destroy(err));
}

const server = http.createServer((req, res) => {
  const started = Date.now();
  const { pathname, search } = new URL(req.url, 'http://localhost');
  res.on('finish', () => {
    console.log(`${req.method} ${pathname} ${res.statusCode} ${Date.now() - started}ms`);
  });

  if (pathname === '/healthz') {
    res.writeHead(200, { 'content-type': 'text/plain' });
    return res.end('ok');
  }

  const cors = corsHeaders(req);
  if (req.method === 'OPTIONS') return handlePreflight(req, res, cors);

  if (req.headers.origin && !cors['access-control-allow-origin']) {
    return sendError(res, 403, {}, 'Origin not allowed');
  }
  if (REQUIRE_ORIGIN && !req.headers.origin) {
    return sendError(res, 403, {}, 'Origin header required');
  }
  if (!ALLOWED_ROUTES.some((r) => r.method === req.method && r.path.test(pathname))) {
    return sendError(res, 404, cors, `Route not allowed: ${req.method} ${pathname}`);
  }

  forward(req, res, cors, pathname, search);
});

// Long-running streamed generations: don't let Node cut requests off.
server.requestTimeout = 0;
server.keepAliveTimeout = 65 * 1000;

server.listen(PORT, () => {
  console.log(`claude-m365-cors-proxy listening on :${PORT} -> ${UPSTREAM.origin}${UPSTREAM_BASE_PATH}`);
  console.log(`allowed origins: ${ALLOWED_ORIGINS.join(', ')}`);
  if (EXTRA_MODELS.length) console.log(`extra models: ${EXTRA_MODELS.map((m) => m.id).join(', ')}`);
});

process.on('SIGTERM', () => server.close(() => process.exit(0)));
