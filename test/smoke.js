'use strict';

// Smoke test: starts a fake Unity Gateway and the proxy, then checks CORS,
// auth forwarding, streaming, route allowlist and error responses.
// Run with: npm test

const assert = require('node:assert/strict');
const http = require('node:http');
const path = require('node:path');
const { spawn } = require('node:child_process');

const ORIGIN = 'https://pivot.claude.ai';
const PROXY_PORT = 18080;
const EXTRA_MODELS = 'dz_hyp.default.claude-5-5, system.ai.claude-sonnet-5-5';
let lastUpstreamRequest;

const upstream = http.createServer((req, res) => {
  lastUpstreamRequest = { method: req.method, url: req.url, headers: req.headers };
  if (req.headers.authorization !== 'Bearer dapi-test') {
    res.writeHead(401, { 'content-type': 'application/json', 'access-control-allow-origin': '*' });
    return res.end(JSON.stringify({ error_code: 'UNAUTHENTICATED' }));
  }
  if (req.url === '/ai-gateway/anthropic/v1/models') {
    res.writeHead(200, { 'content-type': 'application/json' });
    return res.end(JSON.stringify({ data: [{ id: 'system.ai.claude-sonnet-5-5' }], has_more: false }));
  }
  if (req.url === '/ai-gateway/anthropic/v1/messages') {
    res.writeHead(200, { 'content-type': 'text/event-stream', 'request-id': 'req_123' });
    res.write('event: message_start\ndata: {}\n\n');
    setTimeout(() => res.end('event: message_stop\ndata: {}\n\n'), 50);
    return;
  }
  res.writeHead(500);
  res.end('unexpected upstream path');
});

async function main() {
  await new Promise((resolve) => upstream.listen(0, resolve));
  const upstreamUrl = `http://127.0.0.1:${upstream.address().port}/ai-gateway/anthropic/`;

  const proxy = spawn(process.execPath, [path.join(__dirname, '..', 'server.js')], {
    env: { ...process.env, PORT: String(PROXY_PORT), UPSTREAM_URL: upstreamUrl, EXTRA_MODELS },
    stdio: ['ignore', 'pipe', 'inherit'],
  });
  await new Promise((resolve) => proxy.stdout.once('data', resolve));
  const base = `http://127.0.0.1:${PROXY_PORT}`;

  try {
    // 1. Preflight is answered without auth and reflects requested headers.
    let res = await fetch(`${base}/v1/messages`, {
      method: 'OPTIONS',
      headers: {
        origin: ORIGIN,
        'access-control-request-method': 'POST',
        'access-control-request-headers': 'authorization, content-type, anthropic-version',
      },
    });
    assert.equal(res.status, 204);
    assert.equal(res.headers.get('access-control-allow-origin'), ORIGIN);
    assert.match(res.headers.get('access-control-allow-headers'), /authorization/);
    assert.match(res.headers.get('access-control-allow-methods'), /POST/);

    // 2. Unknown origin is rejected.
    res = await fetch(`${base}/v1/messages`, {
      method: 'OPTIONS',
      headers: { origin: 'https://evil.example.com', 'access-control-request-method': 'POST' },
    });
    assert.equal(res.status, 403);
    assert.equal(res.headers.get('access-control-allow-origin'), null);

    // 3. x-api-key is converted to Bearer, response is streamed with CORS headers,
    //    browser headers are stripped before reaching Databricks.
    res = await fetch(`${base}/v1/messages`, {
      method: 'POST',
      headers: { origin: ORIGIN, 'x-api-key': 'dapi-test', 'content-type': 'application/json', 'anthropic-version': '2023-06-01' },
      body: JSON.stringify({ model: 'system.ai.claude-sonnet-5-5', max_tokens: 10, stream: true, messages: [] }),
    });
    assert.equal(res.status, 200);
    assert.equal(res.headers.get('access-control-allow-origin'), ORIGIN);
    assert.equal(res.headers.get('request-id'), 'req_123');
    assert.match(await res.text(), /message_start[\s\S]*message_stop/);
    assert.equal(lastUpstreamRequest.headers.authorization, 'Bearer dapi-test');
    assert.equal(lastUpstreamRequest.headers['x-api-key'], undefined);
    assert.equal(lastUpstreamRequest.headers.origin, undefined);
    assert.equal(lastUpstreamRequest.headers['anthropic-version'], '2023-06-01');

    // 4. Model discovery with Authorization header: gateway list plus EXTRA_MODELS,
    //    without duplicating models the gateway already lists.
    res = await fetch(`${base}/v1/models`, { headers: { origin: ORIGIN, authorization: 'Bearer dapi-test' } });
    assert.equal(res.status, 200);
    assert.equal(res.headers.get('access-control-allow-origin'), ORIGIN);
    const models = (await res.json()).data;
    assert.deepEqual(models.map((m) => m.id), ['system.ai.claude-sonnet-5-5', 'dz_hyp.default.claude-5-5']);
    assert.equal(models[1].type, 'model');
    assert.equal(models[1].display_name, 'dz_hyp.default.claude-5-5');

    // 5. Upstream errors keep our CORS header (and drop upstream's wildcard).
    res = await fetch(`${base}/v1/models`, { headers: { origin: ORIGIN, authorization: 'Bearer wrong' } });
    assert.equal(res.status, 401);
    assert.equal(res.headers.get('access-control-allow-origin'), ORIGIN);

    // 6. Other paths are blocked (no open proxy to workspace APIs).
    res = await fetch(`${base}/api/2.0/clusters/list`, { headers: { origin: ORIGIN, authorization: 'Bearer dapi-test' } });
    assert.equal(res.status, 404);
    assert.equal(res.headers.get('access-control-allow-origin'), ORIGIN);

    // 7. Health check.
    res = await fetch(`${base}/healthz`);
    assert.equal(res.status, 200);

    console.log('All smoke tests passed');
  } finally {
    proxy.kill();
    upstream.close();
  }
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
