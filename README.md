# claude-m365-cors-proxy

Small reverse proxy that lets the **Claude for Microsoft 365** Office add-ins
(Excel, PowerPoint, Word, Outlook) use **Databricks Unity Gateway** as their
LLM gateway.

The add-in's task pane is served from `https://pivot.claude.ai` and calls the
gateway directly from the browser, so every request is cross-origin. Unity
Gateway doesn't answer the unauthenticated CORS preflight or send CORS headers,
so the browser blocks the responses. This proxy fixes exactly that and nothing
else:

```
Office add-in (https://pivot.claude.ai)
   │  OPTIONS  → answered by the proxy (no auth, CORS headers)
   │  POST /v1/messages, GET /v1/models → forwarded unchanged
   ▼
claude-m365-cors-proxy (Azure App Service)
   ▼
https://<workspace>/ai-gateway/anthropic   (auth, governance, logging stay here)
```

What it does:

- Answers `OPTIONS` preflights for allowed origins (default `https://pivot.claude.ai`).
- Forwards only `POST /v1/messages`, `POST /v1/messages/count_tokens` and
  `GET /v1/models` to `UPSTREAM_URL`. Other paths return 404, so it can't be
  used to reach other workspace APIs.
- Adds `Access-Control-Allow-Origin` to every response, including upstream
  errors and proxy errors.
- Streams responses (SSE) without buffering and cancels the upstream call when
  the user stops a generation.
- Converts `x-api-key` (the add-in's default auth header) to
  `Authorization: Bearer`, which Databricks expects.
- Strips browser headers (`Origin`, `Referer`, `Cookie`) before forwarding.

What it doesn't do: it stores no credentials, issues no tokens and doesn't log
request bodies. The user's Databricks token is passed through, and Unity Gateway
still authenticates and authorizes every call.

No npm dependencies. Requires Node.js 20 or later.

## Configuration

| Variable | Required | Default | Description |
| --- | --- | --- | --- |
| `UPSTREAM_URL` | yes | – | Unity Gateway Anthropic endpoint, e.g. `https://adb-1234567890123456.7.azuredatabricks.net/ai-gateway/anthropic` |
| `ALLOWED_ORIGINS` | no | `https://pivot.claude.ai` | Comma-separated list of browser origins allowed to call the proxy |
| `EXTRA_MODELS` | no | – | Models to add to the gateway's `GET /v1/models` list, see [Custom models](#custom-models) |
| `REQUIRE_ORIGIN` | no | `false` | Set to `true` to reject requests without an `Origin` header (blocks curl/server-side use) |
| `UPSTREAM_TIMEOUT_MS` | no | `600000` | Socket idle timeout towards the gateway |
| `PORT` | no | `8080` | Set automatically by App Service |

## Custom models

Unity Gateway's `GET /v1/models` lists only the built-in `system.ai.*` models.
Models it serves under other names, e.g. a model registered in a Unity Catalog
schema such as `dz_hyp.default.claude-5-5`, work in `POST /v1/messages` but
don't show up in the add-in's model picker. `EXTRA_MODELS` appends them to the
gateway's list (models the gateway already lists aren't duplicated):

```bash
# comma-separated ids
EXTRA_MODELS=dz_hyp.default.claude-5-5

# or JSON, to set the label shown in the picker
EXTRA_MODELS='[{"id":"dz_hyp.default.claude-5-5","display_name":"Claude 5.5 (DZ HYP)"}]'
```

The list is only extended when the gateway answers with 200, so users without
gateway access still get the gateway's error. Whether a user may call a model is
still decided by Unity Gateway.

## Run locally

```bash
UPSTREAM_URL=https://<workspace-host>/ai-gateway/anthropic npm start
npm test   # smoke test against a fake gateway, no network needed
```

Quick check against a real workspace (token = Databricks OAuth access token or PAT):

```bash
curl -i -X OPTIONS http://localhost:8080/v1/messages \
  -H 'Origin: https://pivot.claude.ai' \
  -H 'Access-Control-Request-Method: POST' \
  -H 'Access-Control-Request-Headers: authorization, content-type, anthropic-version'

curl -N http://localhost:8080/v1/messages \
  -H 'Origin: https://pivot.claude.ai' \
  -H "Authorization: Bearer $DATABRICKS_TOKEN" \
  -H 'anthropic-version: 2023-06-01' -H 'content-type: application/json' \
  -d '{"model":"<model-name>","max_tokens":50,"stream":true,"messages":[{"role":"user","content":"Hi"}]}'
```

## Deploy to Azure App Service

Linux App Service, Node 22. A **B1** plan is enough for a PoC; the free **F1**
tier also works for a quick test, but it has a daily CPU quota and no Always On.

```bash
RG=rg-claude-m365-proxy
LOCATION=germanywestcentral
PLAN=asp-claude-m365-proxy
APP=claude-m365-proxy-<unique-suffix>      # becomes https://$APP.azurewebsites.net
UPSTREAM_URL=https://<workspace-host>/ai-gateway/anthropic

az group create -n $RG -l $LOCATION
az appservice plan create -g $RG -n $PLAN --is-linux --sku B1
az webapp create -g $RG -p $PLAN -n $APP --runtime "NODE:22-lts"

az webapp config appsettings set -g $RG -n $APP --settings \
  UPSTREAM_URL=$UPSTREAM_URL \
  ALLOWED_ORIGINS=https://pivot.claude.ai \
  SCM_DO_BUILD_DURING_DEPLOYMENT=false

az webapp config set -g $RG -n $APP \
  --startup-file "node server.js" \
  --always-on true \
  --min-tls-version 1.2 \
  --ftps-state Disabled \
  --generic-configurations '{"healthCheckPath": "/healthz"}'
az webapp update -g $RG -n $APP --https-only true

zip -r app.zip server.js package.json
az webapp deploy -g $RG -n $APP --src-path app.zip --type zip

curl https://$APP.azurewebsites.net/healthz   # -> ok
```

Don't enable App Service's built-in CORS feature (`az webapp cors`) in
addition: the proxy sets the headers itself, and both together produce
duplicate headers that browsers reject.

## Configure the add-in

In the add-in manifest (generated with the `claude-for-msft-365-install`
setup wizard), or on the add-in's sign-in screen under
**Cloud provider or gateway → Gateway**:

| Setting | Value |
| --- | --- |
| `gateway_url` | `https://<app>.azurewebsites.net` (no `/v1` suffix) |
| `gateway_api_format` | `anthropic` |
| `gateway_auth_header` | `authorization` (recommended; `x-api-key` also works, the proxy converts it) |
| Token | A Databricks token valid for Unity Gateway (see below) |

Model names: the add-in discovers models via `GET /v1/models`. If the gateway
doesn't return a list there, the add-in asks for a model ID; enter the full
Unity Catalog name of the Claude model.

## Tested

Run locally against an Azure Databricks workspace's Unity Gateway (2026-10-12),
authenticating with a Microsoft Entra ID token for the Azure Databricks
resource (`az account get-access-token --resource 2ff814a6-3304-4ab8-85cb-cd0e6f879c1d`):

| Check | Result |
| --- | --- |
| `OPTIONS /v1/messages` preflight | 204 with CORS headers |
| `GET /v1/models` | 200, returns `system.ai.claude-*` models |
| `POST /v1/messages` streamed, token sent as `x-api-key` | 200, SSE streamed, CORS header present |
| Invalid token | 401 from Unity Gateway, CORS header present |
| `GET /api/2.0/clusters/list` | 404 from the proxy (blocked) |

So Unity Gateway accepts Entra ID tokens directly. Not yet tested: the add-in
itself in Excel.

## Things to check before a PoC

- **Token.** On the LLM gateway path the add-in needs a gateway token. A
  token typed into the add-in is stored in the browser and isn't refreshed, so
  short-lived Databricks OAuth or Entra tokens stop working after about an
  hour; only a long-lived token (in practice a PAT) works without re-login.
  Two PAT-free options with automatic refresh:
  - **Entra token as gateway credential** (documented in the
    [`claude-for-msft-365-install` manifest docs](https://github.com/anthropics/financial-services/blob/main/claude-for-msft-365-install/commands/manifest.md#use-the-entra-token-as-your-gateway-credential)).
    Unity Gateway accepts Entra tokens for the Azure Databricks resource (see
    above), so this should work without any extra service. Not yet tested in
    the add-in:
    ```text
    gateway_url=https://<app>.azurewebsites.net
    gateway_api_format=anthropic
    entra_sso=1
    graph_client_id=<your Entra app; delegated Azure Databricks user_impersonation, admin-consented>
    entra_scope=2ff814a6-3304-4ab8-85cb-cd0e6f879c1d/.default
    gateway_auth_source=entra
    ```
  - **Bootstrap endpoint** (`bootstrap_url`): an HTTPS GET endpoint you run
    that receives the user's Entra token and returns per-user config, e.g. a
    short-lived `gateway_token` with `bootstrap_expires_at`, plus MCP servers
    and skills. Needed for per-user MCP servers and skills, or when the gateway
    doesn't accept Entra tokens. See the
    [bootstrap contract](https://github.com/anthropics/financial-services/blob/main/claude-for-msft-365-install/commands/bootstrap.md)
    and [Python reference](https://github.com/anthropics/financial-services/tree/main/claude-for-msft-365-install/examples/python-bootstrap).
- **Network.** The proxy calls the workspace from App Service's outbound IPs
  (`az webapp show -g $RG -n $APP --query outboundIpAddresses`). Add them to
  the workspace IP access list if one is enabled. With front-end Private Link,
  use App Service VNet integration plus private DNS so the proxy reaches the
  workspace privately.
- **Who can reach the proxy.** Users' browsers call the proxy directly, so
  App Service access restrictions can limit it to the corporate egress IPs.
- **Timeouts.** App Service's front end closes requests that send no response
  for about 230 seconds. Streamed responses are fine because data keeps
  flowing; very long non-streamed requests may hit that limit.
- **Logging.** The proxy logs one line per request (method, path, status,
  duration) to stdout, visible with `az webapp log tail`. Prompts and tokens
  are never logged.

This is a PoC helper, not a supported Databricks or Anthropic component.
