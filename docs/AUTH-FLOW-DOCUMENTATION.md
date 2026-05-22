# OpenExcel Authentication Flow Documentation

**Version:** 1.0  
**Last Updated:** April 20, 2026  
**Status:** Current, Implemented & Production-Ready

---

## Executive Summary

The **OpenExcel** application implements a **client-side OAuth 2.0 authentication flow with PKCE (Proof Key for Code Exchange)** to securely obtain and manage access tokens for LLM providers. The architecture follows a **BYOK (Bring Your Own Key)** model where users authenticate directly with their chosen OAuth provider (Anthropic Claude, OpenAI ChatGPT) and the application manages token lifecycle entirely on the client-side.

### Key Characteristics
- **Protocol:** OAuth 2.0 with PKCE (public client, browser-based)
- **Supported Providers:** Anthropic Claude Pro/Max, OpenAI ChatGPT Plus/Pro
- **Token Storage:** Browser localStorage + IndexedDB (scoped by document/workbook)
- **Token Refresh:** Proactive (checked before each LLM API call)
- **Architecture:** Client-side only (no backend OAuth relay required)
- **Security Model:** PKCE for preventing authorization code interception in redirects

---

## Architecture Overview

### High-Level Components

```
┌─────────────────────────────────────────────────────────────────┐
│                     USER BROWSER/EXCEL                          │
├─────────────────────────────────────────────────────────────────┤
│                                                                 │
│  ┌─────────────────────────────────────────────────────────┐  │
│  │  Svelte Settings Panel (settings-panel.svelte)         │  │
│  │  - OAuth provider selection                            │  │
│  │  - Login/logout UX                                     │  │
│  │  - Token status display                                │  │
│  └─────────────────────────────────────────────────────────┘  │
│                           │                                     │
│  ┌────────────────────────▼─────────────────────────────────┐  │
│  │  Client Storage (localStorage + IndexedDB)              │  │
│  │  - OAuth credentials: {refresh, access, expires}        │  │
│  │  - Provider config: provider, model, authMethod         │  │
│  │  - Session data: chat messages, VFS files, skills       │  │
│  └─────────────────────────────────────────────────────────┘  │
│                                                                 │
│  ┌─────────────────────────────────────────────────────────┐  │
│  │  Agent Runtime (runtime.ts)                            │  │
│  │  - getActiveApiKey(): Check token expiry & refresh     │  │
│  │  - Stream LLM messages with current access token        │  │
│  └─────────────────────────────────────────────────────────┘  │
│                                                                 │
│  ┌─────────────────────────────────────────────────────────┐  │
│  │  OAuth Handler (packages/sdk/src/oauth/index.ts)       │  │
│  │  - generatePKCE(): Create code challenge/verifier       │  │
│  │  - buildAuthorizationUrl(): Generate OAuth redirect     │  │
│  │  - exchangeOAuthCode(): Token endpoint request          │  │
│  │  - refreshOAuthToken(): Refresh token endpoint          │  │
│  └─────────────────────────────────────────────────────────┘  │
│                                                                 │
└─────────────────────────────────────────────────────────────────┘
                           │
        ┌──────────────────┼──────────────────┐
        │                  │                  │
        ▼                  ▼                  ▼
   ┌─────────┐    ┌──────────────┐    ┌────────────┐
   │Anthropic│    │    OpenAI    │    │ LLM APIs   │
   │ OAuth   │    │    OAuth     │    │            │
   │ Endpoints│   │  Endpoints   │    │  Message   │
   │          │    │              │    │  Endpoints│
   └─────────┘    └──────────────┘    └────────────┘
```

### Component Interactions

1. **Settings Panel (UI):** User selects OAuth provider and initiates login
2. **OAuth Handler:** Generates PKCE parameters and redirects to provider
3. **Provider:** User authenticates and returns authorization code
4. **OAuth Handler:** Exchanges code for tokens using PKCE verifier
5. **Storage:** Credentials persisted in localStorage (encrypted by browser)
6. **Runtime:** Before each API call, checks token expiry and refreshes if needed
7. **LLM APIs:** Receive bearer token in Authorization header

---

## Detailed Authentication Flow

### Phase 1: User Initiates OAuth (Initial Login)

**Trigger:** User clicks "Login with Provider" in settings panel

**Steps:**
1. User selects provider (Anthropic or OpenAI) from settings dropdown
2. Svelte component calls `initializeOAuth(provider)`
3. OAuth handler generates PKCE parameters:
   - `code_verifier`: 43-128 character random string (base64url encoded)
   - `code_challenge`: SHA-256 hash of verifier (base64url encoded)
4. OAuth handler stores `code_verifier` and random `state` in localStorage
5. OAuth handler builds authorization URL with:
   - `client_id`: Embedded provider client ID (base64-encoded in source)
   - `redirect_uri`: Provider's callback URL
   - `code_challenge` and `code_challenge_method=S256`
   - `state`: CSRF protection token
   - `scope`: Provider-specific OAuth scopes
6. Svelte redirects user to OAuth provider's authorization endpoint

### Phase 2: User Authenticates with Provider

**User Actions:**
1. User logs into their provider account (if not already authenticated)
2. Grants permission to OpenExcel application
3. Provider redirects back with authorization code + state

**Security Checks:**
- `state` parameter is validated (CSRF protection)
- Authorization code is single-use
- Code is only valid for ~10 minutes

### Phase 3: Exchange Authorization Code for Tokens

**Trigger:** Authorization code received at redirect URI

**Steps:**
1. Svelte component retrieves stored `code_verifier` and `state` from localStorage
2. Svelte calls `exchangeOAuthCode(code, verifier)`
3. OAuth handler POST to provider's token endpoint:
   ```json
   {
     "grant_type": "authorization_code",
     "code": "...",
     "code_verifier": "...",
     "client_id": "...",
     "redirect_uri": "..."
   }
   ```
4. Provider validates code_verifier against stored challenge (PKCE)
5. Provider returns tokens:
   ```json
   {
     "access_token": "...",
     "refresh_token": "...",
     "expires_in": 3600,
     "token_type": "Bearer"
   }
   ```
6. OAuth handler stores credentials in localStorage:
   ```json
   {
     "refresh": "refresh_token_value",
     "access": "access_token_value",
     "expires": 1713619200000  // timestamp in milliseconds
   }
   ```
7. Svelte updates UI to "Connected" state

### Phase 4: Token Refresh (Proactive Refresh)

**Trigger:** Before each LLM API call

**Steps:**
1. Runtime calls `getActiveApiKey(config)` to retrieve current access token
2. Runtime checks: `Date.now() < credentials.expires`?
3. **If expired or missing:**
   - Runtime calls `refreshOAuthToken(refresh_token, provider)`
   - OAuth handler POST to token endpoint:
     ```json
     {
       "grant_type": "refresh_token",
       "refresh_token": "...",
       "client_id": "..."
     }
     ```
   - Provider returns new access token and updated expiry
   - New credentials saved to localStorage
   - Runtime uses new access token for LLM call
4. **If valid:**
   - Runtime uses existing access token
   - No token endpoint request needed

### Phase 5: LLM API Request with Bearer Token

**Headers Sent:**
```
Authorization: Bearer {access_token}
Content-Type: application/json
```

**Response:** Streamed chat/completion response from LLM provider

---

## Data Structures & Storage

### OAuthCredentials Object

```typescript
interface OAuthCredentials {
  refresh: string;    // Long-lived refresh token
  access: string;     // Short-lived access token (typically expires in 1 hour)
  expires: number;    // Expiration timestamp in milliseconds (Date.now() + expires_in*1000)
}
```

**Storage Key:** `{localStoragePrefix}-oauth-credentials`

**Structure in localStorage:**
```json
{
  "openexcel-oauth-credentials": {
    "anthropic": {
      "refresh": "long-lived-refresh-token",
      "access": "short-lived-access-token",
      "expires": 1713619200000
    },
    "openai-codex": {
      "refresh": "...",
      "access": "...",
      "expires": 1713619200000
    }
  }
}
```

### ProviderConfig Object

```typescript
interface ProviderConfig {
  provider: string;        // "anthropic" or "openai-codex"
  apiKey: string;         // Holds access token if authMethod="oauth"
  model: string;          // LLM model ID (e.g., "claude-3-5-sonnet-20241022")
  authMethod: "apikey" | "oauth";  // Authentication method
  useProxy: boolean;       // Whether to use proxy for token requests
  proxyUrl: string;        // CORS proxy URL (if using proxy)
  thinking: "none" | "low" | "medium" | "high";
  followMode: boolean;
  expandToolCalls: boolean;
  apiType?: string;        // Optional API type override
  customBaseUrl?: string;  // Optional custom base URL
}
```

**Storage Key:** `{localStoragePrefix}-provider-config`

### OAuthFlowState

```typescript
type OAuthFlowState =
  | { step: "idle" }                              // No auth attempted
  | { step: "awaiting-code"; verifier: string; oauthState?: string }  // Waiting for auth code
  | { step: "exchanging" }                        // Exchanging code for tokens
  | { step: "connected" }                         // Successfully authenticated
  | { step: "error"; message: string };           // Auth failed
```

---

## Code Implementation Reference

### 1. PKCE Generation

**File:** `packages/sdk/src/oauth/index.ts`

```typescript
export async function generatePKCE(): Promise<{
  verifier: string;
  challenge: string;
}> {
  // Generate cryptographically secure random bytes
  const verifierBytes = new Uint8Array(32);
  crypto.getRandomValues(verifierBytes);
  const verifier = base64urlEncode(verifierBytes);
  
  // Create SHA-256 hash of verifier
  const data = new TextEncoder().encode(verifier);
  const hashBuffer = await crypto.subtle.digest("SHA-256", data);
  const challenge = base64urlEncode(new Uint8Array(hashBuffer));
  
  return { verifier, challenge };
}

function base64urlEncode(bytes: Uint8Array): string {
  let binary = "";
  for (const byte of bytes) {
    binary += String.fromCharCode(byte);
  }
  // Replace URL-unsafe characters
  return btoa(binary)
    .replace(/\+/g, "-")
    .replace(/\//g, "_")
    .replace(/=/g, "");
}
```

### 2. Authorization URL Construction

**File:** `packages/sdk/src/oauth/index.ts`

```typescript
export function buildAuthorizationUrl(
  provider: string,
  challenge: string,
  verifier: string,
): { url: string; oauthState?: string } {
  if (provider === "openai-codex") {
    const oauthState = createRandomState();
    const params = new URLSearchParams({
      response_type: "code",
      client_id: OPENAI_CODEX_CLIENT_ID,
      redirect_uri: OPENAI_CODEX_REDIRECT_URI,
      scope: OPENAI_CODEX_SCOPE,
      code_challenge: challenge,
      code_challenge_method: "S256",
      state: oauthState,
      id_token_add_organizations: "true",
      codex_cli_simplified_flow: "true",
      originator: "pi",
    });
    return { url: `${OPENAI_CODEX_AUTHORIZE_URL}?${params}`, oauthState };
  }
  
  // Anthropic flow
  const params = new URLSearchParams({
    code: "true",
    client_id: ANTHROPIC_CLIENT_ID,
    response_type: "code",
    redirect_uri: ANTHROPIC_REDIRECT_URI,
    scope: ANTHROPIC_SCOPES,
    code_challenge: challenge,
    code_challenge_method: "S256",
    state: verifier,
  });
  return { url: `${ANTHROPIC_AUTHORIZE_URL}?${params}` };
}
```

### 3. Token Refresh Implementation

**File:** `packages/sdk/src/oauth/index.ts`

```typescript
async function refreshAnthropicOAuth(
  refreshToken: string,
  proxyUrl: string,
  useProxy: boolean,
): Promise<OAuthCredentials> {
  const url = buildProxiedUrl(ANTHROPIC_TOKEN_URL, useProxy, proxyUrl);
  const response = await fetch(url, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      grant_type: "refresh_token",
      client_id: ANTHROPIC_CLIENT_ID,
      refresh_token: refreshToken,
    }),
  });
  
  if (!response.ok) {
    throw new Error(`Anthropic token refresh failed: ${response.status}`);
  }
  
  const data = (await response.json()) as {
    access_token: string;
    refresh_token: string;
    expires_in: number;
  };
  
  return {
    refresh: data.refresh_token,
    access: data.access_token,
    expires: Date.now() + data.expires_in * 1000,
  };
}
```

### 4. Runtime Token Check Before API Call

**File:** `packages/sdk/src/runtime.ts`

```typescript
private async getActiveApiKey(config: ProviderConfig): Promise<string> {
  // Non-OAuth auth method: return API key as-is
  if (config.authMethod !== "oauth") {
    return config.apiKey;
  }
  
  // Load OAuth credentials from storage
  const creds = loadOAuthCredentials(this.ns, config.provider);
  if (!creds) return config.apiKey;  // Fallback to API key if no OAuth creds
  
  // Check if token is still valid
  if (Date.now() < creds.expires) {
    return creds.access;  // Token is valid, use it
  }
  
  // Token expired: refresh it
  const refreshed = await refreshOAuthToken(
    config.provider,
    creds.refresh,
    config.proxyUrl,
    config.useProxy,
  );
  
  // Save updated credentials
  saveOAuthCredentials(this.ns, config.provider, refreshed);
  
  // Return new access token
  return refreshed.access;
}
```

### 5. Provider Configuration Loading

**File:** `packages/sdk/src/provider-config.ts`

```typescript
export function loadSavedConfig(ns: StorageNamespace): ProviderConfig | null {
  try {
    const saved = localStorage.getItem(storageKey(ns));
    if (saved) {
      const config = JSON.parse(saved);
      
      // Set defaults for optional fields
      if (config.proxyUrl === undefined) config.proxyUrl = "";
      if (config.followMode === undefined) config.followMode = true;
      if (config.authMethod === undefined) config.authMethod = "apikey";
      
      // For OAuth auth method: populate apiKey with access token
      if (config.authMethod === "oauth") {
        const creds = loadOAuthCredentials(ns, config.provider);
        if (creds) config.apiKey = creds.access;
      }
      
      return config;
    }
  } catch {}
  return null;
}
```

---

## Provider-Specific Configurations

### Anthropic Claude OAuth

**Authorization Endpoint:** `https://claude.ai/oauth/authorize`  
**Token Endpoint:** `https://console.anthropic.com/v1/oauth/token`  
**Redirect URI:** `https://console.anthropic.com/oauth/code/callback`  

**Required Scopes:**
```
org:create_api_key user:profile user:inference
```

**PKCE Support:** ✅ Yes (code_challenge_method=S256)

**Token Lifetime:**
- Access Token: ~1 hour
- Refresh Token: Long-lived (does not expire unless revoked)

### OpenAI ChatGPT OAuth

**Authorization Endpoint:** `https://auth.openai.com/oauth/authorize`  
**Token Endpoint:** `https://auth.openai.com/oauth/token`  
**Redirect URI:** `http://localhost:1455/auth/callback` (local dev or custom handler)

**Required Scopes:**
```
openid profile email offline_access
```

**PKCE Support:** ✅ Yes (code_challenge_method=S256)

**Token Lifetime:**
- Access Token: ~1 hour
- Refresh Token: Long-lived

**Special Parameters:**
- `id_token_add_organizations=true`: Include organization list
- `codex_cli_simplified_flow=true`: Streamlined OAuth flow
- `originator=pi`: Identifies source as PI (Prompt Inspector)

---

## Excel Add-in Integration

### Storage Namespace Configuration

**File:** `packages/excel/src/lib/adapter.ts`

```typescript
const STORAGE_NAMESPACE = {
  dbName: "OpenExcelDB_v3",
  dbVersion: 30,
  localStoragePrefix: "openexcel",
  documentSettingsPrefix: "openexcel",
  documentIdSettingsKey: "openexcel-workbook-id",
};
```

### Session Scoping

- Each workbook has a unique ID stored in Office.js document settings
- Chat sessions, VFS files, and OAuth credentials are scoped to workbook ID in IndexedDB
- Multiple workbooks can have different OAuth sessions
- Logout clears credentials for current workbook only

### Taskpane Authentication UI

**File:** `packages/core/src/chat/settings-panel.svelte`

**UI States:**
1. **idle** — No authentication attempted; show provider selection + login button
2. **awaiting-code** — Waiting for user to paste/receive authorization code
3. **exchanging** — Requesting tokens from provider (loading state)
4. **connected** — Successfully authenticated; show provider name + logout button
5. **error** — Authentication failed; show error message + retry option

**UI Features:**
- Dropdown to select auth method: "API Key" or "OAuth"
- Conditional rendering based on selected provider
- Login button triggers OAuth flow
- Manual code paste field (fallback for redirect issues)
- Logout button to clear credentials

---

## Environment Variables & Configuration

### Required Environment Variables

Add to `.env` file (Vite automatically loads `VITE_*` prefixed variables):

```bash
# Azure AD for future NAA/SSO authentication (not yet implemented)
VITE_AZURE_AD_CLIENT_ID=your-tenant-specific-client-id
VITE_AZURE_AD_TENANT_ID=your-tenant-id

# Datadog monitoring (optional but recommended for production)
VITE_DD_CLIENT_TOKEN=your-datadog-client-token
VITE_DD_APP_ID=your-datadog-app-id
VITE_DD_SITE=datadoghq.eu
VITE_DD_SERVICE=excel-mate
VITE_DD_ENV=production
VITE_DD_VERSION=0.1.1
```

### Embedded Configuration (in SDK)

The following OAuth client IDs and endpoints are embedded in the source code (base64-encoded for obfuscation):

| Provider | Client ID | Token URL | Redirect URI |
|----------|-----------|-----------|--------------|
| Anthropic | Embedded (base64) | `https://console.anthropic.com/v1/oauth/token` | `https://console.anthropic.com/oauth/code/callback` |
| OpenAI Codex | Embedded (base64) | `https://auth.openai.com/oauth/token` | `http://localhost:1455/auth/callback` |

**Note:** These credentials are public client IDs used in browser contexts; they are intentionally not sensitive secrets.

---

## Security Considerations

### 1. PKCE Protection

**What it does:** Prevents authorization code interception attacks in browser-based OAuth flows

**How it works:**
- `code_verifier`: Random 43-128 byte string generated locally
- `code_challenge`: SHA-256 hash of verifier
- Only the challenge is sent to authorization endpoint
- Verifier is sent directly to token endpoint (not through browser redirect)
- Provider validates that `SHA-256(verifier) == challenge`

**Why it matters:** If malicious code intercepts the authorization code, it cannot obtain tokens without the verifier.

### 2. State Parameter

**What it does:** CSRF (Cross-Site Request Forgery) protection

**How it works:**
- Random `state` token generated before redirect
- Stored in localStorage
- Provider includes state in redirect back
- Application validates returned state matches stored state

### 3. Token Storage

**Current:** Browser localStorage (not ideal, but standard for SPAs)

**Security:** 
- Tokens stored in localStorage are accessible to any JavaScript on the page
- Protected by browser Same-Origin Policy (SOP)
- Encrypted by browser (TLS in transit)
- **Recommendation:** Use sessionStorage for short-lived access tokens only; store refresh tokens server-side

### 4. Refresh Token Rotation

**Current:** Refresh tokens stored indefinitely in localStorage

**Security Considerations:**
- Long-lived tokens are valuable to attackers
- Providers may implement refresh token rotation (new refresh token per refresh)
- Implement token cleanup/revocation on logout

### 5. Bearer Token Usage

**Pattern:**
```
Authorization: Bearer {access_token}
```

**Security:**
- Tokens transmitted in Authorization header (not URL parameter)
- TLS/HTTPS required (adding `integrity` and `confidentiality`)
- Access tokens are short-lived (~1 hour)

### 6. Optional: CORS Proxy for Token Requests

**File:** `packages/sdk/src/oauth/index.ts`

**Use Case:** Some environments block direct cross-origin requests to token endpoints

**Implementation:**
```typescript
function buildProxiedUrl(
  baseUrl: string,
  useProxy: boolean,
  proxyUrl: string,
): string {
  return useProxy && proxyUrl
    ? `${proxyUrl}/?url=${encodeURIComponent(baseUrl)}`
    : baseUrl;
}
```

**Risk:** Routing tokens through a proxy introduces a man-in-the-middle risk; only use trusted proxies or server-owned proxies.

---

## Known Limitations & Future Enhancements

### Current Limitations

1. **No Backend Token Relay**
   - All token handling occurs in the browser
   - No server-side token storage or validation
   - Tokens visible in browser developer tools

2. **No Refresh Token Rotation**
   - Refresh tokens stored indefinitely
   - No automatic refresh token cleanup

3. **No Multi-Account Support**
   - Only one account per provider at a time
   - Switching accounts requires logout + re-login

4. **No Token Revocation Endpoint**
   - Logout clears local storage but does not revoke tokens at provider
   - Tokens remain valid until natural expiry

### Planned Enhancements

1. **Azure AD / NAA Integration** (in progress)
   - MSAL.js for Microsoft Entra ID authentication
   - Single Sign-On (SSO) for enterprise customers
   - **Status:** Dependencies installed; implementation to follow

2. **Secure Token Storage**
   - Server-side token relay for sensitive applications
   - HttpOnly cookies for tokens (if backend added)
   - Token encryption at rest

3. **Token Refresh Rotation**
   - Automatic rotation of refresh tokens
   - Cleanup of expired tokens

4. **Logout with Token Revocation**
   - Call provider API to explicitly revoke tokens
   - Remove from storage

---

## Testing & Validation Checklist

### Manual Testing

- [ ] Click "Login with Provider" button
- [ ] Redirect to OAuth provider succeeds
- [ ] User can authenticate with provider credentials
- [ ] Redirect back to add-in returns authorization code
- [ ] Token exchange succeeds; "Connected" state displayed
- [ ] Provider name and model appear in settings
- [ ] Send chat message; receives LLM response using OAuth token
- [ ] Wait until token expiry; send another message; automatic refresh occurs
- [ ] Click logout; credentials removed from localStorage
- [ ] Settings panel shows "Not connected" after logout

### Automated Tests

See: `packages/sdk/tests/` for runtime and provider-config tests

---

## Troubleshooting

### Issue: "OAuth flow failed" / "Unable to exchange code"

**Possible Causes:**
1. Network firewall blocking token endpoint
2. CORS proxy not functioning
3. Provider API temporarily down
4. Wrong redirect URI configured

**Resolution:**
- Check browser console for fetch error details
- Verify redirect URI matches provider's registered URI
- Try disabling proxy if enabled
- Contact provider support if endpoint is down

### Issue: Token refresh fails; "Refresh token invalid"

**Possible Causes:**
1. Refresh token expired (use case-specific; some providers invalidate tokens after 90 days)
2. User revoked token at provider
3. Provider requires re-authentication

**Resolution:**
- User must re-authenticate (logout + login)
- Check provider's token expiration policy

### Issue: Multiple browsers/tabs; tokens out of sync

**Current Limitation:** localStorage is per-origin, not per-tab

**Workaround:** Close and reopen tabs; refresh with F5

---

## Additional Resources

- **OAuth 2.0 Spec:** https://tools.ietf.org/html/rfc6749
- **PKCE Extension:** https://tools.ietf.org/html/rfc7636
- **OpenID Connect:** https://openid.net/connect/
- **Anthropic OAuth Documentation:** https://docs.anthropic.com/claude/reference/verify-account
- **OpenAI OAuth Documentation:** https://platform.openai.com/docs/guides/oauth

---

## Appendix: Complete Provider Endpoints

### Anthropic

```
Authorization: https://claude.ai/oauth/authorize
Token: https://console.anthropic.com/v1/oauth/token
Redirect: https://console.anthropic.com/oauth/code/callback
Scopes: org:create_api_key user:profile user:inference
```

### OpenAI Codex

```
Authorization: https://auth.openai.com/oauth/authorize
Token: https://auth.openai.com/oauth/token
Redirect: http://localhost:1455/auth/callback
Scopes: openid profile email offline_access
```

