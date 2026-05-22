# Azure AD / Entra ID Configuration Checklist & Implementation Guide

**Document Purpose:** This guide provides step-by-step instructions for configuring Azure Active Directory (Entra ID) OAuth for a Static Web App (SWA) to replicate the OpenExcel authentication pattern.

**Pre-Requisites:**
- Azure subscription with Entra ID tenant
- Static Web App created in Azure
- Permissions to manage App Registrations in Entra ID

---

## Part 1: App Registration Setup in Entra ID

### Step 1: Create New App Registration

1. Navigate to **Azure Portal** → **Entra ID** → **App registrations**
2. Click **New registration**
3. Fill in details:
   - **Name:** `OpenExcel SWA` (or your application name)
   - **Supported account types:** Choose based on your needs:
     - ✅ **Accounts in this organizational directory only** — Single-tenant (recommended for enterprise)
     - ⭕ **Accounts in any organizational directory** — Multi-tenant (for broader access)
   - **Redirect URI:**
     - Platform: **Single-page application (SPA)**
     - URI: `https://{your-swa-domain}.azurestaticapps.net/auth/callback`
     - For local dev: `http://localhost:3000/auth/callback`
4. Click **Register**

### Step 2: Note Client Credentials

On the **Overview** page, copy and save:
- **Application (client) ID** — GUID used in OAuth requests
- **Directory (tenant) ID** — Your tenant GUID

These will be used in environment variables:
```
VITE_AZURE_AD_CLIENT_ID={Application (client) ID}
VITE_AZURE_AD_TENANT_ID={Directory (tenant) ID}
```

### Step 3: Configure Redirect URIs

1. Navigate to **Authentication** in the left sidebar
2. Under **Redirect URIs**, add:
   - `https://{your-swa-domain}.azurestaticapps.net/auth/callback`
   - `https://{your-swa-domain}.azurestaticapps.net` (for non-callback redirects)
   - For development: `http://localhost:3000` and `http://localhost:3000/auth/callback`
3. Ensure "Treat `https://{your-custom-domain}/auth/callback` as a public client redirect URI" is **CHECKED**
4. Click **Save**

---

## Part 2: API Permissions & Scopes

### Step 4: Add API Permissions

1. Navigate to **API permissions** in the left sidebar
2. Click **Add a permission**
3. Select **Microsoft Graph**
4. Choose **Delegated permissions** (user-context access)
5. Search for and select:
   - ✅ `User.Read` — Read basic user profile (required)
   - ✅ `profile` — Include profile in token claims
   - ✅ `email` — Include email in token claims
   - ⭕ `offline_access` — Enable refresh token (recommended)
   - ⭕ `Directory.Read.All` — Read directory/groups (if using group-based authorization)

6. Click **Add permissions**

### Step 5: Grant Admin Consent (Optional but Recommended)

1. After adding permissions, click **Grant admin consent for {tenant}**
2. Click **Yes** to confirm
3. Status should show green checkmarks under "Granted"

---

## Part 3: Token Configuration

### Step 6: Configure Token Claims

1. Navigate to **Token configuration** in the left sidebar
2. Click **Add optional claim**
3. Select token type: **ID** (for identity) and/or **Access** (for API calls)
4. Select claims to include:
   - ✅ `email` — User's email address
   - ✅ `given_name` — First name
   - ✅ `family_name` — Last name
   - ✅ `upn` — User Principal Name
   - ⭕ `groups` — Security groups (for role-based authorization)
   - ⭕ `roles` — App Roles (if using RBAC)

5. Click **Add**
6. Repeat for "Access" token if needed

### Step 7: Configure Issued Token Lifetime (Optional)

1. Stay in **Token configuration** → **Protocol settings**
2. Set token lifetimes (defaults are usually fine):
   - **ID token lifetime:** 10 min (default)
   - **Access token lifetime:** 1 hour (default)
   - **Refresh token lifetime:** 14 days (default)
3. Click **Save**

---

## Part 4: Client Secret (For Backend Integration)

### Step 8: Create Client Secret (If Backend OAuth Relay Needed)

**Only required if implementing a backend token relay service** (not needed for pure client-side BYOK model)

1. Navigate to **Certificates & secrets** in the left sidebar
2. Click **New client secret**
3. Fill in:
   - **Description:** `SWA OAuth Backend Secret`
   - **Expires:** Select timeframe (1 year recommended)
4. Click **Add**
5. **IMMEDIATELY** copy and save the secret value
   - ⚠️ This value is only shown once
   - Store securely (e.g., Azure Key Vault, GitHub Secrets)

**Never commit secrets to version control.**

---

## Part 5: Application Roles (RBAC Setup)

### Step 9: Define App Roles (Optional, for Authorization)

1. Navigate to **App roles** in the left sidebar
2. Click **Create app role**
3. Define roles for your application:

**Role 1: User**
- **Display name:** `OpenExcel User`
- **Allowed member types:** Users/Groups
- **Value:** `Excel.User`
- **Description:** `Basic user access to OpenExcel`
- Click **Create**

**Role 2: Admin** (optional)
- **Display name:** `OpenExcel Admin`
- **Allowed member types:** Users/Groups
- **Value:** `Excel.Admin`
- **Description:** `Administrative access to OpenExcel`
- Click **Create**

### Step 10: Assign Users to App Roles

1. Navigate to **Manage** → **Users and groups** in the left sidebar
2. Click **Add user/group**
3. Select users or groups
4. Assign roles
5. Click **Assign**

---

## Part 6: Protected API Configuration (If Using Azure APIM)

### Step 11: Create API in Azure API Management (Optional)

**Only if adding an Azure API Management gateway for token validation**

1. Navigate to **Azure API Management** service
2. Go to **APIs** → **Add API**
3. Define API:
   - **Name:** `OpenExcel API`
   - **Display name:** `OpenExcel API`
   - **Base URL:** `https://{your-backend-function-app}.azurewebsites.net`

### Step 12: Add TOKEN VALIDATION Policy to APIM (Optional)

**In API Management → Policies → Design → Add policy**

```xml
<policies>
  <inbound>
    <validate-jwt header-name="Authorization" failed-validation-httpcode="401" failed-validation-error-message="Unauthorized. Access token is missing or invalid.">
      <openid-config url="https://login.microsoftonline.com/{tenant-id}/.well-known/openid-configuration" />
      <audiences>
        <audience>{application-client-id}</audience>
      </audiences>
      <issuers>
        <issuer>https://sts.windows.net/{tenant-id}/</issuer>
      </issuers>
      <required-claims>
        <claim name="aud">
          <value>{application-client-id}</value>
        </claim>
      </required-claims>
    </validate-jwt>
  </inbound>
  <backend>
    <forward-request />
  </backend>
  <outbound>
    <base />
  </outbound>
  <on-error>
    <base />
  </on-error>
</policies>
```

---

## Part 7: Environment Variables Setup

### Step 13: Create `.env` File with OAuth Configuration

**File:** `.env` (root of SWA project)

```bash
# ============================================
# Azure AD / Entra ID OAuth Configuration
# ============================================

# Azure AD Client ID (Application ID from App Registration)
VITE_AZURE_AD_CLIENT_ID=12345678-1234-1234-1234-123456789012

# Azure AD Tenant ID (Directory ID from App Registration)
VITE_AZURE_AD_TENANT_ID=87654321-4321-4321-4321-210987654321

# OAuth Authority (endpoint for obtaining tokens)
# Format: https://login.microsoftonline.com/{tenant-id}/v2.0
VITE_AZURE_AD_AUTHORITY=https://login.microsoftonline.com/87654321-4321-4321-4321-210987654321/v2.0

# OAuth Redirect URI (must match App Registration configuration)
VITE_AZURE_AD_REDIRECT_URI=https://your-swa-domain.azurestaticapps.net/auth/callback

# OAuth Scopes (space-separated)
# Scopes available: User.Read, profile, email, offline_access, Directory.Read.All
VITE_AZURE_AD_SCOPES=User.Read profile email offline_access

# ============================================
# Application Configuration
# ============================================

# Front-End Base URL
VITE_APP_URL=https://your-swa-domain.azurestaticapps.net

# Back-End API Base URL (if applicable)
VITE_API_URL=https://your-backend-api.azurewebsites.net

# ============================================
# Feature Flags (Optional)
# ============================================

# Enable/disable OAuth SSO
VITE_ENABLE_OAUTH_SSO=true

# Enable/disable BYOK (Bring Your Own Key) mode
VITE_ENABLE_BYOK=true

# ============================================
# Monitoring (Optional)
# ============================================

# Datadog monitoring configuration
VITE_DD_CLIENT_TOKEN=pub123456789abcdef
VITE_DD_APP_ID=12345678-1234-1234-1234-123456789012
VITE_DD_SITE=datadoghq.eu
VITE_DD_ENV=production
```

### Step 14: Development (`localhost`) vs. Production Configuration

**Development (`localhost:3000`):**
```bash
VITE_AZURE_AD_CLIENT_ID=dev-client-id-from-app-registration
VITE_AZURE_AD_REDIRECT_URI=http://localhost:3000/auth/callback
VITE_APP_URL=http://localhost:3000
```

**Production (SWA Domain):**
```bash
VITE_AZURE_AD_CLIENT_ID=prod-client-id-from-app-registration
VITE_AZURE_AD_REDIRECT_URI=https://your-swa-domain.azurestaticapps.net/auth/callback
VITE_APP_URL=https://your-swa-domain.azurestaticapps.net
```

---

## Part 8: MSAL.js Integration in SWA

### Step 15: Install MSAL.js Dependencies

```bash
npm install @azure/msal-browser @azure/msal-react
# or with pnpm
pnpm add @azure/msal-browser @azure/msal-react
```

### Step 16: Initialize MSAL in Application

**File:** `src/main.tsx` or `src/main.ts`

```typescript
import { PublicClientApplication } from "@azure/msal-browser";

const msalConfig = {
  auth: {
    clientId: import.meta.env.VITE_AZURE_AD_CLIENT_ID,
    authority: import.meta.env.VITE_AZURE_AD_AUTHORITY,
    redirectUri: import.meta.env.VITE_AZURE_AD_REDIRECT_URI,
  },
  cache: {
    cacheLocation: "localStorage", // or "sessionStorage"
    storeAuthStateInCookie: false,
  },
  system: {
    allowRedirectInIframe: false,
    loggerOptions: {
      loggerCallback: (logLevel, message, containsPii) => {
        console.log("[MSAL]", message);
      },
      piiLoggingEnabled: false,
    },
  },
};

const msalInstance = new PublicClientApplication(msalConfig);

// Handle redirect after OAuth callback
msalInstance
  .handleRedirectPromise()
  .then((tokenResponse) => {
    if (tokenResponse) {
      console.log("Login successful", tokenResponse);
    }
  })
  .catch((error) => {
    console.error("MSAL error:", error);
  });
```

### Step 17: Wrap Application with MsalProvider

**File:** `src/App.tsx` or root Svelte component

```typescript
import { MsalProvider } from "@azure/msal-react";
import { ChatInterface } from "@office-agents/core";

function App() {
  return (
    <MsalProvider instance={msalInstance}>
      <ChatInterface adapter={excelAdapter} />
    </MsalProvider>
  );
}
```

### Step 18: Implement Login Handler

**File:** `src/components/auth.ts` or similar

```typescript
import { useMsal } from "@azure/msal-react";

export function useAuthHandler() {
  const { instance, accounts } = useMsal();

  const login = async () => {
    try {
      const response = await instance.loginPopup({
        scopes: import.meta.env.VITE_AZURE_AD_SCOPES?.split(" "),
        prompt: "select_account",
      });
      console.log("Login successful", response);
      return response;
    } catch (error) {
      console.error("Login failed", error);
      throw error;
    }
  };

  const logout = async () => {
    try {
      await instance.logoutPopup();
      console.log("Logout successful");
    } catch (error) {
      console.error("Logout failed", error);
    }
  };

  const getAccessToken = async () => {
    if (accounts.length === 0) return null;

    try {
      const response = await instance.acquireTokenSilent({
        scopes: import.meta.env.VITE_AZURE_AD_SCOPES?.split(" "),
        account: accounts[0],
      });
      return response.accessToken;
    } catch (error) {
      console.error("Token acquisition failed", error);
      // Fallback to interactive token acquisition
      const response = await instance.acquireTokenPopup({
        scopes: import.meta.env.VITE_AZURE_AD_SCOPES?.split(" "),
      });
      return response.accessToken;
    }
  };

  return { login, logout, getAccessToken, isAuthenticated: accounts.length > 0 };
}
```

---

## Part 9: Backend API Token Validation (Optional)

### Step 19: Azure Function with Token Validation

**File:** `api/validate-token/index.ts` (Azure Function)

```typescript
import { Context, HttpRequest } from "@azure/functions";
import { jwtDecode } from "jwt-decode";
import fetch from "node-fetch";

export async function validateOAuthToken(
  context: Context,
  req: HttpRequest,
): Promise<void> {
  const authHeader = req.headers.get("Authorization");

  if (!authHeader || !authHeader.startsWith("Bearer ")) {
    context.res = {
      status: 401,
      body: { error: "Missing or invalid Authorization header" },
    };
    return;
  }

  const token = authHeader.substring(7);

  try {
    // Fetch OIDC metadata
    const tenantId = process.env.AZURE_AD_TENANT_ID;
    const response = await fetch(
      `https://login.microsoftonline.com/${tenantId}/.well-known/openid-configuration`,
    );
    const oidcConfig = (await response.json()) as {
      jwks_uri: string;
    };

    // Fetch public keys
    const keysResponse = await fetch(oidcConfig.jwks_uri);
    const keys = await keysResponse.json();

    // Decode token header to get key ID (kid)
    const decoded = jwtDecode(token, { header: true });
    const kid = (decoded as Record<string, any>).kid;

    // Find matching public key
    const key = keys.keys.find(
      (k: Record<string, any>) => k.kid === kid,
    );
    if (!key) {
      throw new Error("Public key not found");
    }

    // Validate token signature and claims
    // (Use a JWT library like jsonwebtoken to handle verification properly)
    context.res = {
      status: 200,
      body: {
        valid: true,
        claims: decoded,
      },
    };
  } catch (error) {
    context.res = {
      status: 401,
      body: { error: "Token validation failed", message: (error as Error).message },
    };
  }
}
```

---

## Part 10: Static Web App Configuration

### Step 20: Configure SWA `staticwebapp.config.json`

**File:** `staticwebapp.config.json` (root of SWA project)

```json
{
  "auth": {
    "identityProviders": {
      "azureActiveDirectory": {
        "registration": {
          "openIdIssuer": "https://login.microsoftonline.com/{tenant-id}/v2.0",
          "clientIdSettingName": "AZURE_CLIENT_ID",
          "clientSecretSettingName": "AZURE_CLIENT_SECRET"
        },
        "login": {
          "loginParameters": ["scope=User.Read profile email offline_access"]
        }
      }
    }
  },
  "routes": [
    {
      "route": "/auth/*",
      "allowedRoles": ["authenticated"]
    },
    {
      "route": "/api/*",
      "allowedRoles": ["authenticated"]
    },
    {
      "route": "/*",
      "serve": "index.html",
      "statusCode": 200
    }
  ]
}
```

---

## Part 11: Azure Static Web App Environment Variables

### Step 21: Set Environment Variables in SWA

1. Navigate to **Azure Portal** → **Static Web Apps** → Your SWA instance
2. Go to **Settings** → **Configuration**
3. Add environment variables:

```
VITE_AZURE_AD_CLIENT_ID = {copy from Step 2}
VITE_AZURE_AD_TENANT_ID = {copy from Step 2}
VITE_AZURE_AD_AUTHORITY = https://login.microsoftonline.com/{tenant-id}/v2.0
VITE_AZURE_AD_REDIRECT_URI = https://{your-swa-domain}.azurestaticapps.net/auth/callback
VITE_AZURE_AD_SCOPES = User.Read profile email offline_access
VITE_APP_URL = https://{your-swa-domain}.azurestaticapps.net
```

---

## Configuration Validation Checklist

Use this checklist to verify all OAuth configurations are in place:

### App Registration

- [ ] App registration created in Entra ID
- [ ] Client ID copied and saved
- [ ] Tenant ID copied and saved
- [ ] Redirect URIs configured (production + localhost dev)
- [ ] SPA platform selected for redirect URI type

### API Permissions

- [ ] `User.Read` permission added
- [ ] `profile` and `email` scopes added
- [ ] `offline_access` scope added (for refresh tokens)
- [ ] Admin consent granted (green checkmarks visible)

### Token Configuration

- [ ] Optional claims configured (email, groups, roles)
- [ ] Token lifetimes set appropriately
- [ ] ID token and Access token lifetimes configured

### MSAL.js Setup

- [ ] MSAL.js dependencies installed
- [ ] MSAL config object created with correct client ID and authority
- [ ] MsalProvider wraps application
- [ ] handleRedirectPromise() called on app startup
- [ ] Login handler implemented with correct scopes
- [ ] Token acquisition implemented (silent + interactive fallback)

### Environment Variables

- [ ] `.env` file created with all OAuth variables
- [ ] `VITE_AZURE_AD_CLIENT_ID` set
- [ ] `VITE_AZURE_AD_TENANT_ID` set
- [ ] `VITE_AZURE_AD_AUTHORITY` set
- [ ] `VITE_AZURE_AD_REDIRECT_URI` set
- [ ] `VITE_AZURE_AD_SCOPES` set
- [ ] SWA environment variables configured in Azure Portal

### SWA Configuration

- [ ] `staticwebapp.config.json` configured
- [ ] Routes restricted to authenticated users where appropriate
- [ ] SWA configured in Azure Portal

### Testing

- [ ] Login flow works locally (`localhost:3000`)
- [ ] OAuth redirect successful
- [ ] Tokens received and stored
- [ ] Login flow works in production SWA domain
- [ ] Logout clears tokens
- [ ] Token refresh works automatically
- [ ] API calls include Authorization header with bearer token

---

## Troubleshooting

### "AADSTS50058: Silent sign-in request failed"

**Cause:** Browser blocked third-party cookies or user not logged into Microsoft account

**Solution:**
- Ensure browser allows third-party cookies for the identity provider
- User must log into Microsoft account separately or in pop-up window
- Use `loginPopup()` instead of `loginRedirect()` for more reliable flow

### "AADSTS650052: The app needs access to a service that your administrator has not yet authorized"

**Cause:** Required API permissions not granted or admin consent not provided

**Solution:**
- Go to App Registration → **API permissions**
- Add missing scopes
- Grant admin consent (green checkmark should appear)

### "Redirect URI mismatch"

**Cause:** Redirect URI in code doesn't match App Registration configuration

**Solution:**
- Navigate to **App Registration** → **Authentication**
- Verify redirect URIs listed match those in MSAL config and environment variables
- Ensure exact match (case-sensitive, including trailing slash)

### Token is undefined after login

**Cause:** Token not being acquired or stored

**Solution:**
- Check browser console for MSAL errors
- Verify scopes include `offline_access` for refresh tokens
- Ensure `acquireTokenSilent()` is called after login
- In incognito window, must use `acquireTokenPopup()` due to cookie restrictions

---

## Additional Resources

- [Microsoft Entra ID Documentation](https://learn.microsoft.com/en-us/entra/identity/)
- [MSAL.js Documentation](https://github.com/AzureAD/microsoft-authentication-library-for-js)
- [OAuth 2.0 Protocol](https://tools.ietf.org/html/rfc6749)
- [Azure Static Web Apps Auth](https://learn.microsoft.com/en-us/azure/static-web-apps/authentication-authorization)
- [Azure API Management JWT Validation](https://learn.microsoft.com/en-us/azure/api-management/api-management-access-restriction-policies)

