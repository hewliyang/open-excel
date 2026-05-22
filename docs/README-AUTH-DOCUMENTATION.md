# OpenExcel Authentication Flow - Complete Documentation Package

**Date:** April 20, 2026  
**Status:** Complete & Ready for Implementation  
**Version:** 1.0

---

## 📋 Documentation Package Overview

This package contains a comprehensive blueprint of the OpenExcel authentication flow, designed to facilitate replication of the proven auth pattern in a new Static Web App (SWA). The documentation is organized into four primary deliverables plus this summary document.

### Included Deliverables

1. **AUTH-FLOW-DOCUMENTATION.md** (Core Technical Guide)
   - End-to-end authentication flow explanation
   - PKCE OAuth 2.0 protocol details
   - Implementation code snippets
   - Security considerations
   - Troubleshooting guide

2. **ENTRA-ID-CONFIG-CHECKLIST.md** (Step-by-Step Setup)
   - Azure AD App Registration creation
   - API permissions configuration
   - Token claims setup
   - MSAL.js integration guide
   - SWA configuration
   - Complete 21-step implementation checklist

3. **ENVIRONMENT-VARIABLES-REFERENCE.md** (Configuration Guide)
   - Detailed environment variable reference
   - Dev/Staging/Production `.env` templates
   - Token storage schema reference
   - 10 production-ready code snippets
   - Debugging tips and quick start commands

4. **Architecture & Sequence Diagrams** (Visual Reference)
   - High-level component architecture diagram
   - Detailed OAuth 2.0 PKCE flow sequence diagram
   - All major components and interactions illustrated

---

## 🎯 Key Findings Summary

### Current Implementation (OpenExcel)

**Authentication Method:** OAuth 2.0 with PKCE  
**Supported Providers:** Anthropic Claude, OpenAI ChatGPT  
**User Model:** BYOK (Bring Your Own Key)  
**Architecture:** Client-side only (browser-based)  
**Token Storage:** Browser localStorage + IndexedDB  
**Token Refresh:** Proactive (before each API call)

### Critical Implementation Details Found

1. **PKCE Protection**
   - Code verifier: 32-byte random string (base64url encoded)
   - Code challenge: SHA-256 hash of verifier (S256 method)
   - Prevents authorization code interception

2. **Token Lifecycle**
   - Access token: ~1 hour lifetime
   - Refresh token: Long-lived (indefinite)
   - Automatic refresh triggered before expired token is used
   - Runtime checks expiry: `Date.now() < credentials.expires`

3. **Storage Architecture**
   - localStorage key: `{prefix}-oauth-credentials`
   - Structure: `{ provider: { refresh, access, expires } }`
   - Scoped per workbook (for Excel add-in)
   - Multiple providers can have simultaneous sessions

4. **UI State Management**
   - Five OAuth flow states tracked in Svelte components
   - States: idle → awaiting-code → exchanging → connected → error
   - User can manually paste auth code (redirect fallback)

5. **Provider Configuration**
   - Client IDs embedded (base64-encoded in source)
   - Endpoints: Anthropic and OpenAI OAuth providers
   - Scopes configured per provider

---

## 🔐 Security Model

### PKCE (Proof Key for Code Exchange)
- **Purpose:** Protect against authorization code interception
- **Implementation:** SHA-256 hash of random verifier
- **Protection:** Verifier never exposed to browser redirect

### State Parameter
- **Purpose:** CSRF (Cross-Site Request Forgery) prevention
- **Implementation:** Random token generated and stored
- **Validation:** Verify returned state matches stored value

### Bearer Token Usage
- **Format:** `Authorization: Bearer {access_token}`
- **Transport:** Authorization header (not URL parameter)
- **Requirement:** TLS/HTTPS enforced

### Token Storage
- **Current:** Browser localStorage (SOP protected)
- **Concern:** Accessible to page JavaScript
- **Future Improvement:** Server-side token relay recommended

---

## 📊 Flow Summary

### Initial Authentication
```
User clicks "Login" 
  → Generate PKCE (verifier + challenge)
  → Redirect to OAuth provider
  → User authenticates
  → Provider returns authorization code
  → Exchange code + PKCE verifier for tokens
  → Store { access, refresh, expires } in localStorage
  → UI shows "Connected"
```

### Token Usage & Refresh
```
User sends chat message
  → Runtime calls getActiveApiKey()
  → Check: is token expired?
  → If expired: refreshOAuthToken() with refresh_token
  → Use access_token in Bearer header
  → Send request to LLM API
  → Stream response back to user
```

### Logout
```
User clicks logout
  → Remove OAuth credentials from localStorage
  → Clear session data
  → UI shows "Not Connected"
```

---

## 🛠️ Implementation Roadmap for New SWA

### Phase 1: Azure AD Setup (2-3 hours)
- [ ] Create Entra ID App Registration
- [ ] Note Client ID and Tenant ID
- [ ] Configure redirect URIs (dev + production)
- [ ] Add API permissions (User.Read, etc.)
- [ ] Grant admin consent
- [ ] Create `.env` file with credentials

**Reference Document:** `ENTRA-ID-CONFIG-CHECKLIST.md` — Follow steps 1-14

### Phase 2: MSAL.js Integration (4-6 hours)
- [ ] Install MSAL.js dependencies
- [ ] Create MSAL config object
- [ ] Initialize MSAL on app startup
- [ ] Implement handleRedirectPromise()
- [ ] Create login/logout handlers
- [ ] Implement token acquisition (silent + interactive)
- [ ] Add authentication state hook/store

**Reference Document:** `ENVIRONMENT-VARIABLES-REFERENCE.md` — See snippets 1-8

### Phase 3: Backend Token Validation (3-4 hours, optional)
- [ ] Create Azure Function for token validation
- [ ] Implement JWT verification with JWKS
- [ ] Add JWT validation policy to APIM (if using gateway)
- [ ] Protect APIs with authentication middleware

**Reference Document:** `AUTH-FLOW-DOCUMENTATION.md` — See "Backend Integration" section

### Phase 4: Testing & Validation (2-3 hours)
- [ ] Test login flow locally
- [ ] Verify token acquisition
- [ ] Test automatic token refresh
- [ ] Test logout and credential cleanup
- [ ] Deploy to staging SWA
- [ ] Run production validation checklist

**Reference Document:** `AUTH-FLOW-DOCUMENTATION.md` — See "Testing & Validation" section

**Total Estimated Time:** 11-16 hours of development

---

## 📁 File Structure for SWA Project

**Recommended directory structure:**

```
src/
├── auth/
│   ├── msal-config.ts         (MSAL setup)
│   ├── auth-handler.ts        (Login/logout/token functions)
│   └── auth-hook.ts          (React hook or Svelte store)
├── components/
│   ├── AuthButton.tsx         (Login/logout button)
│   └── ProtectedRoute.tsx     (Require auth component)
├── api/
│   └── api-client.ts          (Bearer token requests)
├── App.tsx                    (Wrap with MsalProvider)
└── main.tsx                   (Initialize MSAL)

api/
└── validate-token/
    └── index.ts              (Azure Function for token validation)

.env                           (Local development)
.env.staging                   (Staging environment)
.env.production                (Production environment)
staticwebapp.config.json       (SWA routing & auth config)
```

---

## 🔍 Key Files in Original Codebase

For reference, the current OpenExcel implementation uses:

| File | Purpose |
|------|---------|
| `packages/sdk/src/oauth/index.ts` | PKCE generation, token exchange, refresh |
| `packages/sdk/src/provider-config.ts` | OAuth config storage and loading |
| `packages/sdk/src/runtime.ts` | Token validation before API calls |
| `packages/core/src/chat/settings-panel.svelte` | OAuth UI and flow management |
| `packages/excel/src/lib/adapter.ts` | Excel add-in integration |
| `packages/sdk/src/storage/` | IndexedDB and localStorage management |

---

## 🚀 Quick Start for New SWA

### 1. Create App Registration
```bash
# Navigate to: Azure Portal → Entra ID → App Registrations → New registration
# Follow: ENTRA-ID-CONFIG-CHECKLIST.md steps 1-5
```

### 2. Clone/Create SWA Project
```bash
npx create-react-app my-swa-app
# or with Next.js
npx create-next-app@latest my-swa-app
# or with Svelte
npm create vite@latest my-swa-app -- --template svelte
```

### 3. Install Dependencies
```bash
npm install @azure/msal-browser @azure/msal-react jose
```

### 4. Add Configuration
```bash
# Create .env file with values from App Registration
cp .env.template .env
# Edit with your Client ID, Tenant ID, etc.
```

### 5. Integrate MSAL Code
```bash
# Use code snippets from ENVIRONMENT-VARIABLES-REFERENCE.md
# Snippets 1-8 provide complete MSAL setup
```

### 6. Deploy to SWA
```bash
# Push to GitHub
git push origin main

# GitHub Actions will build and deploy to Azure Static Web Apps
```

---

## ⚠️ Important Notes & Limitations

### Current Limitations
1. **No backend token storage** — Tokens visible in browser localStorage
2. **No refresh token rotation** — Tokens stored indefinitely
3. **No multi-account support** — One account per provider at a time
4. **No token revocation** — Logout doesn't revoke at provider

### Security Recommendations
1. Use **HttpOnly cookies** with a backend token relay for production
2. Implement **refresh token rotation** if using long-lived tokens
3. Add **token cleanup** on logout (revoke at provider)
4. Consider **Content Security Policy (CSP)** to protect against XSS

### Future Enhancements Tracked in Codebase
- Azure AD NAA/SSO integration (MSAL as fallback)
- Server-side token relay implementation
- Token storage encryption
- Multi-provider simultaneous sessions
- Built-in token revocation on logout

---

## 📚 External References

### OAuth & Security Standards
- [OAuth 2.0 Specification](https://tools.ietf.org/html/rfc6749)
- [PKCE Extension (RFC 7636)](https://tools.ietf.org/html/rfc7636)
- [OpenID Connect 1.0](https://openid.net/specs/openid-connect-core-1_0.html)

### Azure & MSAL Documentation
- [Microsoft Entra ID Overview](https://learn.microsoft.com/en-us/entra/identity/)
- [MSAL.js GitHub Repository](https://github.com/AzureAD/microsoft-authentication-library-for-js)
- [Azure Static Web Apps Authentication](https://learn.microsoft.com/en-us/azure/static-web-apps/authentication-authorization)

### Provider Documentation
- [Anthropic OAuth](https://docs.anthropic.com/claude/reference/verify-account)
- [OpenAI OAuth](https://platform.openai.com/docs/guides/oauth)

---

## 📞 Troubleshooting Quick Reference

### Common Issues & Solutions

| Issue | Cause | Solution |
|-------|-------|----------|
| AADSTS50058: Silent sign-in failed | Browser blocked cookies | Ensure 3rd-party cookies allowed; use popup login |
| Redirect URI mismatch | Config doesn't match App Reg | Verify exact URI match (case-sensitive) in App Registration |
| Token is undefined | Token not acquired | Check scopes include `offline_access`; verify token acquisition called |
| "No user logged in" after refresh | Token not persisted | Check localStorage cache location set correctly in MSAL config |
| API 401 Unauthorized | Invalid/expired token | Verify Bearer token format; check token expiry and refresh logic |

For detailed troubleshooting, see `AUTH-FLOW-DOCUMENTATION.md` → "Troubleshooting" section.

---

## ✅ Pre-Deployment Checklist

### Before Going to Production

**Configuration:**
- [ ] All environment variables set in Azure Portal
- [ ] Client ID and Tenant ID verified
- [ ] Redirect URI points to production SWA domain
- [ ] staticwebapp.config.json configured properly

**Code Quality:**
- [ ] MSAL logging disabled (piiLoggingEnabled=false)
- [ ] No console errors in production build
- [ ] CSP headers configured if using CSP
- [ ] No secrets committed to version control

**Testing:**
- [ ] Login flow tested end-to-end
- [ ] Token refresh tested (wait for expiry)
- [ ] Logout tested
- [ ] Protected routes require authentication
- [ ] Error states handled gracefully

**Security:**
- [ ] HTTPS enforced (TLS/SSL certificate)
- [ ] CORS configured correctly
- [ ] Backend token validation implemented (if applicable)
- [ ] CSP policy deployed
- [ ] Rate limiting configured (if applicable)

**Monitoring:**
- [ ] Datadog RUM configured (if monitoring enabled)
- [ ] Error logging setup
- [ ] Authentication success/failure metrics tracked

---

## 📝 Document Cross-References

Each document includes references to the others:

- **AUTH-FLOW-DOCUMENTATION.md**
  - Detailed technical implementation
  - Sequences and state machines
  - Code patterns and examples
  - References ENVIRONMENT-VARIABLES and ENTRA-ID-CONFIG docs

- **ENTRA-ID-CONFIG-CHECKLIST.md**
  - Step-by-step setup instructions
  - Configuration screenshots/paths
  - References AUTH-FLOW for protocol details
  - References ENVIRONMENT-VARIABLES for variable setup

- **ENVIRONMENT-VARIABLES-REFERENCE.md**
  - Configuration reference for all variables
  - Ready-to-use code snippets
  - Environment-specific templates
  - References AUTH-FLOW for implementation details

- **Architecture & Sequence Diagrams**
  - Visual representation of all components
  - Used alongside AUTH-FLOW-DOCUMENTATION
  - Supplement all three markdown documents

---

## 🎓 Learning Path

**For new team members onboarding to this auth pattern:**

1. **Start here:** Read this summary document (5 min)
2. **Understand the flow:** Review architecture and sequence diagrams (10 min)
3. **Learn the protocol:** Read AUTH-FLOW-DOCUMENTATION.md sections 1-4 (20 min)
4. **See code examples:** Review ENVIRONMENT-VARIABLES-REFERENCE.md snippets 1-5 (15 min)
5. **Implement:** Follow ENTRA-ID-CONFIG-CHECKLIST.md step-by-step (4-6 hours)
6. **Reference:** Use all three markdown docs and diagrams during development

**Total learning time:** ~45 min theory + 4-6 hours implementation

---

## 🎯 Next Steps

### Immediate Actions

1. **Assign ownership:** Designate team member to lead SWA implementation
2. **Review documentation:** Have team read this package
3. **Create Entra ID App Reg:** Allocate 2-3 hours for setup
4. **Create dev SWA:** Deploy test instance to Azure
5. **Implement MSAL:** Follow the code snippets in Environment Variables doc

### Proof of Concept Timeline

- **Week 1:** Azure AD setup + MSAL integration (Phase 1-2)
- **Week 2:** Testing and validation (Phase 4)
- **Week 3:** Documentation review and handoff

### Production Deployment Timeline

- **After POC validation:** Deploy to staging SWA
- **After staging validation:** Deploy to production SWA
- **Ongoing:** Monitor auth flows with observability tools

---

## 📄 Document Index

| Document | Purpose | Audience | Time to Read |
|----------|---------|----------|--------------|
| **This Summary** | Overview & roadmap | All | 10 min |
| **AUTH-FLOW-DOCUMENTATION.md** | Technical deep-dive | Engineers | 30 min |
| **ENTRA-ID-CONFIG-CHECKLIST.md** | Implementation guide | DevOps/Engineers | 45 min |
| **ENVIRONMENT-VARIABLES-REFERENCE.md** | Configuration reference | Engineers | 20 min |
| **Architecture Diagram** | System visualization | All | 5 min |
| **Sequence Diagram** | OAuth flow visualization | All | 5 min |

---

## 🏆 Success Criteria

Your new SWA authentication implementation is successful when:

✅ Users can log in with Azure AD credentials  
✅ Access tokens acquired and stored securely  
✅ Tokens automatically refreshed before expiry  
✅ Backend APIs accept and validate bearer tokens  
✅ Users can log out and credentials are cleared  
✅ OAuth errors handled gracefully with user feedback  
✅ Monitoring and logging capture auth flow metrics  
✅ No sensitive data logged or exposed to frontend  

---

## 📧 Questions or Issues?

When implementing this auth pattern:

1. **Protocol questions?** → See AUTH-FLOW-DOCUMENTATION.md
2. **Setup questions?** → See ENTRA-ID-CONFIG-CHECKLIST.md
3. **Configuration questions?** → See ENVIRONMENT-VARIABLES-REFERENCE.md
4. **Visual understanding?** → See Architecture/Sequence diagrams
5. **Code snippets needed?** → See ENVIRONMENT-VARIABLES-REFERENCE.md snippets 1-10

---

**End of Documentation Package Summary**

This comprehensive documentation package is complete and ready for implementation. Proceed with Phase 1 (Azure AD Setup) when team is ready.

