import {
  Check,
  ChevronDown,
  ChevronUp,
  ExternalLink,
  Eye,
  EyeOff,
  FolderUp,
  LogOut,
  Plus,
  Trash2,
} from "lucide-react";
import { useCallback, useRef, useState } from "react";
import {
  buildAuthorizationUrl,
  exchangeOAuthCode,
  generatePKCE,
  loadOAuthCredentials,
  OAUTH_PROVIDERS,
  type OAuthFlowState,
  removeOAuthCredentials,
  saveOAuthCredentials,
} from "../../../lib/oauth";
import {
  API_TYPES,
  loadSavedConfig,
  saveConfig,
  THINKING_LEVELS,
  type ThinkingLevel,
} from "../../../lib/provider-config";
import { loadWebConfig, saveWebConfig } from "../../../lib/web/config";
import { listFetchProviders } from "../../../lib/web/fetch";
import { listSearchProviders } from "../../../lib/web/search";
import { useChat } from "./chat-context";

function SkillsSection() {
  const { state, installSkill, uninstallSkill } = useChat();
  const folderInputRef = useRef<HTMLInputElement>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);
  const [installing, setInstalling] = useState(false);

  const handleFolderSelect = useCallback(
    async (e: React.ChangeEvent<HTMLInputElement>) => {
      const files = e.target.files;
      if (!files || files.length === 0) return;
      setInstalling(true);
      try {
        await installSkill(Array.from(files));
      } finally {
        setInstalling(false);
        if (folderInputRef.current) folderInputRef.current.value = "";
      }
    },
    [installSkill],
  );

  const handleFileSelect = useCallback(
    async (e: React.ChangeEvent<HTMLInputElement>) => {
      const files = e.target.files;
      if (!files || files.length === 0) return;
      setInstalling(true);
      try {
        await installSkill(Array.from(files));
      } finally {
        setInstalling(false);
        if (fileInputRef.current) fileInputRef.current.value = "";
      }
    },
    [installSkill],
  );

  return (
    <div>
      <div className="text-[10px] uppercase tracking-widest text-(--chat-text-muted) mb-4">
        agent skills
      </div>

      <div className="space-y-3">
        {state.skills.length > 0 ? (
          <div className="space-y-1">
            {state.skills.map((skill) => (
              <div
                key={skill.name}
                className="flex items-start justify-between gap-2 px-3 py-2 bg-(--chat-input-bg) border border-(--chat-border)"
                style={{ borderRadius: "var(--chat-radius)" }}
              >
                <div className="min-w-0 flex-1">
                  <div className="text-xs text-(--chat-text-primary) font-medium truncate">
                    {skill.name}
                  </div>
                  <div className="text-[10px] text-(--chat-text-muted) mt-0.5 line-clamp-2">
                    {skill.description}
                  </div>
                </div>
                <button
                  type="button"
                  onClick={() => uninstallSkill(skill.name)}
                  className="shrink-0 p-1 text-(--chat-text-muted) hover:text-(--chat-error) transition-colors"
                  title="Remove skill"
                >
                  <Trash2 size={12} />
                </button>
              </div>
            ))}
          </div>
        ) : (
          <p className="text-xs text-(--chat-text-muted)">
            No skills installed
          </p>
        )}

        <div className="flex gap-2">
          <button
            type="button"
            onClick={() => folderInputRef.current?.click()}
            disabled={installing}
            className="flex-1 flex items-center justify-center gap-1.5 px-3 py-2 text-xs
                       bg-(--chat-input-bg) border border-(--chat-border) text-(--chat-text-secondary)
                       hover:border-(--chat-border-active) hover:text-(--chat-text-primary)
                       disabled:opacity-50 disabled:cursor-not-allowed transition-colors"
            style={{ borderRadius: "var(--chat-radius)" }}
          >
            <FolderUp size={12} />
            {installing ? "Installing…" : "Add Folder"}
          </button>
          <button
            type="button"
            onClick={() => fileInputRef.current?.click()}
            disabled={installing}
            className="flex-1 flex items-center justify-center gap-1.5 px-3 py-2 text-xs
                       bg-(--chat-input-bg) border border-(--chat-border) text-(--chat-text-secondary)
                       hover:border-(--chat-border-active) hover:text-(--chat-text-primary)
                       disabled:opacity-50 disabled:cursor-not-allowed transition-colors"
            style={{ borderRadius: "var(--chat-radius)" }}
          >
            <Plus size={12} />
            {installing ? "Installing…" : "Add File"}
          </button>
        </div>
        <p className="text-[10px] text-(--chat-text-muted)">
          Add a skill folder or a single SKILL.md file. Skills must have valid
          frontmatter with name and description.
        </p>
      </div>

      <input
        ref={folderInputRef}
        type="file"
        className="hidden"
        {...({
          webkitdirectory: "",
          directory: "",
        } as React.InputHTMLAttributes<HTMLInputElement>)}
        onChange={handleFolderSelect}
      />
      <input
        ref={fileInputRef}
        type="file"
        accept=".md"
        className="hidden"
        onChange={handleFileSelect}
      />
    </div>
  );
}

export function SettingsPanel() {
  const { state, setProviderConfig, availableProviders, getModelsForProvider } =
    useChat();

  const [saved] = useState(loadSavedConfig);
  const [provider, setProvider] = useState(() => saved?.provider || "");
  const [apiKey, setApiKey] = useState(() => saved?.apiKey || "");
  const [model, setModel] = useState(() => saved?.model || "");
  const [showKey, setShowKey] = useState(false);
  const [useProxy, setUseProxy] = useState(() => saved?.useProxy !== false);
  const [proxyUrl, setProxyUrl] = useState(() => saved?.proxyUrl || "");
  const [thinking, setThinking] = useState<ThinkingLevel>(
    () => saved?.thinking || "none",
  );
  const [apiType, setApiType] = useState(
    () => saved?.apiType || "openai-completions",
  );
  const [customBaseUrl, setCustomBaseUrl] = useState(
    () => saved?.customBaseUrl || "",
  );
  const [authMethod, setAuthMethod] = useState<"apikey" | "oauth">(
    () => saved?.authMethod || "apikey",
  );

  const [savedWeb] = useState(loadWebConfig);
  const [webSearchProvider, setWebSearchProvider] = useState(
    () => savedWeb.searchProvider,
  );
  const [webFetchProvider, setWebFetchProvider] = useState(
    () => savedWeb.fetchProvider,
  );
  const [braveApiKey, setBraveApiKey] = useState(
    () => savedWeb.apiKeys.brave || "",
  );
  const [serperApiKey, setSerperApiKey] = useState(
    () => savedWeb.apiKeys.serper || "",
  );
  const [exaApiKey, setExaApiKey] = useState(() => savedWeb.apiKeys.exa || "");
  const [showAdvancedWebKeys, setShowAdvancedWebKeys] = useState(false);

  // OAuth flow state
  const [oauthFlow, setOauthFlow] = useState<OAuthFlowState>(() => {
    if (saved?.authMethod === "oauth") {
      const creds = loadOAuthCredentials(saved.provider);
      return creds ? { step: "connected" } : { step: "idle" };
    }
    return { step: "idle" };
  });
  const [oauthCodeInput, setOauthCodeInput] = useState("");

  const followMode = state.providerConfig?.followMode ?? true;
  const isCustom = provider === "custom";

  const updateAndSync = useCallback(
    (
      updates: Partial<{
        provider: string;
        apiKey: string;
        model: string;
        useProxy: boolean;
        proxyUrl: string;
        thinking: ThinkingLevel;
        apiType: string;
        customBaseUrl: string;
        authMethod: "apikey" | "oauth";
      }>,
    ) => {
      const p = updates.provider ?? provider;
      const k = updates.apiKey ?? apiKey;
      const m = updates.model ?? model;
      const up = updates.useProxy ?? useProxy;
      const pu = updates.proxyUrl ?? proxyUrl;
      const t = updates.thinking ?? thinking;
      const at = updates.apiType ?? apiType;
      const cb = updates.customBaseUrl ?? customBaseUrl;
      const am = updates.authMethod ?? authMethod;

      if ("provider" in updates) setProvider(p);
      if ("apiKey" in updates) setApiKey(k);
      if ("model" in updates) setModel(m);
      if ("useProxy" in updates) setUseProxy(up);
      if ("proxyUrl" in updates) setProxyUrl(pu);
      if ("thinking" in updates) setThinking(t);
      if ("apiType" in updates) setApiType(at);
      if ("customBaseUrl" in updates) setCustomBaseUrl(cb);
      if ("authMethod" in updates) setAuthMethod(am);

      const isCustomEndpoint = p === "custom";
      const isValid = isCustomEndpoint ? p && at && cb && m && k : p && k && m;

      if (isValid) {
        saveConfig({
          provider: p,
          apiKey: k,
          model: m,
          useProxy: up,
          proxyUrl: pu,
          thinking: t,
          followMode,
          apiType: at,
          customBaseUrl: cb,
          authMethod: am,
        });
        setProviderConfig({
          provider: p,
          apiKey: k,
          model: m,
          useProxy: up,
          proxyUrl: pu,
          thinking: t,
          followMode,
          apiType: at,
          customBaseUrl: cb,
          authMethod: am,
        });
      }
    },
    [
      provider,
      apiKey,
      model,
      useProxy,
      proxyUrl,
      thinking,
      apiType,
      customBaseUrl,
      authMethod,
      followMode,
      setProviderConfig,
    ],
  );

  const models = provider && !isCustom ? getModelsForProvider(provider) : [];

  const hasOAuth = provider in OAUTH_PROVIDERS;
  const searchProviders = listSearchProviders();
  const fetchProviders = listFetchProviders();
  const needsBraveKey = webSearchProvider === "brave";
  const needsSerperKey = webSearchProvider === "serper";
  const needsExaKey = webSearchProvider === "exa" || webFetchProvider === "exa";

  const updateWebSettings = useCallback(
    (
      updates: Partial<{
        searchProvider: string;
        fetchProvider: string;
        braveApiKey: string;
        serperApiKey: string;
        exaApiKey: string;
      }>,
    ) => {
      const nextSearchProvider = updates.searchProvider ?? webSearchProvider;
      const nextFetchProvider = updates.fetchProvider ?? webFetchProvider;
      const nextBraveApiKey = updates.braveApiKey ?? braveApiKey;
      const nextSerperApiKey = updates.serperApiKey ?? serperApiKey;
      const nextExaApiKey = updates.exaApiKey ?? exaApiKey;

      if ("searchProvider" in updates) setWebSearchProvider(nextSearchProvider);
      if ("fetchProvider" in updates) setWebFetchProvider(nextFetchProvider);
      if ("braveApiKey" in updates) setBraveApiKey(nextBraveApiKey);
      if ("serperApiKey" in updates) setSerperApiKey(nextSerperApiKey);
      if ("exaApiKey" in updates) setExaApiKey(nextExaApiKey);

      saveWebConfig({
        searchProvider: nextSearchProvider,
        fetchProvider: nextFetchProvider,
        apiKeys: {
          brave: nextBraveApiKey,
          serper: nextSerperApiKey,
          exa: nextExaApiKey,
        },
      });
    },
    [webSearchProvider, webFetchProvider, braveApiKey, serperApiKey, exaApiKey],
  );

  const handleProviderChange = (newProvider: string) => {
    if (newProvider === "custom") {
      updateAndSync({ provider: newProvider, model: "", authMethod: "apikey" });
    } else {
      const providerModels = newProvider
        ? getModelsForProvider(newProvider)
        : [];
      const keepOAuth = newProvider in OAUTH_PROVIDERS ? authMethod : "apikey";
      updateAndSync({
        provider: newProvider,
        model: providerModels[0]?.id || "",
        authMethod: keepOAuth,
      });
    }
    // Reset OAuth flow when switching away from an OAuth-capable provider
    if (!(newProvider in OAUTH_PROVIDERS)) {
      setOauthFlow({ step: "idle" });
    }
  };

  const handleAuthMethodChange = (newMethod: "apikey" | "oauth") => {
    if (newMethod === "oauth") {
      const creds = loadOAuthCredentials(provider);
      if (creds) {
        setOauthFlow({ step: "connected" });
        updateAndSync({ authMethod: "oauth", apiKey: creds.access });
      } else {
        setAuthMethod("oauth");
        setOauthFlow({ step: "idle" });
      }
    } else {
      setOauthFlow({ step: "idle" });
      updateAndSync({ authMethod: "apikey", apiKey: "" });
    }
  };

  const startOAuthLogin = async () => {
    try {
      const { verifier, challenge } = await generatePKCE();
      const { url, oauthState } = buildAuthorizationUrl(
        provider,
        challenge,
        verifier,
      );
      window.open(url, "_blank");
      setOauthFlow({ step: "awaiting-code", verifier, oauthState });
    } catch (err) {
      setOauthFlow({
        step: "error",
        message: err instanceof Error ? err.message : "Failed to start OAuth",
      });
    }
  };

  const submitOAuthCode = async () => {
    if (oauthFlow.step !== "awaiting-code" || !oauthCodeInput.trim()) return;
    const { verifier } = oauthFlow;
    setOauthFlow({ step: "exchanging" });

    try {
      const creds = await exchangeOAuthCode({
        provider,
        rawInput: oauthCodeInput.trim(),
        verifier,
        expectedState: oauthFlow.oauthState,
        useProxy,
        proxyUrl,
      });
      saveOAuthCredentials(provider, creds);
      setOauthFlow({ step: "connected" });
      setOauthCodeInput("");
      updateAndSync({ apiKey: creds.access, authMethod: "oauth" });
    } catch (err) {
      setOauthFlow({
        step: "error",
        message: err instanceof Error ? err.message : "OAuth failed",
      });
    }
  };

  const logoutOAuth = () => {
    removeOAuthCredentials(provider);
    setOauthFlow({ step: "idle" });
    updateAndSync({ authMethod: "apikey", apiKey: "" });
  };

  const isConfigured = state.providerConfig !== null;
  const showApiKeyInput = !(hasOAuth && authMethod === "oauth");

  const inputStyle = {
    borderRadius: "var(--chat-radius)",
    fontFamily: "var(--chat-font-mono)",
  };

  return (
    <div
      className="flex-1 overflow-y-auto p-4 space-y-6"
      style={{ fontFamily: "var(--chat-font-mono)" }}
    >
      <div>
        <div className="text-[10px] uppercase tracking-widest text-(--chat-text-muted) mb-4">
          api configuration
        </div>

        <div className="space-y-4">
          {/* Provider */}
          <label className="block">
            <span className="block text-xs text-(--chat-text-secondary) mb-1.5">
              Provider
            </span>
            <select
              value={provider}
              onChange={(e) => handleProviderChange(e.target.value)}
              className="w-full bg-(--chat-input-bg) text-(--chat-text-primary)
                         text-sm px-3 py-2 border border-(--chat-border)
                         focus:outline-none focus:border-(--chat-border-active)"
              style={inputStyle}
            >
              <option value="">Select provider...</option>
              {availableProviders.map((p) => (
                <option key={p} value={p}>
                  {p}
                </option>
              ))}
              <option disabled>──────────</option>
              <option value="custom">Custom Endpoint</option>
            </select>
          </label>

          {/* Custom Endpoint fields */}
          {isCustom && (
            <>
              <label className="block">
                <span className="block text-xs text-(--chat-text-secondary) mb-1.5">
                  API Type
                </span>
                <select
                  value={apiType}
                  onChange={(e) => updateAndSync({ apiType: e.target.value })}
                  className="w-full bg-(--chat-input-bg) text-(--chat-text-primary)
                             text-sm px-3 py-2 border border-(--chat-border)
                             focus:outline-none focus:border-(--chat-border-active)"
                  style={inputStyle}
                >
                  {API_TYPES.map((at) => (
                    <option key={at.id} value={at.id}>
                      {at.name}
                    </option>
                  ))}
                </select>
                <p className="text-[10px] text-(--chat-text-muted) mt-1">
                  {API_TYPES.find((at) => at.id === apiType)?.hint}
                </p>
              </label>

              <label className="block">
                <span className="block text-xs text-(--chat-text-secondary) mb-1.5">
                  Base URL
                </span>
                <input
                  type="text"
                  value={customBaseUrl}
                  onChange={(e) =>
                    updateAndSync({ customBaseUrl: e.target.value })
                  }
                  placeholder="https://api.openai.com/v1"
                  className="w-full bg-(--chat-input-bg) text-(--chat-text-primary)
                             text-sm px-3 py-2 border border-(--chat-border)
                             placeholder:text-(--chat-text-muted)
                             focus:outline-none focus:border-(--chat-border-active)"
                  style={inputStyle}
                />
                <p className="text-[10px] text-(--chat-text-muted) mt-1">
                  The API endpoint URL for your provider
                </p>
              </label>

              <label className="block">
                <span className="block text-xs text-(--chat-text-secondary) mb-1.5">
                  Model ID
                </span>
                <input
                  type="text"
                  value={model}
                  onChange={(e) => updateAndSync({ model: e.target.value })}
                  placeholder="gpt-4o"
                  className="w-full bg-(--chat-input-bg) text-(--chat-text-primary)
                             text-sm px-3 py-2 border border-(--chat-border)
                             placeholder:text-(--chat-text-muted)
                             focus:outline-none focus:border-(--chat-border-active)"
                  style={inputStyle}
                />
              </label>
            </>
          )}

          {/* Model dropdown — built-in providers only */}
          {!isCustom && provider && (
            <label className="block">
              <span className="block text-xs text-(--chat-text-secondary) mb-1.5">
                Model
              </span>
              <select
                value={model}
                onChange={(e) => updateAndSync({ model: e.target.value })}
                disabled={!provider}
                className="w-full bg-(--chat-input-bg) text-(--chat-text-primary)
                           text-sm px-3 py-2 border border-(--chat-border)
                           focus:outline-none focus:border-(--chat-border-active)
                           disabled:opacity-50 disabled:cursor-not-allowed"
                style={inputStyle}
              >
                <option value="">Select model...</option>
                {models.map((m) => (
                  <option key={m.id} value={m.id}>
                    {m.name}
                  </option>
                ))}
              </select>
            </label>
          )}

          {/* Auth method toggle — providers with OAuth support */}
          {hasOAuth && (
            <div>
              <span className="block text-xs text-(--chat-text-secondary) mb-1.5">
                Authentication
              </span>
              <div className="flex gap-1">
                <button
                  type="button"
                  onClick={() => handleAuthMethodChange("apikey")}
                  className={`flex-1 py-1.5 text-xs border transition-colors ${
                    authMethod === "apikey"
                      ? "bg-(--chat-accent) border-(--chat-accent) text-white"
                      : "bg-(--chat-input-bg) border-(--chat-border) text-(--chat-text-secondary) hover:border-(--chat-border-active)"
                  }`}
                  style={{ borderRadius: "var(--chat-radius)" }}
                >
                  API Key
                </button>
                <button
                  type="button"
                  onClick={() => handleAuthMethodChange("oauth")}
                  className={`flex-1 py-1.5 text-xs border transition-colors ${
                    authMethod === "oauth"
                      ? "bg-(--chat-accent) border-(--chat-accent) text-white"
                      : "bg-(--chat-input-bg) border-(--chat-border) text-(--chat-text-secondary) hover:border-(--chat-border-active)"
                  }`}
                  style={{ borderRadius: "var(--chat-radius)" }}
                >
                  {OAUTH_PROVIDERS[provider]?.label ?? "OAuth"}
                </button>
              </div>
            </div>
          )}

          {/* OAuth flow — providers with OAuth support */}
          {hasOAuth && authMethod === "oauth" && (
            <div className="space-y-2">
              {oauthFlow.step === "idle" && (
                <button
                  type="button"
                  onClick={startOAuthLogin}
                  className="w-full flex items-center justify-center gap-2 px-3 py-2.5 text-xs
                             bg-(--chat-input-bg) border border-(--chat-border) text-(--chat-text-primary)
                             hover:border-(--chat-accent) hover:text-(--chat-accent) transition-colors"
                  style={{ borderRadius: "var(--chat-radius)" }}
                >
                  <ExternalLink size={12} />
                  {OAUTH_PROVIDERS[provider]?.buttonText ?? "Login"}
                </button>
              )}

              {oauthFlow.step === "awaiting-code" && (
                <div className="space-y-2">
                  <p className="text-[10px] text-(--chat-text-muted)">
                    {provider === "openai-codex"
                      ? "Complete login in the opened tab. The page will redirect to localhost and fail — copy the full URL from your browser's address bar and paste it below:"
                      : "Authorize in the opened tab, then paste the code shown on the redirect page:"}
                  </p>
                  <div className="flex gap-1">
                    <input
                      type="text"
                      value={oauthCodeInput}
                      onChange={(e) => setOauthCodeInput(e.target.value)}
                      placeholder={
                        provider === "openai-codex"
                          ? "Paste the full redirect URL here"
                          : "Paste code#state here"
                      }
                      className="flex-1 bg-(--chat-input-bg) text-(--chat-text-primary)
                                 text-sm px-3 py-2 border border-(--chat-border)
                                 placeholder:text-(--chat-text-muted)
                                 focus:outline-none focus:border-(--chat-border-active)"
                      style={inputStyle}
                      onKeyDown={(e) => e.key === "Enter" && submitOAuthCode()}
                    />
                    <button
                      type="button"
                      onClick={submitOAuthCode}
                      disabled={!oauthCodeInput.trim()}
                      className="px-3 py-2 text-xs bg-(--chat-accent) text-white border border-(--chat-accent)
                                 hover:opacity-90 disabled:opacity-50 disabled:cursor-not-allowed transition-colors"
                      style={{ borderRadius: "var(--chat-radius)" }}
                    >
                      Submit
                    </button>
                  </div>
                  <p className="text-[10px] text-(--chat-text-muted)">
                    Requires CORS proxy to be enabled for token exchange.
                  </p>
                </div>
              )}

              {oauthFlow.step === "exchanging" && (
                <div
                  className="px-3 py-2.5 text-xs text-(--chat-text-muted) bg-(--chat-input-bg) border border-(--chat-border)"
                  style={{ borderRadius: "var(--chat-radius)" }}
                >
                  Exchanging authorization code…
                </div>
              )}

              {oauthFlow.step === "connected" && (
                <div
                  className="flex items-center justify-between px-3 py-2.5 bg-(--chat-input-bg) border border-(--chat-border)"
                  style={{ borderRadius: "var(--chat-radius)" }}
                >
                  <div className="flex items-center gap-2 text-xs">
                    <Check size={12} className="text-(--chat-success)" />
                    <span className="text-(--chat-text-secondary)">
                      Connected via OAuth
                    </span>
                  </div>
                  <button
                    type="button"
                    onClick={logoutOAuth}
                    className="flex items-center gap-1 text-[10px] text-(--chat-text-muted) hover:text-(--chat-error) transition-colors"
                  >
                    <LogOut size={10} />
                    Logout
                  </button>
                </div>
              )}

              {oauthFlow.step === "error" && (
                <div className="space-y-2">
                  <div
                    className="px-3 py-2 text-xs text-(--chat-error) bg-(--chat-input-bg) border border-(--chat-error)/30"
                    style={{ borderRadius: "var(--chat-radius)" }}
                  >
                    {oauthFlow.message}
                  </div>
                  <button
                    type="button"
                    onClick={() => setOauthFlow({ step: "idle" })}
                    className="text-[10px] text-(--chat-text-muted) hover:text-(--chat-text-secondary) transition-colors"
                  >
                    Try again
                  </button>
                </div>
              )}
            </div>
          )}

          {/* API Key input — hidden when using OAuth */}
          {showApiKeyInput && (
            <label className="block">
              <span className="block text-xs text-(--chat-text-secondary) mb-1.5">
                API Key
              </span>
              <div className="relative">
                <input
                  type={showKey ? "text" : "password"}
                  value={apiKey}
                  onChange={(e) => updateAndSync({ apiKey: e.target.value })}
                  placeholder="Enter your API key"
                  className="w-full bg-(--chat-input-bg) text-(--chat-text-primary)
                             text-sm px-3 py-2 pr-10 border border-(--chat-border)
                             placeholder:text-(--chat-text-muted)
                             focus:outline-none focus:border-(--chat-border-active)"
                  style={inputStyle}
                />
                <button
                  type="button"
                  onClick={() => setShowKey(!showKey)}
                  className="absolute right-2 top-1/2 -translate-y-1/2 text-(--chat-text-muted)
                             hover:text-(--chat-text-secondary)"
                >
                  {showKey ? <EyeOff size={14} /> : <Eye size={14} />}
                </button>
              </div>
            </label>
          )}

          {/* CORS Proxy */}
          <div className="flex items-center justify-between">
            <div>
              <span className="text-xs text-(--chat-text-secondary)">
                CORS Proxy
              </span>
              <p className="text-[10px] text-(--chat-text-muted) mt-0.5">
                Required for Anthropic and some providers
              </p>
            </div>
            <button
              type="button"
              onClick={() => updateAndSync({ useProxy: !useProxy })}
              className={`
                w-10 h-5 rounded-full transition-colors relative
                ${useProxy ? "bg-(--chat-accent)" : "bg-(--chat-border)"}
              `}
            >
              <span
                className={`
                  absolute top-0.5 w-4 h-4 rounded-full bg-white transition-transform
                  ${useProxy ? "left-5" : "left-0.5"}
                `}
              />
            </button>
          </div>

          {useProxy && (
            <label className="block">
              <span className="block text-xs text-(--chat-text-secondary) mb-1.5">
                Proxy URL
              </span>
              <input
                type="text"
                value={proxyUrl}
                onChange={(e) => updateAndSync({ proxyUrl: e.target.value })}
                placeholder="https://your-proxy.com/proxy"
                className="w-full bg-(--chat-input-bg) text-(--chat-text-primary)
                           text-sm px-3 py-2 border border-(--chat-border)
                           placeholder:text-(--chat-text-muted)
                           focus:outline-none focus:border-(--chat-border-active)"
                style={inputStyle}
              />
              <p className="text-[10px] text-(--chat-text-muted) mt-1">
                Your proxy should accept ?url=encoded_url format
              </p>
            </label>
          )}

          {/* Thinking Level */}
          <div>
            <span className="block text-xs text-(--chat-text-secondary) mb-1.5">
              Thinking Level
            </span>
            <div className="flex gap-1">
              {THINKING_LEVELS.map((level) => (
                <button
                  key={level.value}
                  type="button"
                  onClick={() => updateAndSync({ thinking: level.value })}
                  className={`
                    flex-1 py-1.5 text-xs border transition-colors
                    ${
                      thinking === level.value
                        ? "bg-(--chat-accent) border-(--chat-accent) text-white"
                        : "bg-(--chat-input-bg) border-(--chat-border) text-(--chat-text-secondary) hover:border-(--chat-border-active)"
                    }
                  `}
                  style={{ borderRadius: "var(--chat-radius)" }}
                >
                  {level.label}
                </button>
              ))}
            </div>
            <p className="text-[10px] text-(--chat-text-muted) mt-1">
              Extended thinking for supported models
            </p>
          </div>

          <div className="border-t border-(--chat-border) pt-4 space-y-3">
            <div className="text-[10px] uppercase tracking-widest text-(--chat-text-muted)">
              web tools
            </div>

            <label className="block">
              <span className="block text-xs text-(--chat-text-secondary) mb-1.5">
                Default Search Provider
              </span>
              <select
                value={webSearchProvider}
                onChange={(e) =>
                  updateWebSettings({ searchProvider: e.target.value })
                }
                className="w-full bg-(--chat-input-bg) text-(--chat-text-primary)
                           text-sm px-3 py-2 border border-(--chat-border)
                           focus:outline-none focus:border-(--chat-border-active)"
                style={inputStyle}
              >
                {searchProviders.map((p) => (
                  <option key={p.id} value={p.id}>
                    {p.label}
                  </option>
                ))}
              </select>
              <p className="text-[10px] text-(--chat-text-muted) mt-1">
                Used by web-search.
              </p>
            </label>

            <label className="block">
              <span className="block text-xs text-(--chat-text-secondary) mb-1.5">
                Default Fetch Provider
              </span>
              <select
                value={webFetchProvider}
                onChange={(e) =>
                  updateWebSettings({ fetchProvider: e.target.value })
                }
                className="w-full bg-(--chat-input-bg) text-(--chat-text-primary)
                           text-sm px-3 py-2 border border-(--chat-border)
                           focus:outline-none focus:border-(--chat-border-active)"
                style={inputStyle}
              >
                {fetchProviders.map((p) => (
                  <option key={p} value={p}>
                    {p}
                  </option>
                ))}
              </select>
              <p className="text-[10px] text-(--chat-text-muted) mt-1">
                Used by web-fetch.
              </p>
            </label>

            {needsBraveKey && (
              <label className="block">
                <span className="block text-xs text-(--chat-text-secondary) mb-1.5">
                  Brave API Key
                </span>
                <input
                  type="password"
                  value={braveApiKey}
                  onChange={(e) =>
                    updateWebSettings({ braveApiKey: e.target.value })
                  }
                  placeholder="Required for Brave search"
                  className="w-full bg-(--chat-input-bg) text-(--chat-text-primary)
                           text-sm px-3 py-2 border border-(--chat-border)
                           placeholder:text-(--chat-text-muted)
                           focus:outline-none focus:border-(--chat-border-active)"
                  style={inputStyle}
                />
              </label>
            )}

            {needsSerperKey && (
              <label className="block">
                <span className="block text-xs text-(--chat-text-secondary) mb-1.5">
                  Serper API Key
                </span>
                <input
                  type="password"
                  value={serperApiKey}
                  onChange={(e) =>
                    updateWebSettings({ serperApiKey: e.target.value })
                  }
                  placeholder="Required for Serper search"
                  className="w-full bg-(--chat-input-bg) text-(--chat-text-primary)
                           text-sm px-3 py-2 border border-(--chat-border)
                           placeholder:text-(--chat-text-muted)
                           focus:outline-none focus:border-(--chat-border-active)"
                  style={inputStyle}
                />
              </label>
            )}

            {needsExaKey && (
              <label className="block">
                <span className="block text-xs text-(--chat-text-secondary) mb-1.5">
                  Exa API Key
                </span>
                <input
                  type="password"
                  value={exaApiKey}
                  onChange={(e) =>
                    updateWebSettings({ exaApiKey: e.target.value })
                  }
                  placeholder="Required for Exa search/fetch"
                  className="w-full bg-(--chat-input-bg) text-(--chat-text-primary)
                           text-sm px-3 py-2 border border-(--chat-border)
                           placeholder:text-(--chat-text-muted)
                           focus:outline-none focus:border-(--chat-border-active)"
                  style={inputStyle}
                />
              </label>
            )}

            <div className="pt-1">
              <button
                type="button"
                onClick={() => setShowAdvancedWebKeys(!showAdvancedWebKeys)}
                className="inline-flex items-center gap-1.5 text-xs text-(--chat-text-secondary) hover:text-(--chat-text-primary)"
              >
                {showAdvancedWebKeys ? (
                  <ChevronUp size={12} />
                ) : (
                  <ChevronDown size={12} />
                )}
                <span>
                  {showAdvancedWebKeys ? "Hide" : "Show"} advanced saved API
                  keys
                </span>
              </button>
            </div>

            {showAdvancedWebKeys && (
              <div className="space-y-3 border border-(--chat-border) p-3 bg-(--chat-input-bg)">
                {!needsBraveKey && (
                  <label className="block">
                    <span className="block text-xs text-(--chat-text-secondary) mb-1.5">
                      Brave API Key
                    </span>
                    <input
                      type="password"
                      value={braveApiKey}
                      onChange={(e) =>
                        updateWebSettings({ braveApiKey: e.target.value })
                      }
                      placeholder="Optional"
                      className="w-full bg-(--chat-bg) text-(--chat-text-primary)
                           text-sm px-3 py-2 border border-(--chat-border)
                           placeholder:text-(--chat-text-muted)
                           focus:outline-none focus:border-(--chat-border-active)"
                      style={inputStyle}
                    />
                  </label>
                )}

                {!needsSerperKey && (
                  <label className="block">
                    <span className="block text-xs text-(--chat-text-secondary) mb-1.5">
                      Serper API Key
                    </span>
                    <input
                      type="password"
                      value={serperApiKey}
                      onChange={(e) =>
                        updateWebSettings({ serperApiKey: e.target.value })
                      }
                      placeholder="Optional"
                      className="w-full bg-(--chat-bg) text-(--chat-text-primary)
                           text-sm px-3 py-2 border border-(--chat-border)
                           placeholder:text-(--chat-text-muted)
                           focus:outline-none focus:border-(--chat-border-active)"
                      style={inputStyle}
                    />
                  </label>
                )}

                {!needsExaKey && (
                  <label className="block">
                    <span className="block text-xs text-(--chat-text-secondary) mb-1.5">
                      Exa API Key
                    </span>
                    <input
                      type="password"
                      value={exaApiKey}
                      onChange={(e) =>
                        updateWebSettings({ exaApiKey: e.target.value })
                      }
                      placeholder="Optional"
                      className="w-full bg-(--chat-bg) text-(--chat-text-primary)
                           text-sm px-3 py-2 border border-(--chat-border)
                           placeholder:text-(--chat-text-muted)
                           focus:outline-none focus:border-(--chat-border-active)"
                      style={inputStyle}
                    />
                  </label>
                )}
              </div>
            )}
          </div>
        </div>
      </div>

      {/* Status */}
      <div className="border-t border-(--chat-border) pt-4">
        <div className="flex items-center gap-2 text-xs">
          {isConfigured ? (
            <>
              <Check size={12} className="text-(--chat-success)" />
              <span className="text-(--chat-text-secondary)">
                Using{" "}
                {state.providerConfig?.provider === "custom"
                  ? `custom (${state.providerConfig?.apiType})`
                  : state.providerConfig?.provider}
                {state.providerConfig?.authMethod === "oauth" && " via OAuth"}
              </span>
            </>
          ) : (
            <span className="text-(--chat-text-muted)">
              Fill in all fields above to get started
            </span>
          )}
        </div>
      </div>

      {/* Skills */}
      <div className="border-t border-(--chat-border) pt-4">
        <SkillsSection />
      </div>

      {/* About */}
      <div className="border-t border-(--chat-border) pt-4">
        <div className="text-[10px] uppercase tracking-widest text-(--chat-text-muted) mb-2">
          about
        </div>
        <p className="text-xs text-(--chat-text-secondary) leading-relaxed">
          OpenExcel uses your own API key to connect to LLM providers. Your key
          is stored locally in the browser.
        </p>
        {isCustom && (
          <p className="text-xs text-(--chat-text-muted) leading-relaxed mt-2">
            Custom Endpoint: Point to any OpenAI-compatible API (Ollama, vLLM,
            LMStudio) or other supported API types.
          </p>
        )}
        {useProxy && (
          <p className="text-xs text-(--chat-text-muted) leading-relaxed mt-2">
            CORS Proxy: Requests route through your proxy to bypass browser CORS
            restrictions. Required for Claude OAuth and some providers.
          </p>
        )}
        <p className="text-[10px] text-(--chat-text-muted) mt-3">
          v{__APP_VERSION__}
        </p>
      </div>
    </div>
  );
}
