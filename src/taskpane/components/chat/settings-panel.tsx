import { Check, Eye, EyeOff } from "lucide-react";
import { useCallback, useState } from "react";
import { type ThinkingLevel, useChat } from "./chat-context";

const STORAGE_KEY = "openexcel-provider-config";

interface SavedConfig {
  provider: string;
  apiKey: string;
  model: string;
  useProxy: boolean;
  proxyUrl: string;
  thinking: ThinkingLevel;
  followMode?: boolean;
}

function loadSavedConfig(): SavedConfig | null {
  try {
    const saved = localStorage.getItem(STORAGE_KEY);
    if (saved) {
      const config = JSON.parse(saved);
      if (config.proxyUrl === undefined) config.proxyUrl = "";
      if (config.followMode === undefined) config.followMode = true;
      return config;
    }
  } catch {}
  return null;
}

function saveConfig(
  provider: string,
  apiKey: string,
  model: string,
  useProxy: boolean,
  proxyUrl: string,
  thinking: ThinkingLevel,
  followMode: boolean,
) {
  localStorage.setItem(
    STORAGE_KEY,
    JSON.stringify({ provider, apiKey, model, useProxy, proxyUrl, thinking, followMode }),
  );
}

const THINKING_LEVELS: { value: ThinkingLevel; label: string }[] = [
  { value: "none", label: "None" },
  { value: "low", label: "Low" },
  { value: "medium", label: "Medium" },
  { value: "high", label: "High" },
];

export function SettingsPanel() {
  const { state, setProviderConfig, availableProviders, getModelsForProvider } = useChat();

  const [saved] = useState(loadSavedConfig);
  const [provider, setProvider] = useState(() => saved?.provider || "");
  const [apiKey, setApiKey] = useState(() => saved?.apiKey || "");
  const [model, setModel] = useState(() => saved?.model || "");
  const [showKey, setShowKey] = useState(false);
  const [useProxy, setUseProxy] = useState(() => saved?.useProxy !== false);
  const [proxyUrl, setProxyUrl] = useState(() => saved?.proxyUrl || "");
  const [thinking, setThinking] = useState<ThinkingLevel>(() => saved?.thinking || "none");

  // Preserve followMode from current state (managed via header toggle)
  const followMode = state.providerConfig?.followMode ?? true;

  const updateAndSync = useCallback(
    (
      updates: Partial<{
        provider: string;
        apiKey: string;
        model: string;
        useProxy: boolean;
        proxyUrl: string;
        thinking: ThinkingLevel;
      }>,
    ) => {
      const p = updates.provider ?? provider;
      const k = updates.apiKey ?? apiKey;
      const m = updates.model ?? model;
      const up = updates.useProxy ?? useProxy;
      const pu = updates.proxyUrl ?? proxyUrl;
      const t = updates.thinking ?? thinking;

      if ("provider" in updates) setProvider(p);
      if ("apiKey" in updates) setApiKey(k);
      if ("model" in updates) setModel(m);
      if ("useProxy" in updates) setUseProxy(up);
      if ("proxyUrl" in updates) setProxyUrl(pu);
      if ("thinking" in updates) setThinking(t);

      if (p && k && m) {
        saveConfig(p, k, m, up, pu, t, followMode);
        setProviderConfig({ provider: p, apiKey: k, model: m, useProxy: up, proxyUrl: pu, thinking: t, followMode });
      }
    },
    [provider, apiKey, model, useProxy, proxyUrl, thinking, followMode, setProviderConfig],
  );

  const models = provider ? getModelsForProvider(provider) : [];

  const handleProviderChange = (newProvider: string) => {
    const providerModels = newProvider ? getModelsForProvider(newProvider) : [];
    updateAndSync({ provider: newProvider, model: providerModels[0]?.id || "" });
  };

  const isConfigured = state.providerConfig !== null;

  const inputStyle = {
    borderRadius: "var(--chat-radius)",
    fontFamily: "var(--chat-font-mono)",
  };

  return (
    <div className="flex-1 overflow-y-auto p-4 space-y-6" style={{ fontFamily: "var(--chat-font-mono)" }}>
      <div>
        <div className="text-[10px] uppercase tracking-widest text-(--chat-text-muted) mb-4">api configuration</div>

        <div className="space-y-4">
          <label className="block">
            <span className="block text-xs text-(--chat-text-secondary) mb-1.5">Provider</span>
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
            </select>
          </label>

          <label className="block">
            <span className="block text-xs text-(--chat-text-secondary) mb-1.5">Model</span>
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

          <label className="block">
            <span className="block text-xs text-(--chat-text-secondary) mb-1.5">API Key</span>
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

          <div className="flex items-center justify-between">
            <div>
              <span className="text-xs text-(--chat-text-secondary)">CORS Proxy</span>
              <p className="text-[10px] text-(--chat-text-muted) mt-0.5">Required for Anthropic and some providers</p>
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
              <span className="block text-xs text-(--chat-text-secondary) mb-1.5">Proxy URL</span>
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

          <div>
            <span className="block text-xs text-(--chat-text-secondary) mb-1.5">Thinking Level</span>
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
            <p className="text-[10px] text-(--chat-text-muted) mt-1">Extended thinking for supported models</p>
          </div>
        </div>
      </div>

      <div className="border-t border-(--chat-border) pt-4">
        <div className="flex items-center gap-2 text-xs">
          {isConfigured ? (
            <>
              <Check size={12} className="text-(--chat-success)" />
              <span className="text-(--chat-text-secondary)">Using {state.providerConfig?.provider}</span>
            </>
          ) : (
            <span className="text-(--chat-text-muted)">Fill in all fields above to get started</span>
          )}
        </div>
      </div>

      <div className="border-t border-(--chat-border) pt-4">
        <div className="text-[10px] uppercase tracking-widest text-(--chat-text-muted) mb-2">about</div>
        <p className="text-xs text-(--chat-text-secondary) leading-relaxed">
          OpenExcel uses your own API key to connect to LLM providers. Your key is stored locally in the browser.
        </p>
        {useProxy && (
          <p className="text-xs text-(--chat-text-muted) leading-relaxed mt-2">
            CORS Proxy: Requests route through your proxy to bypass browser CORS restrictions. Required for Claude OAuth
            and Z.ai.
          </p>
        )}
        <p className="text-[10px] text-(--chat-text-muted) mt-3">v{__APP_VERSION__}</p>
      </div>
    </div>
  );
}
