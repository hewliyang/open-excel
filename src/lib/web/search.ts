import type {
  SearchOptions,
  SearchProvider,
  SearchResult,
  WebContext,
} from "./types";

function parseHTML(html: string): Document {
  return new DOMParser().parseFromString(html, "text/html");
}

function textOf(el: Element | null): string {
  return (el?.textContent ?? "").replace(/\s+/g, " ").trim();
}

function getApiKey(
  context: WebContext,
  providerId: string,
): string | undefined {
  return context.apiKeys?.[providerId];
}

async function fetchWithProxy(
  url: string,
  context: WebContext,
  init?: RequestInit,
): Promise<Response> {
  if (context.proxyUrl) {
    try {
      return await fetch(
        `${context.proxyUrl}/?url=${encodeURIComponent(url)}`,
        init,
      );
    } catch (err) {
      throw new Error(
        `CORS proxy search failed: ${err instanceof Error ? err.message : String(err)}`,
      );
    }
  }

  try {
    return await fetch(url, init);
  } catch {
    throw new Error(
      "Search blocked by CORS and no CORS proxy is configured. Enable the CORS proxy in Settings.",
    );
  }
}

const ddgsProvider: SearchProvider = {
  id: "ddgs",
  async search(query, options, context) {
    const { region = "us-en", timelimit, maxResults = 10, page = 1 } = options;

    const body = new URLSearchParams({ q: query, l: region, b: "" });
    if (page > 1) body.set("s", String(10 + (page - 2) * 15));
    if (timelimit) body.set("df", timelimit);

    const target = "https://html.duckduckgo.com/html/";

    const resp = await fetchWithProxy(target, context, {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: body.toString(),
    });
    if (!resp.ok) throw new Error(`Search failed: ${resp.status}`);

    const doc = parseHTML(await resp.text());
    if (doc.querySelector(".anomaly-modal, #challenge-form")) {
      throw new Error(
        "DuckDuckGo is rate-limiting requests (bot challenge). Try again later or use a different CORS proxy.",
      );
    }

    const results: SearchResult[] = [];
    for (const item of doc.querySelectorAll(".result")) {
      const titleEl = item.querySelector(".result__title a, .result__a");
      const bodyEl = item.querySelector(".result__snippet");
      let href = titleEl?.getAttribute("href") ?? "";
      if (!href || href.includes("duckduckgo.com/y.js")) continue;
      if (href.startsWith("//")) href = `https:${href}`;

      results.push({
        title: textOf(titleEl),
        href,
        body: textOf(bodyEl),
      });

      if (results.length >= maxResults) break;
    }

    return results;
  },
};

const braveProvider: SearchProvider = {
  id: "brave",
  requiresApiKey: true,
  async search(query, options, context) {
    const apiKey = getApiKey(context, "brave");
    if (!apiKey) {
      throw new Error(
        "Brave search requires an API key. Configure it in openexcel-web-config.apiKeys.brave.",
      );
    }

    const maxResults = options.maxResults ?? 10;
    const offset = ((options.page ?? 1) - 1) * maxResults;
    const country = options.region?.split("-")[0]?.toUpperCase() || "US";
    const url = new URL("https://api.search.brave.com/res/v1/web/search");
    url.searchParams.set("q", query);
    url.searchParams.set("count", String(maxResults));
    url.searchParams.set("offset", String(offset));
    url.searchParams.set("country", country);

    const resp = await fetchWithProxy(url.toString(), context, {
      headers: {
        Accept: "application/json",
        "X-Subscription-Token": apiKey,
      },
    });

    if (!resp.ok) {
      throw new Error(`Brave search failed: ${resp.status} ${resp.statusText}`);
    }

    const data = (await resp.json()) as {
      web?: {
        results?: Array<{ title?: string; url?: string; description?: string }>;
      };
    };

    return (data.web?.results || []).map((r) => ({
      title: r.title || "",
      href: r.url || "",
      body: r.description || "",
    }));
  },
};

const serperProvider: SearchProvider = {
  id: "serper",
  requiresApiKey: true,
  async search(query, options, context) {
    const apiKey = getApiKey(context, "serper");
    if (!apiKey) {
      throw new Error(
        "Serper search requires an API key. Configure it in openexcel-web-config.apiKeys.serper.",
      );
    }

    const [countryRaw, languageRaw] = (options.region || "us-en").split("-");
    const country = countryRaw?.toLowerCase() || "us";
    const language = languageRaw?.toLowerCase() || "en";

    const body: {
      q: string;
      num?: number;
      page?: number;
      gl?: string;
      hl?: string;
      tbs?: string;
    } = {
      q: query,
      num: options.maxResults ?? 10,
      page: options.page ?? 1,
      gl: country,
      hl: language,
    };

    if (options.timelimit) body.tbs = `qdr:${options.timelimit}`;

    const endpoint = "https://google.serper.dev/search";

    const resp = await fetchWithProxy(endpoint, context, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "X-API-KEY": apiKey,
      },
      body: JSON.stringify(body),
    });

    if (!resp.ok) {
      throw new Error(
        `Serper search failed: ${resp.status} ${resp.statusText}`,
      );
    }

    const data = (await resp.json()) as {
      organic?: Array<{ title?: string; link?: string; snippet?: string }>;
    };

    return (data.organic || []).map((r) => ({
      title: r.title || "",
      href: r.link || "",
      body: r.snippet || "",
    }));
  },
};

const exaProvider: SearchProvider = {
  id: "exa",
  requiresApiKey: true,
  async search(query, options, context) {
    const apiKey = getApiKey(context, "exa");
    if (!apiKey) {
      throw new Error(
        "Exa search requires an API key. Configure it in openexcel-web-config.apiKeys.exa.",
      );
    }

    const body = {
      query,
      numResults: options.maxResults ?? 10,
      type: "auto",
    };

    const targetUrl = "https://api.exa.ai/search";

    const resp = await fetchWithProxy(targetUrl, context, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "x-api-key": apiKey,
      },
      body: JSON.stringify(body),
    });

    if (!resp.ok) {
      throw new Error(`Exa search failed: ${resp.status} ${resp.statusText}`);
    }

    const data = (await resp.json()) as {
      results?: Array<{ title?: string; url?: string; text?: string }>;
    };

    return (data.results || []).map((r) => ({
      title: r.title || "",
      href: r.url || "",
      body: r.text || "",
    }));
  },
};

const PROVIDERS: Record<string, SearchProvider> = {
  ddgs: ddgsProvider,
  brave: braveProvider,
  serper: serperProvider,
  exa: exaProvider,
};

const PROVIDER_LABELS: Record<string, string> = {
  ddgs: "ddgs (free, easily rate limited)",
};

export function listSearchProviders(): { id: string; label: string }[] {
  return Object.keys(PROVIDERS).map((id) => ({
    id,
    label: PROVIDER_LABELS[id] ?? id,
  }));
}

export function getSearchProvider(providerId?: string): SearchProvider {
  if (!providerId) return ddgsProvider;
  return PROVIDERS[providerId] || ddgsProvider;
}

export async function searchWeb(
  query: string,
  options: SearchOptions = {},
  context: WebContext = {},
  providerId?: string,
): Promise<SearchResult[]> {
  const provider = getSearchProvider(providerId);
  return provider.search(query, options, context);
}
