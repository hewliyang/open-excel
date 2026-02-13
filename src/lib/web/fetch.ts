import { Readability } from "@mozilla/readability";
import TurndownService from "turndown";
import type { FetchProvider, FetchResult, WebContext } from "./types";

function isHtmlContentType(contentType: string): boolean {
  const ct = contentType.split(";")[0].trim().toLowerCase();
  return ct === "text/html" || ct === "application/xhtml+xml";
}

function isTextContentType(contentType: string): boolean {
  const ct = contentType.split(";")[0].trim().toLowerCase();
  if (ct.startsWith("text/")) return true;
  return (
    ct === "application/json" ||
    ct === "application/xml" ||
    ct === "application/javascript"
  );
}

function getApiKey(
  context: WebContext,
  providerId: string,
): string | undefined {
  return context.apiKeys?.[providerId];
}

async function fetchWithProxy(
  url: string,
  proxyUrl?: string,
  init?: RequestInit,
): Promise<Response> {
  if (proxyUrl) {
    try {
      return await fetch(`${proxyUrl}/?url=${encodeURIComponent(url)}`, init);
    } catch (err) {
      throw new Error(
        `CORS proxy fetch failed: ${err instanceof Error ? err.message : err}`,
      );
    }
  }

  try {
    return await fetch(url, init);
  } catch {
    throw new Error(
      "Fetch blocked by CORS and no CORS proxy is configured. Enable the CORS proxy in Settings.",
    );
  }
}

function extractContentFromHtml(url: string, html: string): FetchResult {
  const doc = new DOMParser().parseFromString(html, "text/html");
  const base = doc.createElement("base");
  base.href = url;
  doc.head.prepend(base);

  const reader = new Readability(doc);
  const article = reader.parse();

  const td = new TurndownService({
    headingStyle: "atx",
    codeBlockStyle: "fenced",
    bulletListMarker: "-",
  });
  td.remove(["script", "style", "footer"]);

  let title: string;
  let content: string;
  const metadata: Record<string, string> = { URL: url };

  if (article) {
    title = article.title;
    if (article.byline) metadata.Author = article.byline;
    if (article.siteName) metadata.Site = article.siteName;
    content = td.turndown(article.content).trim();
  } else {
    title = doc.querySelector("title")?.textContent ?? url;
    content = td.turndown(doc.body?.innerHTML ?? "").trim();
  }

  return {
    kind: "text",
    contentType: "text/markdown",
    text: content,
    title,
    metadata,
  };
}

const basicFetchProvider: FetchProvider = {
  id: "basic",
  async fetch(url, context): Promise<FetchResult> {
    const resp = await fetchWithProxy(url, context.proxyUrl);
    if (!resp.ok) {
      throw new Error(`Fetch failed: ${resp.status} ${resp.statusText}`);
    }

    const contentType = resp.headers.get("content-type") ?? "";

    if (isHtmlContentType(contentType)) {
      return extractContentFromHtml(url, await resp.text());
    }

    if (isTextContentType(contentType)) {
      return {
        kind: "text",
        contentType,
        text: await resp.text(),
      };
    }

    return {
      kind: "binary",
      contentType,
      data: new Uint8Array(await resp.arrayBuffer()),
    };
  },
};

type ExaTextOptions =
  | boolean
  | {
      maxCharacters?: number;
      includeHtmlTags?: boolean;
    };

interface ExaContentsRequest {
  urls: string[];
  text?: ExaTextOptions;
}

interface ExaContentsResult {
  id?: string;
  url?: string;
  title?: string;
  text?: string;
  summary?: string;
  highlights?: string[];
  author?: string;
  publishedDate?: string;
}

interface ExaContentsResponse {
  requestId?: string;
  results?: ExaContentsResult[];
}

function toExaTextResult(
  inputUrl: string,
  response: ExaContentsResponse,
): FetchResult | null {
  const first = response.results?.[0];
  if (!first) return null;

  const text =
    first.text?.trim() ||
    first.summary?.trim() ||
    first.highlights?.filter(Boolean).join("\n\n").trim() ||
    "";

  if (!text) return null;

  const metadata: Record<string, string> = {
    URL: first.url || inputUrl,
  };

  if (first.author) metadata.Author = first.author;
  if (first.publishedDate) metadata.Published = first.publishedDate;
  if (response.requestId) metadata.ExaRequestId = response.requestId;
  if (first.id) metadata.ExaId = first.id;

  return {
    kind: "text",
    contentType: "text/markdown",
    title: first.title || first.url || inputUrl,
    text,
    metadata,
  };
}

const exaFetchProvider: FetchProvider = {
  id: "exa",
  requiresApiKey: true,
  async fetch(url, context): Promise<FetchResult> {
    const apiKey = getApiKey(context, "exa");
    if (!apiKey) {
      throw new Error(
        "Exa fetch requires an API key. Configure it in openexcel-web-config.apiKeys.exa.",
      );
    }

    const endpoint = "https://api.exa.ai/contents";
    const body: ExaContentsRequest = {
      urls: [url],
      text: true,
    };

    const resp = await fetchWithProxy(endpoint, context.proxyUrl, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "x-api-key": apiKey,
      },
      body: JSON.stringify(body),
    });

    if (resp.ok) {
      const json = (await resp.json()) as ExaContentsResponse;
      const parsed = toExaTextResult(url, json);
      if (parsed) return parsed;
      return basicFetchProvider.fetch(url, context);
    }

    if (resp.status === 401 || resp.status === 403) {
      throw new Error(`Exa fetch failed: ${resp.status} ${resp.statusText}`);
    }

    return basicFetchProvider.fetch(url, context);
  },
};

const PROVIDERS: Record<string, FetchProvider> = {
  basic: basicFetchProvider,
  exa: exaFetchProvider,
};

export function listFetchProviders(): string[] {
  return Object.keys(PROVIDERS);
}

export function getFetchProvider(providerId?: string): FetchProvider {
  if (!providerId) return basicFetchProvider;
  return PROVIDERS[providerId] || basicFetchProvider;
}

export async function fetchWeb(
  url: string,
  context: WebContext = {},
  providerId?: string,
): Promise<FetchResult> {
  const provider = getFetchProvider(providerId);
  return provider.fetch(url, context);
}
