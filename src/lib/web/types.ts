export interface SearchResult {
  title: string;
  href: string;
  body: string;
}

export interface SearchOptions {
  region?: string;
  timelimit?: "d" | "w" | "m" | "y";
  maxResults?: number;
  page?: number;
}

export interface WebContext {
  proxyUrl?: string;
  apiKeys?: Record<string, string | undefined>;
}

export interface SearchProvider {
  id: string;
  requiresApiKey?: boolean;
  search: (
    query: string,
    options: SearchOptions,
    context: WebContext,
  ) => Promise<SearchResult[]>;
}

export type FetchResult =
  | {
      kind: "text";
      contentType: string;
      text: string;
      title?: string;
      metadata?: Record<string, string>;
    }
  | {
      kind: "binary";
      contentType: string;
      data: Uint8Array;
    };

export interface FetchProvider {
  id: string;
  requiresApiKey?: boolean;
  fetch: (url: string, context: WebContext) => Promise<FetchResult>;
}
