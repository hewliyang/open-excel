export interface SearchPage {
  totalFound: number;
  returned: number;
  offset: number;
  hasMore: boolean;
  nextOffset: number | null;
}

export interface SearchPageCollector<T> {
  readonly pageOffset: number;
  readonly pageSize: number;
  readonly matches: T[];
  add(match: T): boolean;
  toPage(): SearchPage;
}

export function createSearchPageCollector<T>(
  offset: number,
  maxResults: number,
): SearchPageCollector<T> {
  const pageOffset = Math.max(0, Math.floor(offset));
  const pageSize = Math.max(1, Math.floor(maxResults));

  const matches: T[] = [];
  let totalMatched = 0;
  let hasMore = false;

  return {
    pageOffset,
    pageSize,
    matches,
    add(match: T) {
      if (hasMore) return true;

      const matchIndex = totalMatched;
      totalMatched += 1;

      if (matchIndex < pageOffset) return false;

      if (matches.length < pageSize) {
        matches.push(match);
        return false;
      }

      hasMore = true;
      return true;
    },
    toPage() {
      return {
        totalFound: totalMatched,
        returned: matches.length,
        offset: pageOffset,
        hasMore,
        nextOffset: hasMore ? pageOffset + matches.length : null,
      };
    },
  };
}
