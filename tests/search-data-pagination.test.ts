import { describe, expect, it } from "vitest";
import { createSearchPageCollector } from "../src/lib/excel/search-data-pagination";

describe("createSearchPageCollector", () => {
  it("returns first page and sets hasMore with a sentinel match", () => {
    const collector = createSearchPageCollector<number>(0, 2);

    expect(collector.add(1)).toBe(false);
    expect(collector.add(2)).toBe(false);
    expect(collector.add(3)).toBe(true);

    expect(collector.matches).toEqual([1, 2]);
    expect(collector.toPage()).toEqual({
      totalFound: 3,
      returned: 2,
      offset: 0,
      hasMore: true,
      nextOffset: 2,
    });
  });

  it("handles offset pagination correctly", () => {
    const collector = createSearchPageCollector<number>(2, 2);

    expect(collector.add(10)).toBe(false);
    expect(collector.add(20)).toBe(false);
    expect(collector.add(30)).toBe(false);
    expect(collector.add(40)).toBe(false);
    expect(collector.add(50)).toBe(true);

    expect(collector.matches).toEqual([30, 40]);
    expect(collector.toPage()).toEqual({
      totalFound: 5,
      returned: 2,
      offset: 2,
      hasMore: true,
      nextOffset: 4,
    });
  });

  it("returns empty page when offset is beyond all matches", () => {
    const collector = createSearchPageCollector<number>(10, 5);

    collector.add(1);
    collector.add(2);
    collector.add(3);

    expect(collector.matches).toEqual([]);
    expect(collector.toPage()).toEqual({
      totalFound: 3,
      returned: 0,
      offset: 10,
      hasMore: false,
      nextOffset: null,
    });
  });

  it("normalizes invalid offset and maxResults", () => {
    const collector = createSearchPageCollector<number>(-3.9, 0.4);

    collector.add(7);
    expect(collector.add(8)).toBe(true);

    expect(collector.matches).toEqual([7]);
    expect(collector.toPage()).toEqual({
      totalFound: 2,
      returned: 1,
      offset: 0,
      hasMore: true,
      nextOffset: 1,
    });
  });
});
