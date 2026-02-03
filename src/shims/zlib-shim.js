// Shim for node:zlib in browser
// gzip/gunzip commands won't work in browser, but this prevents build errors

export const constants = {
  Z_DEFAULT_COMPRESSION: -1,
  Z_BEST_COMPRESSION: 9,
  Z_BEST_SPEED: 1,
};

export function gunzipSync() {
  throw new Error("gzip/gunzip not supported in browser environment");
}

export function gzipSync() {
  throw new Error("gzip/gunzip not supported in browser environment");
}

export default {
  constants,
  gunzipSync,
  gzipSync,
};
