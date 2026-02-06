// Browser shim for node:util/types
export function isArrayBuffer(value) {
  return value instanceof ArrayBuffer;
}

export function isTypedArray(value) {
  return ArrayBuffer.isView(value) && !(value instanceof DataView);
}

export function isDataView(value) {
  return value instanceof DataView;
}

export function isSharedArrayBuffer(value) {
  return typeof SharedArrayBuffer !== "undefined" && value instanceof SharedArrayBuffer;
}

export function isDate(value) {
  return value instanceof Date;
}

export function isRegExp(value) {
  return value instanceof RegExp;
}

export function isMap(value) {
  return value instanceof Map;
}

export function isSet(value) {
  return value instanceof Set;
}

export function isPromise(value) {
  return value instanceof Promise;
}
