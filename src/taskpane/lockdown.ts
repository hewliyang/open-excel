import "ses";

/* global lockdown */

const PRESERVE_FUNCTION_PROPS = new Set(["length", "name", "prototype"]);
const OFFICE_FUNCTION_STUBS = new Set([
  "_validateParams",
  "_validateParameterCount",
  "_validateParameter",
  "_validateParameterType",
]);

function saveFunctionProperties(): Map<string, unknown> {
  const saved = new Map<string, unknown>();
  for (const key of Object.getOwnPropertyNames(Function)) {
    if (!PRESERVE_FUNCTION_PROPS.has(key)) {
      saved.set(key, (Function as unknown as Record<string, unknown>)[key]);
    }
  }
  return saved;
}

function restoreFunctionProperties(saved: Map<string, unknown>) {
  if (saved.size === 0) return;

  const noop = () => null;
  for (const key of OFFICE_FUNCTION_STUBS) {
    if (saved.has(key)) saved.set(key, noop);
  }

  const fn = globalThis.Function;

  // Fast path: if Function is still extensible, define properties directly
  if (Object.isExtensible(fn)) {
    try {
      for (const [key, value] of saved) {
        Object.defineProperty(fn, key, {
          value,
          writable: false,
          enumerable: false,
          configurable: false,
        });
      }
      return;
    } catch {
      // Fall through to proxy path
    }
  }

  // Slow path: Function is frozen by lockdown, wrap it in a Proxy
  const desc = Object.getOwnPropertyDescriptor(globalThis, "Function");
  if (!desc || (!desc.writable && !desc.configurable)) {
    console.warn("[lockdown] Cannot restore Function properties — Function is neither writable nor configurable");
    return;
  }

  const proxy = new Proxy(fn, {
    get(target, prop, receiver) {
      return saved.has(prop as string) ? saved.get(prop as string) : Reflect.get(target, prop, receiver);
    },
    has(target, prop) {
      return saved.has(prop as string) ? true : Reflect.has(target, prop);
    },
    getOwnPropertyDescriptor(target, prop) {
      if (saved.has(prop as string)) {
        return { value: saved.get(prop as string), writable: false, enumerable: false, configurable: true };
      }
      return Reflect.getOwnPropertyDescriptor(target, prop);
    },
  });

  Object.defineProperty(globalThis, "Function", {
    value: proxy,
    writable: desc.writable ?? false,
    configurable: desc.configurable ?? false,
  });
}

let locked = false;

export function ensureLockdown() {
  if (locked) return;
  locked = true;
  try {
    const savedFnProps = saveFunctionProperties();
    lockdown({
      errorTaming: "unsafe",
      consoleTaming: "unsafe",
      overrideTaming: "severe",
      stackFiltering: "verbose",
    });
    restoreFunctionProperties(savedFnProps);
  } catch (e) {
    if (!(e instanceof TypeError && String(e).includes("SES_ALREADY_LOCKED_DOWN"))) {
      throw e;
    }
  }
}
