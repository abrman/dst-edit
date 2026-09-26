import { useCallback, useState } from "react";

const LIMIT = 200;

interface History<T> {
  past: T[];
  present: T;
  future: T[];
}

/** State with undo/redo. Every set() is one undo step. */
export function useHistory<T>(initial: T) {
  const [history, setHistory] = useState<History<T>>({ past: [], present: initial, future: [] });

  const set = useCallback((update: T | ((prev: T) => T)) => {
    setHistory((h) => {
      const next = typeof update === "function" ? (update as (prev: T) => T)(h.present) : update;
      if (next === h.present) return h;
      return { past: [...h.past, h.present].slice(-LIMIT), present: next, future: [] };
    });
  }, []);

  const reset = useCallback((value: T) => setHistory({ past: [], present: value, future: [] }), []);

  const undo = useCallback(() => {
    setHistory((h) =>
      h.past.length ? { past: h.past.slice(0, -1), present: h.past[h.past.length - 1], future: [h.present, ...h.future] } : h,
    );
  }, []);

  const redo = useCallback(() => {
    setHistory((h) =>
      h.future.length ? { past: [...h.past, h.present], present: h.future[0], future: h.future.slice(1) } : h,
    );
  }, []);

  return {
    value: history.present,
    set,
    reset,
    undo,
    redo,
    canUndo: history.past.length > 0,
    canRedo: history.future.length > 0,
  };
}
