import { useCallback, useState } from 'react';

/** Multi-row selection for tables. Rows are identified by a stable key (e.g. rowNumber). */
export function useRowSelection<K extends string | number = number>() {
  const [selected, setSelected] = useState<Set<K>>(new Set());

  const has = useCallback((key: K) => selected.has(key), [selected]);

  const toggle = useCallback((key: K) => {
    setSelected((prev) => {
      const next = new Set(prev);
      if (next.has(key)) next.delete(key);
      else next.add(key);
      return next;
    });
  }, []);

  const allChecked = (keys: K[]) => keys.length > 0 && keys.every((k) => selected.has(k));

  const toggleAll = useCallback((keys: K[]) => {
    setSelected((prev) => {
      const everyChecked = keys.length > 0 && keys.every((k) => prev.has(k));
      return everyChecked ? new Set() : new Set(keys);
    });
  }, []);

  const clear = useCallback(() => setSelected(new Set()), []);

  return { selected: Array.from(selected), count: selected.size, has, toggle, toggleAll, allChecked, clear };
}
