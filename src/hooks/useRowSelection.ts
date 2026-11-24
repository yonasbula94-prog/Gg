import { useCallback, useEffect, useRef, useState } from 'react';

interface DragState<K> {
  anchor: K;
  select: boolean;
  baseline: Set<K>;
  order: K[];
  moved: boolean;
}

/**
 * Multi-row selection for tables. Rows are identified by a stable key (e.g. rowNumber).
 * - Click a checkbox to toggle one row.
 * - Shift-click selects the range from the last clicked row.
 * - Press on a checkbox and drag across rows to select (or clear) a band of rows.
 * Every row can be selected, including empty rows.
 */
export function useRowSelection<K extends string | number = number>() {
  const [selected, setSelected] = useState<Set<K>>(new Set());
  const anchorRef = useRef<K | null>(null);
  const dragRef = useRef<DragState<K> | null>(null);

  const has = useCallback((key: K) => selected.has(key), [selected]);

  const rangeBetween = (order: K[], from: K, to: K): K[] => {
    const a = order.indexOf(from);
    const b = order.indexOf(to);
    if (a < 0 || b < 0) return [to];
    const lo = Math.min(a, b);
    const hi = Math.max(a, b);
    return order.slice(lo, hi + 1);
  };

  /** Checkbox click. `order` lists the row keys in table order (needed for shift-click). */
  const click = useCallback((key: K, order: K[], shift: boolean) => {
    const anchor = anchorRef.current;
    if (shift && anchor !== null) {
      setSelected(new Set(rangeBetween(order, anchor, key)));
      return;
    }
    anchorRef.current = key;
    setSelected((prev) => {
      const next = new Set(prev);
      if (next.has(key)) next.delete(key);
      else next.add(key);
      return next;
    });
  }, []);

  const toggle = useCallback((key: K) => click(key, [], false), [click]);

  /** Mouse down on a checkbox cell: starts a drag band. */
  const beginDrag = useCallback(
    (key: K, order: K[]) => {
      dragRef.current = {
        anchor: key,
        select: !selected.has(key),
        baseline: new Set(selected),
        order,
        moved: false,
      };
    },
    [selected]
  );

  /** Pointer enters a row while dragging: extends the band from the anchor row. */
  const dragOver = useCallback((key: K) => {
    const drag = dragRef.current;
    if (!drag) return;
    if (key !== drag.anchor) drag.moved = true;
    const band = rangeBetween(drag.order, drag.anchor, key);
    const next = new Set(drag.baseline);
    band.forEach((k) => (drag.select ? next.add(k) : next.delete(k)));
    setSelected(next);
  }, []);

  useEffect(() => {
    const end = () => {
      const drag = dragRef.current;
      if (drag && drag.moved) anchorRef.current = drag.anchor;
      dragRef.current = null;
    };
    window.addEventListener('mouseup', end);
    return () => window.removeEventListener('mouseup', end);
  }, []);

  const allChecked = (keys: K[]) => keys.length > 0 && keys.every((k) => selected.has(k));

  const toggleAll = useCallback((keys: K[]) => {
    setSelected((prev) => {
      const everyChecked = keys.length > 0 && keys.every((k) => prev.has(k));
      return everyChecked ? new Set() : new Set(keys);
    });
  }, []);

  const clear = useCallback(() => {
    anchorRef.current = null;
    setSelected(new Set());
  }, []);

  return {
    selected: Array.from(selected),
    count: selected.size,
    has,
    click,
    toggle,
    beginDrag,
    dragOver,
    toggleAll,
    allChecked,
    clear,
  };
}
