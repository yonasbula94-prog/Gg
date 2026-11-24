import { useCallback, useEffect, useRef, useState } from 'react';

export interface CellRangeConfig {
  /** Row numbers in table order */
  rowOrder: number[];
  /** Editable field for a table column (cell index in the row), or null when the column cannot be copied or pasted */
  fieldAt: (cellIndex: number) => string | null;
  /** Current display text of one cell (what Ctrl+C copies) */
  readCell: (rowNumber: number, field: string) => string;
  /** Writes one cell from pasted text (the view parses numbers and text) */
  writeCell: (rowNumber: number, field: string, text: string) => void;
  /** Table element attribute value, used to highlight the selected cells */
  tableKey: string;
}

interface Pos {
  row: number;
  col: number;
}

/**
 * Free cell-range selection for editable tables: drag across rows and columns, Ctrl+C copies the block
 * as tab-separated text, Ctrl+V pastes it starting at the top-left cell of the selection.
 * Single-cell clicks behave as normal typing; the range only takes over when it covers more than one cell.
 */
export function useCellRange(config: CellRangeConfig) {
  const [anchor, setAnchor] = useState<Pos | null>(null);
  const [focus, setFocus] = useState<Pos | null>(null);
  const dragging = useRef(false);
  const configRef = useRef(config);
  configRef.current = config;
  const memoryClip = useRef<string>('');
  const clipboardWriteFailed = useRef(false);

  const bounds = (a: Pos | null, b: Pos | null) => {
    if (!a || !b) return null;
    return {
      r0: Math.min(a.row, b.row),
      r1: Math.max(a.row, b.row),
      c0: Math.min(a.col, b.col),
      c1: Math.max(a.col, b.col),
    };
  };
  const box = bounds(anchor, focus);
  const cellCount = box ? (box.r1 - box.r0 + 1) * (box.c1 - box.c0 + 1) : 0;

  /** Row positions are indexes into rowOrder, so the selection is independent of row numbers */
  const pos = (rowNumber: number, cellIndex: number): Pos => ({
    row: configRef.current.rowOrder.indexOf(rowNumber),
    col: cellIndex,
  });

  const onCellDown = useCallback((rowNumber: number, cellIndex: number, e: React.MouseEvent) => {
    if (e.button !== 0 || cellIndex < 2) return;
    dragging.current = true;
    const p = pos(rowNumber, cellIndex);
    setAnchor(p);
    setFocus(p);
  }, []);

  const onCellOver = useCallback((rowNumber: number, cellIndex: number) => {
    if (!dragging.current) return;
    setFocus(pos(rowNumber, cellIndex));
  }, []);

  useEffect(() => {
    const up = () => {
      dragging.current = false;
    };
    window.addEventListener('mouseup', up);
    return () => window.removeEventListener('mouseup', up);
  }, []);

  /** Cells of the current range, as text rows */
  const selectedText = (): string => {
    const c = configRef.current;
    if (!box) return '';
    const lines: string[] = [];
    for (let r = box.r0; r <= box.r1; r++) {
      const rn = c.rowOrder[r];
      const cells: string[] = [];
      for (let col = box.c0; col <= box.c1; col++) {
        const field = c.fieldAt(col);
        cells.push(field && rn !== undefined ? c.readCell(rn, field) : '');
      }
      lines.push(cells.join('\t'));
    }
    return lines.join('\n');
  };

  const copy = useCallback(async () => {
    const text = selectedText();
    memoryClip.current = text;
    try {
      await navigator.clipboard.writeText(text);
      clipboardWriteFailed.current = false;
    } catch {
      // Clipboard blocked: the in-app copy is used for paste, so unrelated clipboard text is never pasted
      clipboardWriteFailed.current = true;
    }
    return cellCount;
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [cellCount, box?.r0, box?.r1, box?.c0, box?.c1]);

  const paste = useCallback(async () => {
    if (!box) return 0;
    const c = configRef.current;
    let text = memoryClip.current;
    if (!clipboardWriteFailed.current) {
      try {
        const fromClipboard = await navigator.clipboard.readText();
        if (fromClipboard) text = fromClipboard;
      } catch {
        // Clipboard read blocked: use the copy made on this page
      }
    }
    if (!text) return 0;
    const lines = text.replace(/\r/g, '').split('\n');
    if (lines.length > 0 && lines[lines.length - 1] === '') lines.pop();
    let written = 0;
    lines.forEach((line, i) => {
      const rn = c.rowOrder[box.r0 + i];
      if (rn === undefined) return;
      line.split('\t').forEach((value, j) => {
        const field = c.fieldAt(box.c0 + j);
        if (!field || value === '') return;
        c.writeCell(rn, field, value.trim());
        written += 1;
      });
    });
    return written;
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [box?.r0, box?.r1, box?.c0, box?.c1]);

  // Ctrl+C / Ctrl+V take over only for a multi-cell range
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (!(e.ctrlKey || e.metaKey) || cellCount < 2) return;
      // Typing in a field outside the table keeps normal copy and paste
      const target = e.target as HTMLElement | null;
      const inTable = !!target?.closest?.('table[data-range-table]');
      if (!inTable && target && /^(INPUT|TEXTAREA|SELECT)$/.test(target.tagName)) return;
      const key = e.key.toLowerCase();
      if (key === 'c') {
        e.preventDefault();
        copy();
      } else if (key === 'v') {
        e.preventDefault();
        paste();
      }
    };
    window.addEventListener('keydown', onKey);
    return () => window.removeEventListener('keydown', onKey);
  }, [cellCount, copy, paste]);

  // Highlight the selected cells (toggled on the DOM so table markup stays unchanged)
  useEffect(() => {
    const rows = document.querySelectorAll<HTMLTableRowElement>(`table[data-range-table="${config.tableKey}"] tbody tr`);
    rows.forEach((tr, r) => {
      Array.from(tr.cells).forEach((td, col) => {
        const on = box !== null && r >= box.r0 && r <= box.r1 && col >= box.c0 && col <= box.c1 && col >= 2;
        td.classList.toggle('cell-range-on', on);
      });
    });
  });

  return {
    cellCount,
    onCellDown,
    onCellOver,
    copy,
    paste,
    clear: () => {
      setAnchor(null);
      setFocus(null);
    },
  };
}
