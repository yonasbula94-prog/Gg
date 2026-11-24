import React, { useEffect, useRef, useState } from 'react';

interface HelpTipProps {
  text: string;
  align?: 'left' | 'center' | 'right';
  className?: string;
}

export const HelpTip: React.FC<HelpTipProps> = ({
  text,
  align = 'left',
  className = '',
}) => {
  const [open, setOpen] = useState(false);
  const [pinned, setPinned] = useState(false);
  const wrapperRef = useRef<HTMLSpanElement | null>(null);

  useEffect(() => {
    if (!open) return;
    const handleOutside = (e: MouseEvent | TouchEvent) => {
      if (wrapperRef.current && !wrapperRef.current.contains(e.target as Node)) {
        setOpen(false);
        setPinned(false);
      }
    };
    const handleKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') {
        setOpen(false);
        setPinned(false);
      }
    };
    document.addEventListener('mousedown', handleOutside);
    document.addEventListener('touchstart', handleOutside);
    window.addEventListener('keydown', handleKey);
    return () => {
      document.removeEventListener('mousedown', handleOutside);
      document.removeEventListener('touchstart', handleOutside);
      window.removeEventListener('keydown', handleKey);
    };
  }, [open]);

  const posClass =
    align === 'right'
      ? 'right-0'
      : align === 'center'
      ? 'left-1/2 -translate-x-1/2'
      : 'left-0';

  return (
    <span
      ref={wrapperRef}
      className={`no-print relative inline-flex items-center align-middle ${className}`}
      onMouseEnter={() => setOpen(true)}
      onMouseLeave={() => {
        if (!pinned) setOpen(false);
      }}
    >
      <button
        type="button"
        aria-label="Help information"
        aria-expanded={open}
        onClick={(e) => {
          e.stopPropagation();
          if (pinned) {
            setPinned(false);
            setOpen(false);
          } else {
            setPinned(true);
            setOpen(true);
          }
        }}
        className={`inline-flex items-center justify-center w-4 h-4 rounded-full border text-[10px] font-bold leading-none transition-colors cursor-pointer select-none ${
          open
            ? 'border-slate-900 bg-slate-900 text-white'
            : 'border-slate-300 bg-white text-slate-500 hover:border-slate-700 hover:text-slate-900'
        }`}
      >
        ?
      </button>

      {open && (
        <span
          role="tooltip"
          onClick={(e) => e.stopPropagation()}
          className={`absolute top-full mt-1.5 z-50 w-64 sm:w-72 rounded-lg border border-slate-200 bg-slate-900 px-3 py-2 text-[11px] font-normal leading-relaxed text-white shadow-xl whitespace-normal text-left ${posClass}`}
        >
          {text}
        </span>
      )}
    </span>
  );
};
