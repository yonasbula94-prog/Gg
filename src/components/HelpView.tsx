import React, { useMemo, useState } from 'react';
import { ArrowLeft, CheckSquare, ChevronRight, Search, X } from 'lucide-react';
import { HELP_TOPICS, HelpTopic, searchHelpTopics } from '../constants/helpTopics';

// Real screenshots bundled with the app (inlined in the offline packages). Missing files are skipped.
const SHOTS = import.meta.glob('../assets/help/*.jpg', { eager: true, query: '?url', import: 'default' }) as Record<
  string,
  string
>;

function shotUrl(name?: string): string | undefined {
  if (!name) return undefined;
  const key = Object.keys(SHOTS).find((k) => k.endsWith(`/${name}`));
  return key ? SHOTS[key] : undefined;
}

interface HelpViewProps {
  initialTopicId: string;
  returnLabel: string;
  onBack: () => void;
  accentColor: string;
  accentText: string;
}

export const HelpView: React.FC<HelpViewProps> = ({ initialTopicId, returnLabel, onBack, accentColor, accentText }) => {
  const [selectedId, setSelectedId] = useState(initialTopicId);
  const [query, setQuery] = useState('');
  const [zoom, setZoom] = useState<string | null>(null);

  const topics = useMemo(() => searchHelpTopics(query), [query]);
  const topic: HelpTopic = HELP_TOPICS.find((t) => t.id === selectedId) ?? HELP_TOPICS[0];

  return (
    <div className="space-y-4">
      <div className="flex flex-wrap items-center justify-between gap-3">
        <button
          type="button"
          onClick={onBack}
          className="inline-flex items-center gap-1.5 rounded-lg border border-slate-200 bg-white px-3.5 py-2 text-xs font-medium text-slate-700 hover:bg-slate-50 cursor-pointer min-h-[40px]"
        >
          <ArrowLeft className="w-3.5 h-3.5" />
          <span>Back to {returnLabel}</span>
        </button>
        <p className="text-xs text-slate-500">Works offline. Pick a topic, or search by what you see on screen.</p>
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-[280px_1fr] gap-4">
        {/* Topic index */}
        <aside className="rounded-lg border border-slate-200 bg-white p-3 space-y-3 h-fit">
          <div className="relative">
            <Search className="w-4 h-4 text-slate-400 absolute left-3 top-1/2 -translate-y-1/2" />
            <input
              type="search"
              value={query}
              onChange={(e) => setQuery(e.target.value)}
              placeholder="e.g. issue blocked, scanned PDF"
              className="w-full min-h-[40px] rounded-lg border border-slate-200 pl-9 pr-3 py-2 text-xs focus:border-slate-900 focus:outline-none"
            />
          </div>
          <ul className="space-y-1">
            {topics.map((t) => {
              const active = t.id === topic.id;
              return (
                <li key={t.id}>
                  <button
                    type="button"
                    onClick={() => setSelectedId(t.id)}
                    className={`w-full flex items-center justify-between gap-2 rounded-lg px-3 py-2 text-left text-xs font-medium cursor-pointer min-h-[40px] ${
                      active ? '' : 'text-slate-700 hover:bg-slate-100'
                    }`}
                    style={active ? { backgroundColor: accentColor, color: accentText } : undefined}
                  >
                    <span>{t.title}</span>
                    <ChevronRight className="w-3.5 h-3.5 shrink-0 opacity-70" />
                  </button>
                </li>
              );
            })}
            {topics.length === 0 && <li className="px-3 py-2 text-xs text-slate-500">No topic matches. Try fewer words.</li>}
          </ul>
        </aside>

        {/* Topic content */}
        <article className="rounded-lg border border-slate-200 bg-white p-5 space-y-6 min-w-0">
          <header className="space-y-1">
            <h2 className="text-base font-semibold text-slate-900">{topic.title}</h2>
            <p className="text-xs text-slate-600">{topic.summary}</p>
          </header>

          {topic.shots.map((shot) => {
            const src = shotUrl(shot.file);
            if (!src) return null;
            return (
              <figure key={shot.file} className="space-y-2">
                <button type="button" onClick={() => setZoom(src)} className="block w-full cursor-zoom-in" aria-label="Enlarge screenshot">
                  <img src={src} alt={shot.caption} className="w-full rounded-md border border-slate-200" />
                </button>
                <figcaption className="text-[11px] text-slate-500">{shot.caption} (tap to enlarge)</figcaption>
              </figure>
            );
          })}

          <section className="space-y-2">
            <h3 className="text-sm font-semibold text-slate-900">Step by step</h3>
            <ol className="space-y-2 list-decimal pl-5 text-xs text-slate-700">
              {topic.steps.map((s, i) => (
                <li key={i}>{s}</li>
              ))}
            </ol>
          </section>

          {topic.checklist && topic.checklist.length > 0 && (
            <section className="space-y-2">
              <h3 className="flex items-center gap-1.5 text-sm font-semibold text-slate-900">
                <CheckSquare className="w-4 h-4" /> Before you continue
              </h3>
              <ul className="space-y-1.5 text-xs text-slate-700">
                {topic.checklist.map((c, i) => (
                  <li key={i} className="flex gap-2">
                    <span className="text-slate-400">☐</span>
                    <span>{c}</span>
                  </li>
                ))}
              </ul>
            </section>
          )}

          <section className="space-y-2">
            <h3 className="text-sm font-semibold text-slate-900">Common problems and fixes</h3>
            <div className="space-y-3">
              {topic.problems.map((p, i) => (
                <div key={i} className="rounded-lg border border-slate-200 p-3 text-xs space-y-1">
                  <p className="font-semibold text-slate-900">{p.symptom}</p>
                  <p className="text-slate-600">
                    <span className="font-medium">Why:</span> {p.cause}
                  </p>
                  <p className="text-slate-700">
                    <span className="font-medium">Fix:</span> {p.fix}
                  </p>
                </div>
              ))}
            </div>
          </section>
        </article>
      </div>

      {zoom && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/80 p-4" onClick={() => setZoom(null)}>
          <button type="button" className="absolute top-4 right-4 rounded-full bg-white p-2 cursor-pointer" aria-label="Close">
            <X className="w-4 h-4" />
          </button>
          <img src={zoom} alt="Screenshot" className="max-h-full max-w-full rounded-md" />
        </div>
      )}
    </div>
  );
};
