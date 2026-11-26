import { useCallback, useEffect, useRef, useState } from 'react';
import {
  fetchExecutiveSnapshot,
  getApiBaseUrl,
  requestStreamTicket,
} from '../lib/apiClient';
import type {
  ExecutiveSnapshot,
  StreamConnectionState,
} from '../types/enterprise';
import type { ProcurementJob, RegistryEntry } from '../types/procurement';

interface UseExecutiveStreamOptions {
  orgId: string | null;
  enabled: boolean;
  onRemoteJobUpdated?: (job: ProcurementJob, updatedByUid: string) => void;
  onRemoteJobDeleted?: (jobId: string, deletedByUid: string) => void;
  onRemoteDocIssued?: (job: ProcurementJob, registryEntry: RegistryEntry) => void;
  onRemoteRegistryVoided?: (docNumber: string, voidReason: string) => void;
  onSeatRevoked?: (reason: string) => void;
}

export function useExecutiveStream({
  orgId,
  enabled,
  onRemoteJobUpdated,
  onRemoteJobDeleted,
  onRemoteDocIssued,
  onRemoteRegistryVoided,
  onSeatRevoked,
}: UseExecutiveStreamOptions) {
  const [snapshot, setSnapshot] = useState<ExecutiveSnapshot | null>(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [streamState, setStreamState] = useState<StreamConnectionState>('offline');
  const [lastPulseAt, setLastPulseAt] = useState<string | null>(null);

  const callbacksRef = useRef({
    onRemoteJobUpdated,
    onRemoteJobDeleted,
    onRemoteDocIssued,
    onRemoteRegistryVoided,
    onSeatRevoked,
  });
  callbacksRef.current = {
    onRemoteJobUpdated,
    onRemoteJobDeleted,
    onRemoteDocIssued,
    onRemoteRegistryVoided,
    onSeatRevoked,
  };

  const refreshSnapshot = useCallback(async () => {
    if (!orgId || !enabled) return;
    setLoading(true);
    try {
      const data = await fetchExecutiveSnapshot(orgId);
      setSnapshot(data);
      setLastPulseAt(data.generatedAt);
      setError(null);
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Could not load executive snapshot.');
    } finally {
      setLoading(false);
    }
  }, [orgId, enabled]);

  useEffect(() => {
    if (!orgId || !enabled) {
      setSnapshot(null);
      setStreamState('offline');
      setError(null);
      return;
    }

    let cancelled = false;
    let source: EventSource | null = null;
    let reconnectTimer: ReturnType<typeof setTimeout> | null = null;
    let attempt = 0;

    const connect = async () => {
      if (cancelled) return;
      setStreamState(attempt === 0 ? 'connecting' : 'reconnecting');

      try {
        // Fetch initial snapshot immediately so UI renders without waiting for SSE handshake
        if (attempt === 0) {
          setLoading(true);
          try {
            const initial = await fetchExecutiveSnapshot(orgId);
            if (!cancelled) {
              setSnapshot(initial);
              setLastPulseAt(initial.generatedAt);
              setError(null);
            }
          } catch (snapErr) {
            if (!cancelled) {
              setError(
                snapErr instanceof Error ? snapErr.message : 'Could not load executive dashboard.'
              );
            }
          } finally {
            if (!cancelled) setLoading(false);
          }
        }

        const ticketRes = await requestStreamTicket(orgId);
        if (cancelled) return;

        const apiBase = getApiBaseUrl();
        const streamUrl = `${apiBase}${ticketRes.sseUrl}`;
        source = new EventSource(streamUrl);

        source.addEventListener('connected', () => {
          if (cancelled) return;
          attempt = 0;
          setStreamState('live');
          setError(null);
        });

        source.addEventListener('executive:snapshot', (ev) => {
          if (cancelled) return;
          try {
            const env = JSON.parse((ev as MessageEvent).data);
            if (env?.payload) {
              setSnapshot(env.payload as ExecutiveSnapshot);
              setLastPulseAt(env.at || new Date().toISOString());
              setStreamState('live');
              setError(null);
            }
          } catch {
            // Ignore malformed SSE frames
          }
        });

        source.addEventListener('job:updated', (ev) => {
          if (cancelled) return;
          try {
            const env = JSON.parse((ev as MessageEvent).data);
            if (env?.payload?.job) {
              callbacksRef.current.onRemoteJobUpdated?.(
                env.payload.job as ProcurementJob,
                String(env.payload.updatedByUid || '')
              );
              setLastPulseAt(env.at || new Date().toISOString());
            }
          } catch {
            // Ignore
          }
        });

        source.addEventListener('job:deleted', (ev) => {
          if (cancelled) return;
          try {
            const env = JSON.parse((ev as MessageEvent).data);
            if (env?.payload?.jobId) {
              callbacksRef.current.onRemoteJobDeleted?.(
                String(env.payload.jobId),
                String(env.payload.deletedByUid || '')
              );
            }
          } catch {
            // Ignore
          }
        });

        source.addEventListener('doc:issued', (ev) => {
          if (cancelled) return;
          try {
            const env = JSON.parse((ev as MessageEvent).data);
            if (env?.payload?.job && env?.payload?.registryEntry) {
              callbacksRef.current.onRemoteDocIssued?.(
                env.payload.job as ProcurementJob,
                env.payload.registryEntry as RegistryEntry
              );
              setLastPulseAt(env.at || new Date().toISOString());
            }
          } catch {
            // Ignore
          }
        });

        source.addEventListener('registry:updated', (ev) => {
          if (cancelled) return;
          try {
            const env = JSON.parse((ev as MessageEvent).data);
            if (env?.payload?.docNumber && env?.payload?.status === 'VOID') {
              callbacksRef.current.onRemoteRegistryVoided?.(
                String(env.payload.docNumber),
                String(env.payload.voidReason || '')
              );
            }
          } catch {
            // Ignore
          }
        });

        source.addEventListener('heartbeat', (ev) => {
          if (cancelled) return;
          try {
            const env = JSON.parse((ev as MessageEvent).data);
            setStreamState('live');
            if (typeof env?.payload?.activeListeners === 'number') {
              setSnapshot((prev) =>
                prev ? { ...prev, activeListeners: env.payload.activeListeners } : prev
              );
            }
          } catch {
            // Ignore
          }
        });

        source.addEventListener('seat:revoked', (ev) => {
          if (cancelled) return;
          try {
            const env = JSON.parse((ev as MessageEvent).data);
            const reason = String(env?.payload?.reason || 'Your seat was suspended.');
            source?.close();
            setStreamState('offline');
            callbacksRef.current.onSeatRevoked?.(reason);
          } catch {
            // Ignore
          }
        });

        source.onerror = () => {
          if (cancelled) return;
          source?.close();
          source = null;
          setStreamState(navigator.onLine ? 'reconnecting' : 'offline');
          attempt += 1;
          const delayMs = Math.min(30_000, 1500 * Math.pow(1.6, Math.min(attempt, 6)));
          reconnectTimer = setTimeout(connect, delayMs);
        };
      } catch {
        if (cancelled) return;
        setStreamState(navigator.onLine ? 'reconnecting' : 'offline');
        attempt += 1;
        const delayMs = Math.min(30_000, 2000 * Math.pow(1.6, Math.min(attempt, 6)));
        reconnectTimer = setTimeout(connect, delayMs);
      }
    };

    connect();

    const handleOnline = () => {
      if (!cancelled && !source) {
        if (reconnectTimer) clearTimeout(reconnectTimer);
        attempt = 0;
        connect();
      }
    };
    const handleOffline = () => {
      if (!cancelled) {
        setStreamState('offline');
      }
    };

    window.addEventListener('online', handleOnline);
    window.addEventListener('offline', handleOffline);

    return () => {
      cancelled = true;
      window.removeEventListener('online', handleOnline);
      window.removeEventListener('offline', handleOffline);
      if (reconnectTimer) clearTimeout(reconnectTimer);
      source?.close();
    };
  }, [orgId, enabled]);

  return {
    snapshot,
    loading,
    error,
    streamState,
    lastPulseAt,
    refreshSnapshot,
  };
}
