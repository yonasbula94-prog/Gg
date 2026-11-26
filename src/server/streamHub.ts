import crypto from 'node:crypto';
import type { Response } from 'express';
import type { WebSocket } from 'ws';
import type { OrgRole, StreamEnvelope } from '../types/enterprise.ts';

interface StreamTicket {
  ticket: string;
  orgId: string;
  uid: string;
  role: OrgRole;
  displayName: string;
  expiresAt: number;
}

interface SseClient {
  id: string;
  orgId: string;
  uid: string;
  role: OrgRole;
  res: Response;
}

interface WsClient {
  id: string;
  orgId: string;
  uid: string;
  role: OrgRole;
  socket: WebSocket;
}

class OrgStreamHub {
  private tickets = new Map<string, StreamTicket>();
  private sseClients = new Map<string, Set<SseClient>>();
  private wsClients = new Map<string, Set<WsClient>>();
  private heartbeatTimer: ReturnType<typeof setInterval> | null = null;

  constructor() {
    this.heartbeatTimer = setInterval(() => {
      this.cleanupExpiredTickets();
      this.broadcastHeartbeats();
    }, 25000);
    if (this.heartbeatTimer.unref) {
      this.heartbeatTimer.unref();
    }
  }

  private cleanupExpiredTickets() {
    const now = Date.now();
    for (const [key, item] of this.tickets.entries()) {
      if (item.expiresAt <= now) {
        this.tickets.delete(key);
      }
    }
  }

  createStreamTicket(params: {
    orgId: string;
    uid: string;
    role: OrgRole;
    displayName: string;
  }): { ticket: string; expiresInSeconds: number } {
    this.cleanupExpiredTickets();
    const ticket = crypto.randomBytes(24).toString('base64url');
    this.tickets.set(ticket, {
      ticket,
      orgId: params.orgId,
      uid: params.uid,
      role: params.role,
      displayName: params.displayName,
      expiresAt: Date.now() + 60_000,
    });
    return { ticket, expiresInSeconds: 60 };
  }

  consumeStreamTicket(ticket: string, expectedOrgId?: string): StreamTicket | null {
    const entry = this.tickets.get(ticket);
    if (!entry) return null;
    this.tickets.delete(ticket);
    if (entry.expiresAt <= Date.now()) return null;
    if (expectedOrgId && entry.orgId !== expectedOrgId) return null;
    return entry;
  }

  registerSseClient(params: {
    orgId: string;
    uid: string;
    role: OrgRole;
    res: Response;
  }): SseClient {
    const client: SseClient = {
      id: crypto.randomUUID(),
      orgId: params.orgId,
      uid: params.uid,
      role: params.role,
      res: params.res,
    };
    let room = this.sseClients.get(params.orgId);
    if (!room) {
      room = new Set();
      this.sseClients.set(params.orgId, room);
    }
    room.add(client);

    this.writeSse(client, {
      type: 'connected',
      orgId: params.orgId,
      at: new Date().toISOString(),
      payload: {
        clientId: client.id,
        activeListeners: this.getListenerCount(params.orgId),
      },
    });

    return client;
  }

  unregisterSseClient(client: SseClient) {
    const room = this.sseClients.get(client.orgId);
    if (!room) return;
    room.delete(client);
    if (room.size === 0) {
      this.sseClients.delete(client.orgId);
    }
  }

  registerWsClient(params: {
    orgId: string;
    uid: string;
    role: OrgRole;
    socket: WebSocket;
  }): WsClient {
    const client: WsClient = {
      id: crypto.randomUUID(),
      orgId: params.orgId,
      uid: params.uid,
      role: params.role,
      socket: params.socket,
    };
    let room = this.wsClients.get(params.orgId);
    if (!room) {
      room = new Set();
      this.wsClients.set(params.orgId, room);
    }
    room.add(client);

    this.writeWs(client, {
      type: 'connected',
      orgId: params.orgId,
      at: new Date().toISOString(),
      payload: {
        clientId: client.id,
        activeListeners: this.getListenerCount(params.orgId),
      },
    });

    return client;
  }

  unregisterWsClient(client: WsClient) {
    const room = this.wsClients.get(client.orgId);
    if (!room) return;
    room.delete(client);
    if (room.size === 0) {
      this.wsClients.delete(client.orgId);
    }
  }

  getListenerCount(orgId: string): number {
    const sse = this.sseClients.get(orgId)?.size ?? 0;
    const ws = this.wsClients.get(orgId)?.size ?? 0;
    return sse + ws;
  }

  broadcastToOrg(orgId: string, type: StreamEnvelope['type'], payload: unknown) {
    const envelope: StreamEnvelope = {
      type,
      orgId,
      at: new Date().toISOString(),
      payload,
    };

    const sseRoom = this.sseClients.get(orgId);
    if (sseRoom) {
      for (const client of sseRoom) {
        this.writeSse(client, envelope);
      }
    }

    const wsRoom = this.wsClients.get(orgId);
    if (wsRoom) {
      for (const client of wsRoom) {
        this.writeWs(client, envelope);
      }
    }
  }

  revokeUserStreams(orgId: string, uid: string, reason: string) {
    const envelope: StreamEnvelope = {
      type: 'seat:revoked',
      orgId,
      at: new Date().toISOString(),
      payload: { reason },
    };

    const sseRoom = this.sseClients.get(orgId);
    if (sseRoom) {
      for (const client of Array.from(sseRoom)) {
        if (client.uid === uid) {
          this.writeSse(client, envelope);
          try {
            client.res.end();
          } catch {
            // Ignore already-closed responses
          }
          this.unregisterSseClient(client);
        }
      }
    }

    const wsRoom = this.wsClients.get(orgId);
    if (wsRoom) {
      for (const client of Array.from(wsRoom)) {
        if (client.uid === uid) {
          this.writeWs(client, envelope);
          try {
            client.socket.close(4003, reason);
          } catch {
            // Ignore already-closed sockets
          }
          this.unregisterWsClient(client);
        }
      }
    }
  }

  private writeSse(client: SseClient, envelope: StreamEnvelope) {
    try {
      client.res.write(`event: ${envelope.type}\ndata: ${JSON.stringify(envelope)}\n\n`);
    } catch {
      this.unregisterSseClient(client);
    }
  }

  private writeWs(client: WsClient, envelope: StreamEnvelope) {
    try {
      if (client.socket.readyState === 1) {
        client.socket.send(JSON.stringify(envelope));
      }
    } catch {
      this.unregisterWsClient(client);
    }
  }

  private broadcastHeartbeats() {
    const nowIso = new Date().toISOString();
    for (const [orgId, room] of this.sseClients.entries()) {
      for (const client of room) {
        this.writeSse(client, {
          type: 'heartbeat',
          orgId,
          at: nowIso,
          payload: { activeListeners: this.getListenerCount(orgId) },
        });
      }
    }
    for (const [orgId, room] of this.wsClients.entries()) {
      for (const client of room) {
        this.writeWs(client, {
          type: 'heartbeat',
          orgId,
          at: nowIso,
          payload: { activeListeners: this.getListenerCount(orgId) },
        });
      }
    }
  }
}

export const streamHub = new OrgStreamHub();
