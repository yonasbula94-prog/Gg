import http from 'node:http';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import express from 'express';
import { WebSocketServer } from 'ws';
import { createApiRouter, computeOrgExecutiveSnapshot } from './src/server/routes.ts';
import { streamHub } from './src/server/streamHub.ts';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

async function startServer() {
  const app = express();
  const PORT = Number(process.env.PORT || 3000);
  const isProd = process.env.NODE_ENV === 'production';

  app.set('trust proxy', 1);
  app.use(express.json({ limit: '15mb' }));

  // Mount API routes before Vite / static middleware
  app.use('/api', createApiRouter());
  app.use('/api', (_req, res) => {
    res.status(404).json({ error: 'API endpoint not found.' });
  });
  app.use(
    '/api',
    (err: unknown, _req: express.Request, res: express.Response, _next: express.NextFunction) => {
      console.error('Unhandled API error:', err);
      if (!res.headersSent) {
        res.status(500).json({ error: 'Internal server error.' });
      }
    }
  );

  const server = http.createServer(app);

  // Attach authenticated WebSocket server for bidirectional live workspace & CEO streaming
  const wss = new WebSocketServer({ server, path: '/ws' });
  wss.on('connection', (socket) => {
    let registeredClient: ReturnType<typeof streamHub.registerWsClient> | null = null;

    const authTimeout = setTimeout(() => {
      if (!registeredClient) {
        socket.close(4001, 'Authentication handshake timeout');
      }
    }, 10_000);

    socket.on('message', async (raw) => {
      try {
        const msg = JSON.parse(raw.toString('utf8'));
        if (msg?.type === 'auth' && typeof msg.ticket === 'string') {
          const ticket = streamHub.consumeStreamTicket(msg.ticket, msg.orgId);
          if (!ticket) {
            socket.close(4001, 'Invalid or expired stream ticket');
            return;
          }
          clearTimeout(authTimeout);
          registeredClient = streamHub.registerWsClient({
            orgId: ticket.orgId,
            uid: ticket.uid,
            role: ticket.role,
            socket,
          });
          const snapshot = await computeOrgExecutiveSnapshot(ticket.orgId);
          if (snapshot && socket.readyState === 1) {
            socket.send(
              JSON.stringify({
                type: 'executive:snapshot',
                orgId: ticket.orgId,
                at: new Date().toISOString(),
                payload: snapshot,
              })
            );
          }
        }
      } catch {
        // Ignore malformed frames
      }
    });

    socket.on('close', () => {
      clearTimeout(authTimeout);
      if (registeredClient) {
        streamHub.unregisterWsClient(registeredClient);
      }
    });
  });

  if (!isProd) {
    const { createServer: createViteServer } = await import('vite');
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: 'spa',
    });
    app.use(vite.middlewares);
  } else {
    const distDir = path.resolve(__dirname, 'dist');
    app.use(express.static(distDir));
    app.use((req, res, next) => {
      if (req.method === 'GET' && !req.path.startsWith('/api')) {
        res.sendFile(path.join(distDir, 'index.html'));
        return;
      }
      next();
    });
  }

  server.listen(PORT, '0.0.0.0', () => {
    console.log(`Procurement platform server listening on http://0.0.0.0:${PORT}`);
  });
}

startServer().catch((err) => {
  console.error('Fatal server startup error:', err);
  process.exit(1);
});
