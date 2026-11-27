# Cloud Run container: builds the React app and runs the Express server (server.ts),
# which serves both the API and the built frontend from dist/.
FROM node:22-slim

WORKDIR /app

COPY package.json ./
RUN npm install --no-package-lock --legacy-peer-deps --no-audit --no-fund

COPY . .
RUN npm run build

ENV NODE_ENV=production
# Cloud Run sets PORT automatically; server.ts reads it.
EXPOSE 8080
CMD ["npm", "start"]
