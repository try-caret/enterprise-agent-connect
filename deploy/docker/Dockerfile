# Built from the repository root as the context: docker build -f deploy/docker/Dockerfile .
FROM node:22-alpine AS build
WORKDIR /app
COPY package.json package-lock.json tsconfig.json ./
RUN npm ci
COPY packages/server/src ./packages/server/src
COPY packages/server/migrations ./packages/server/migrations
COPY packages/connector-kit/src ./packages/connector-kit/src
COPY connectors ./connectors
RUN npm run build

FROM node:22-alpine
WORKDIR /app
ENV NODE_ENV=production
COPY package.json package-lock.json ./
RUN npm ci --omit=dev
# npm is not used at runtime and its bundled tar carries CVEs.
RUN rm -rf /usr/local/lib/node_modules/npm /usr/local/bin/npm /usr/local/bin/npx
COPY --from=build /app/dist ./dist
EXPOSE 3000
USER node
CMD ["node", "dist/packages/server/src/server.js"]
