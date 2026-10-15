# syntax=docker/dockerfile:1
FROM node:22-bookworm-slim AS build
WORKDIR /app
ENV ELECTRON_SKIP_BINARY_DOWNLOAD=1 \
    PLAYWRIGHT_SKIP_BROWSER_DOWNLOAD=1
RUN apt-get update \
    && apt-get install -y --no-install-recommends python3 make g++ \
    && rm -rf /var/lib/apt/lists/* \
    && corepack enable \
    && corepack prepare yarn@4.9.2 --activate

COPY package.json yarn.lock .yarnrc.yml .npmrc ./
COPY packages/common/crdt/package.json packages/common/crdt/package.json
COPY packages/common/nbstore/package.json packages/common/nbstore/package.json
COPY packages/editor/framework/package.json packages/editor/framework/package.json
COPY packages/editor/blocks/package.json packages/editor/blocks/package.json
COPY packages/frontend/core/package.json packages/frontend/core/package.json
COPY packages/frontend/apps/web/package.json packages/frontend/apps/web/package.json
COPY packages/frontend/apps/desktop/package.json packages/frontend/apps/desktop/package.json
RUN yarn install --immutable
COPY packages ./packages
COPY tsconfig.json ./tsconfig.json
COPY scripts/build-server.mjs ./scripts/build-server.mjs
RUN yarn server:build

FROM node:22-bookworm-slim AS runtime
ARG MYDOC_VERSION=dev
LABEL org.opencontainers.image.title="MyDoc Web" \
    org.opencontainers.image.version=$MYDOC_VERSION
RUN apt-get update \
    && apt-get install -y --no-install-recommends gosu ca-certificates \
    && rm -rf /var/lib/apt/lists/*
WORKDIR /app
ENV NODE_ENV=production \
    MYDOC_VERSION=$MYDOC_VERSION \
    MYDOC_CONTAINER=1 \
    MYDOC_HOST=0.0.0.0 \
    MYDOC_PORT=8080 \
    MYDOC_CONFIG_PATH=/app/data/server-config.json \
    MYDOC_CLOUD_DB_PATH=/app/data/cloud.sqlite
COPY --from=build /app/packages/frontend/apps/web/dist ./packages/frontend/apps/web/dist
COPY --from=build /app/packages/frontend/apps/web/server-dist ./packages/frontend/apps/web/server-dist
# The server bundle leaves only the native SQLite module external.
COPY --from=build /app/node_modules/better-sqlite3 ./node_modules/better-sqlite3
COPY --from=build /app/node_modules/bindings ./node_modules/bindings
COPY --from=build /app/node_modules/file-uri-to-path ./node_modules/file-uri-to-path
COPY docker/entrypoint.sh /usr/local/bin/mydoc-entrypoint
RUN chmod 755 /usr/local/bin/mydoc-entrypoint && mkdir -p /app/data && chown node:node /app/data
# Require an explicit mount; anonymous volumes are not reused by a new container.
EXPOSE 8080
HEALTHCHECK --interval=30s --timeout=5s --start-period=30s --retries=3 \
    CMD node -e "fetch('http://127.0.0.1:8080/api/hosting').then(async r=>{if(!r.ok||!(await r.json()).hosted)process.exit(1)}).catch(()=>process.exit(1))"
ENTRYPOINT ["mydoc-entrypoint"]
CMD ["node", "packages/frontend/apps/web/server-dist/main.cjs"]
