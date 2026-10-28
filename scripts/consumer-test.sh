#!/usr/bin/env bash
# Installs the packed server and client into a fresh project outside the repo, type-checks
# scripts/consumer/ against them and runs it: fails when a public import or type breaks.
# Usage: DATABASE_URL=postgres://postgres@localhost:5432/ac_test scripts/consumer-test.sh
set -euo pipefail
: "${DATABASE_URL:?set DATABASE_URL to a scratch Postgres database}"
root=$(cd "$(dirname "$0")/.." && pwd)
work=$(mktemp -d)
trap 'rm -rf "$work"' EXIT

cd "$root"
npm run build >/dev/null
npx tsc -p packages/client
npm pack --silent --pack-destination "$work" >/dev/null
npm pack ./packages/client --silent --pack-destination "$work" >/dev/null

cd "$work"
cp "$root"/scripts/consumer/*.ts "$root"/scripts/consumer/tsconfig.json .
echo '{"name":"consumer","private":true,"type":"module"}' >package.json
npm install --no-audit --no-fund --silent ./arbium-enterprise-agent-connect-[0-9]*.tgz \
  ./arbium-enterprise-agent-connect-client-*.tgz \
  zod@4 postgres@3 hono@4 express@5 @types/express@5 @types/node@22 typescript@5
npx tsc -p .
node dist/main.js
