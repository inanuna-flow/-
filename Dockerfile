FROM node:20-alpine

# CVE fix: upgrade openssl (CVE-2026-34182 CRITICAL, and 18 other HIGH/MEDIUM/LOW openssl CVEs)
RUN apk upgrade --no-cache openssl

# CVE fix: upgrade npm to patch its bundled dependencies
# (tar, minimatch, glob, cross-spawn, brace-expansion, ip-address, @sigstore/core, diff)
RUN npm install -g npm@latest

WORKDIR /app

ENV NODE_ENV=production

COPY package*.json ./
RUN npm ci --omit=dev

COPY server.js index.html ./
COPY kpl-dashboard ./kpl-dashboard

EXPOSE 8080

CMD ["node", "server.js"]
