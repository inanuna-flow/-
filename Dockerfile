FROM node:20-alpine

WORKDIR /app

ENV NODE_ENV=production

COPY package*.json ./
RUN npm ci --omit=dev

COPY server.js index.html ./
COPY kpl-dashboard ./kpl-dashboard

EXPOSE 8080

CMD ["node", "server.js"]
