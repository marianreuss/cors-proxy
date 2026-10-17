# Zero-dependency proxy: no npm install needed.
FROM node:22-alpine
WORKDIR /app
COPY package.json server.js ./
ENV NODE_ENV=production PORT=8080
USER node
EXPOSE 8080
CMD ["node", "server.js"]
