FROM node:22-slim
WORKDIR /app
ENV NODE_ENV=production
COPY package.json package-lock.json ./
RUN npm ci --omit=dev
COPY *.js *.sql ./
USER node
EXPOSE 3000
CMD ["node", "index.js"]
