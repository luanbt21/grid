FROM oven/bun:latest AS builder
WORKDIR /app
COPY . .
RUN bun install
RUN bun run build

FROM node:22-alpine
WORKDIR /app
COPY --from=builder /app/build build/
COPY --from=builder /app/node_modules node_modules/
COPY package.json .

EXPOSE 3000

ENV NODE_ENV=production
ENV ORIGIN=http://localhost:3000

ENV BODY_SIZE_LIMIT=10M

CMD ["node", "build"]