FROM node:18-alpine

WORKDIR /app

COPY package.json package-lock.json* ./
RUN npm install --omit=dev

COPY attendance_api.js db.js env.js status_master.js ./
COPY schema ./schema
COPY public ./public

ENV NODE_ENV=production
ENV DDO_API_PORT=3000
ENV DDO_API_BIND=0.0.0.0

EXPOSE 3000

CMD ["node", "attendance_api.js"]
