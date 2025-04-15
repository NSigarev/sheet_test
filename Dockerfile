# Базовый образ
FROM node:22

# Установка переменной окружения
ENV NODE_ENV=production

# Установка рабочей директории
WORKDIR /var/www/test-nest

# Копируем package.json и устанавливаем зависимости
COPY package.json package-lock.json ./
RUN npm ci

# Копируем исходники
COPY . .


# Компиляция TypeScript (если используем)
RUN npm run build

# Открываем порт
ARG APP_PORT=3000
EXPOSE ${APP_PORT}

# Запускаем сервер
ENTRYPOINT ["node", "dist/main.js"]
