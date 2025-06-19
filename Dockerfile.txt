<<<<<<< HEAD
# Используем официальный Python образ
FROM python:3.9-slim

# Устанавливаем рабочую директорию
WORKDIR /app

# Копируем файлы проекта в контейнер
COPY . .

# Устанавливаем зависимости
RUN pip install --upgrade pip
RUN pip install -r requirements.txt

# Запускаем бота
CMD ["python", "bot.py"]
=======
# Используем официальный Python образ
FROM python:3.9-slim

# Устанавливаем рабочую директорию
WORKDIR /app

# Копируем файлы проекта в контейнер
COPY . .

# Устанавливаем зависимости
RUN pip install --upgrade pip
RUN pip install -r requirements.txt

# Устанавливаем переменные окружения для работы с dotenv
ENV PYTHONUNBUFFERED=1
ENV BOT_TOKEN=$BOT_TOKEN
ENV MASTER_ID=$MASTER_ID
ENV MASTER_PHONE=$MASTER_PHONE
ENV SPREADSHEET_URL=$SPREADSHEET_URL
ENV SPREADSHEET_NAME=$SPREADSHEET_NAME
ENV CHANNEL_USERNAME=$CHANNEL_USERNAME
ENV CREDENTIALS_FILE=$CREDENTIALS_FILE
ENV SECRET_TOKEN=$SECRET_TOKEN
ENV RENDER_SERVICE_NAME=$RENDER_SERVICE_NAME

# Запускаем бота
CMD ["python", "bot.py"]
>>>>>>> 7e6fef3c641a57c352e42509d55be631f3e6158f
