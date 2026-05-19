FROM python:3.11-slim

RUN apt-get update && apt-get install -y \
    libreoffice-writer \
    wget \
    unzip \
    --no-install-recommends \
    && mkdir -p /app/fonts \
    && wget -q "https://github.com/alif-type/amiri/releases/download/1.000/Amiri-1.000.zip" \
       -O /tmp/amiri.zip \
    && unzip -q /tmp/amiri.zip -d /tmp/amiri \
    && cp /tmp/amiri/Amiri-1.000/Amiri-Regular.ttf /app/fonts/ \
    && cp /tmp/amiri/Amiri-1.000/Amiri-Bold.ttf /app/fonts/ \
    && rm -rf /tmp/amiri* \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

RUN mkdir -p logs bot_data

CMD ["python", "bot.py"]
