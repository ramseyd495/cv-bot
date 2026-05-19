FROM python:3.11-slim

# Install LibreOffice for PDF conversion (minimal)
RUN apt-get update && apt-get install -y \
    libreoffice-writer \
    --no-install-recommends \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/*

# Set working directory
WORKDIR /app

# Copy and install dependencies first (for Docker cache)
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy project files
COPY . .

# Create directory for logs and bot data
RUN mkdir -p logs bot_data

# Run the bot
CMD ["python", "bot.py"]
