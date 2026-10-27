FROM python:3.11-slim

# Avoid prompts from apt
ENV DEBIAN_FRONTEND=noninteractive
ENV PYTHONUNBUFFERED=1

# Install LibreOffice for headless PDF conversion, essential fonts, and curl
RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice-writer \
    fonts-liberation \
    fonts-dejavu \
    fontconfig \
    curl \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# Copy dependency requirements
COPY requirements.txt /app/requirements.txt
RUN pip install --no-cache-dir -r requirements.txt

# Copy application source code
COPY bidv_engine.py /app/bidv_engine.py
COPY app.py /app/app.py

# Create outputs directory
RUN mkdir -p /app/outputs && chmod 777 /app/outputs

EXPOSE 20135

CMD ["uvicorn", "app:app", "--host", "0.0.0.0", "--port", "20135"]
