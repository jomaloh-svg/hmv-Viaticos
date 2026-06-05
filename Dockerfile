FROM python:3.11-slim

RUN apt-get update && \
    apt-get install -y --no-install-recommends \
        libreoffice-writer \
        fonts-liberation \
        fonts-dejavu-core \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

# Render injects PORT automatically — default to 10000 if not set
ENV PORT=10000
EXPOSE 10000

CMD ["python3", "server.py"]
