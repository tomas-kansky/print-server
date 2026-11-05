FROM python:3.11-slim

ENV PYTHONDONTWRITEBYTECODE=1
ENV PYTHONUNBUFFERED=1
ENV TEMP_DIR=/tmp/print_jobs
ENV PORT=8000

RUN apt-get update && \
    apt-get install -y --no-install-recommends cups-client && \
    rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY app.py .

RUN mkdir -p /tmp/print_jobs && chmod 777 /tmp/print_jobs

EXPOSE 8000

CMD ["python", "app.py"]
