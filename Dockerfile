FROM python:3.12-slim

ENV PYTHONDONTWRITEBYTECODE=1
ENV PYTHONUNBUFFERED=1
ENV HOST=0.0.0.0
ENV PORT=8787
ENV LEAD_DATA_DIR=/data/leads
ENV LEAD_EXPORT_DIR=/data/exports

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

RUN mkdir -p /data/leads /data/exports

EXPOSE 8787

CMD ["python", "server.py"]
