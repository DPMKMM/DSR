FROM python:3.12-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    DATA_DIR=/tmp/khammam-attendance-data

WORKDIR /app

COPY requirements.txt ./
RUN pip install --no-cache-dir --upgrade pip \
    && pip install --no-cache-dir -r requirements.txt \
    && useradd --system --uid 10001 --create-home appuser \
    && mkdir -p /tmp/khammam-attendance-data \
    && chown -R appuser:appuser /app /tmp/khammam-attendance-data

COPY --chown=appuser:appuser app.py report_generator.py integrations.py ./

USER 10001
EXPOSE 8080

CMD ["sh", "-c", "exec uvicorn app:app --host 0.0.0.0 --port ${PORT:-8080} --workers 1 --proxy-headers"]
