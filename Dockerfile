FROM python:3.12-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PIP_NO_CACHE_DIR=1 \
    PIP_DISABLE_PIP_VERSION_CHECK=1 \
    HOST=0.0.0.0 \
    OPEN_BROWSER=0 \
    COOKIE_SECURE=1 \
    NOMINAL_ROLLS_DIR=/app/NOMINAL ROLLS

WORKDIR /app

COPY requirements.txt ./
RUN pip install --no-cache-dir -r requirements.txt

COPY attendance-studio ./attendance-studio
COPY ["NOMINAL ROLLS", "./NOMINAL ROLLS"]

RUN useradd --create-home --uid 10001 appuser \
    && chown -R appuser:appuser /app
USER appuser

EXPOSE 8000

# Single worker on purpose: sessions and cached reports live in process memory.
CMD ["sh", "-c", "uvicorn backend.main:app --app-dir attendance-studio --host 0.0.0.0 --port ${PORT:-8000} --proxy-headers --workers 1"]
