FROM python:3.12-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PIP_NO_CACHE_DIR=1 \
    PIP_DISABLE_PIP_VERSION_CHECK=1

# A Unicode font is required for Arabic/non-Latin text in generated PDFs.
RUN apt-get update \
 && apt-get install -y --no-install-recommends fonts-dejavu-core \
 && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# Dependencies first so this layer is cached until requirements.txt changes.
COPY requirements.txt .
RUN pip install -r requirements.txt

# Run as an unprivileged user.
RUN useradd --system --create-home --uid 10001 appuser
COPY --chown=appuser:appuser . .
USER appuser

ENV PORT=8080
EXPOSE 8080

# Shell form so ${PORT} is expanded; `exec` makes uvicorn PID 1 so it receives SIGTERM.
CMD exec uvicorn main:app --host 0.0.0.0 --port ${PORT} --no-server-header --timeout-keep-alive 75