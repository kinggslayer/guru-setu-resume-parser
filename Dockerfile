FROM python:3.12-slim

# pdfplumber leans on system libraries that aren't in the slim image.
# libgl1 + libglib2.0-0 cover its image handling; curl is only for the
# container healthcheck below.
RUN apt-get update && apt-get install -y --no-install-recommends \
        libgl1 \
        libglib2.0-0 \
        curl \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# Copy requirements first so a code change doesn't invalidate the pip
# layer — rebuilds after an edit then take seconds, not minutes.
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

# Don't run as root. The app only ever writes to temp_resumes/, so that's
# the one directory that needs to be owned by the app user.
RUN useradd --create-home --uid 10001 parser \
    && mkdir -p /app/temp_resumes \
    && chown -R parser:parser /app/temp_resumes

USER parser

EXPOSE 8501

# Streamlit's own health endpoint. Docker restarts the container if the
# app wedges rather than leaving a dead page being served.
HEALTHCHECK --interval=30s --timeout=5s --start-period=30s --retries=3 \
    CMD curl --fail http://localhost:8501/_stcore/health || exit 1

# Server binding lives here, not in .streamlit/config.toml — Community
# Cloud assigns its own port, so a port pinned in that file breaks the
# free deploy. These flags only ever apply to this container.
CMD ["streamlit", "run", "app.py", \
     "--server.address=0.0.0.0", \
     "--server.port=8501", \
     "--server.headless=true"]
