FROM ubuntu:24.04

ENV DEBIAN_FRONTEND=noninteractive
ENV PYTHONUNBUFFERED=1

# Install LibreOffice + Python UNO bridge + fonts + supervisor
# Accept Microsoft core fonts EULA non-interactively
RUN echo "ttf-mscorefonts-installer msttcorefonts/accepted-mscorefonts-eula select true" | debconf-set-selections

RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice-writer \
    libreoffice-calc \
    python3 \
    python3-pip \
    python3-uno \
    supervisor \
    curl \
    # Metric-compatible replacements for old MS fonts (Arial→Liberation Sans, TNR→Liberation Serif)
    fonts-liberation \
    fonts-dejavu-core \
    # Calibri-compatible (Word default since 2007 — most customer templates use this)
    fonts-crosextra-carlito \
    # Cambria-compatible (common Word heading font)
    fonts-crosextra-caladea \
    # Actual Microsoft core fonts (Arial, Verdana, Georgia, Courier New, etc.)
    ttf-mscorefonts-installer \
    fontconfig \
    && fc-cache -f -v \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/*

# Custom font support: mount a volume at /fonts with .ttf/.otf files
# They will be picked up automatically by the entrypoint font refresh

# Create working directories
WORKDIR /app
RUN mkdir -p /documents/templates /documents/output /documents/temp

# Install Python dependencies
COPY requirements.txt .
RUN pip3 install --no-cache-dir --break-system-packages -r requirements.txt

# Copy application code
COPY . .

# Copy supervisord config and entrypoint
COPY supervisord.conf /etc/supervisor/conf.d/lola.conf
COPY entrypoint.sh /entrypoint.sh
RUN chmod +x /entrypoint.sh

EXPOSE 8080

ENTRYPOINT ["/entrypoint.sh"]
