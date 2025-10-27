# Prebuilt image with Playwright + Chromium + all system dependencies
FROM mcr.microsoft.com/playwright/python:v1.55.0-jammy

# Set working directory
WORKDIR /app

# Copy and install Python dependencies
COPY requirements.txt /app/requirements.txt
RUN pip install --no-cache-dir -r requirements.txt

# Copy your project files
COPY . /app

# Optional: set timezone to Paris
ENV TZ=Europe/Paris

# Default command (Render will override this in render.yaml)
CMD ["python", "batch_submit_summ3.py"]
