FROM python:3.12-slim

WORKDIR /app

# No system-level deps needed (no tesseract) — Claude Vision handles OCR
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

ENV PORT=8501
EXPOSE ${PORT}

# Railway overrides via startCommand in railway.json; this is the fallback
CMD streamlit run app.py \
    --server.port=${PORT} \
    --server.address=0.0.0.0 \
    --server.headless=true \
    --browser.gatherUsageStats=false
