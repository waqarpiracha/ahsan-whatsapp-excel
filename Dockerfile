FROM python:3.12-slim

WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt gunicorn

COPY app/ ./app/
COPY server.py parse_chat.py write_workbook.py ./

# output is ephemeral on free hosting — processed per session
RUN mkdir -p output

EXPOSE 10000
CMD ["sh", "-c", "gunicorn server:app --bind 0.0.0.0:${PORT:-10000} --timeout 300 --workers 2"]
