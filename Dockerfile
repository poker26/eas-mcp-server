FROM python:3.12-slim

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY eas_client.py server.py ./

EXPOSE 8000

CMD ["python", "server.py", "--http", "--port=8000"]
