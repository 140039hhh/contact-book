FROM python:3.12-slim
WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY server.py .
COPY static ./static
ENV HOST=0.0.0.0 PORT=8000 DATABASE_PATH=/data/contacts.db
RUN mkdir /data && useradd -m app && chown app /data
USER app
EXPOSE 8000
CMD ["python", "server.py"]
