FROM python:3.12-slim
WORKDIR /app
COPY requirements.txt requirements-excel.txt ./
RUN pip install --no-cache-dir -r requirements.txt -r requirements-excel.txt
COPY pulse ./pulse
COPY excel_engine ./excel_engine
COPY data/raw ./data/raw
ENV PYTHONUNBUFFERED=1
ENV PULSE_DATA_DIR=/home/data
CMD ["python", "-m", "excel_engine", "--from-servicebus"]
