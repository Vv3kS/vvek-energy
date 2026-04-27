FROM python:3.11

RUN apt-get update && apt-get install -y libreoffice

WORKDIR /app

COPY . .

RUN pip install -r requirements.txt

CMD ["gunicorn", "server:app", "--bind", "0.0.0.0:8000"]