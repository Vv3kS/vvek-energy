# Create a Dockerfile with the provided content

content = """FROM python:3.11

RUN apt-get update && apt-get install -y libreoffice

WORKDIR /app

COPY . .

RUN pip install -r requirements.txt

CMD ["gunicorn", "server:app", "--bind", "0.0.0.0:10000"]
"""

file_path = "/mnt/data/Dockerfile"

with open(file_path, "w") as f:
    f.write(content)

file_path