FROM python:3.12-slim

RUN apt-get update && \
    apt-get install -y tesseract-ocr libtesseract-dev poppler-utils libmagic1 libmagic-dev tesseract-ocr-rus && \
    apt-get clean

COPY requirements.txt requirements.txt
RUN pip install -r requirements.txt

COPY . /app
WORKDIR /app

CMD ["python3", "scraper.py"]
