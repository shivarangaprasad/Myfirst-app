FROM python:3.12-slim

WORKDIR /app

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1

RUN python -m pip install --no-cache-dir pandas openpyxl

COPY Kaggle.py BoxOfficeCollections.csv ./

CMD ["python", "Kaggle.py"]