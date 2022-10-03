FROM 3.9-alpine3.16

WORKDIR /app
COPY requirements.txt /app

RUN \
  pip install -U pip --no-cache-dir && \
  pip install -r requirements.txt --no-cache-dir

COPY . /app

RUN pip install -e . --no-cache-dir
CMD excel-comparison-bot
