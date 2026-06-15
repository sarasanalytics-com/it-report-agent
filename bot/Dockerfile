# Always-on IT Slack bot. Build from the repo root:
#   docker build -f bot/Dockerfile -t it-bot .
#   docker run --env-file bot/.env it-bot
FROM python:3.12-slim

WORKDIR /app

# Install data-pipeline deps + bot deps.
COPY requirements.txt ./requirements.txt
COPY bot/requirements.txt ./bot-requirements.txt
RUN pip install --no-cache-dir -r requirements.txt -r bot-requirements.txt

# Bring in the whole repo (scripts/, prompts/, bot/).
COPY . .

# Writable runtime dirs for downloaded data + generated reports.
RUN mkdir -p data output

# app.py lives in bot/ and imports answer.py from the same dir.
CMD ["python", "bot/app.py"]
