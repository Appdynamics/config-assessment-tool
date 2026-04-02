# Dockerfile

FROM python:3.12 AS base

LABEL org.opencontainers.image.description "Configuration Assessment Tool (CAT) is a comprehensive solution designed to evaluate and enhance the security of your software configurations. By analyzing configuration files, CAT identifies potential vulnerabilities and provides actionable recommendations to mitigate risks. With support for various configuration formats and integration capabilities, CAT empowers organizations to maintain secure and compliant configurations across their software ecosystems."

ENV LANG=C.UTF-8
ENV LC_ALL=C.UTF-8
ENV PYTHONDONTWRITEBYTECODE=1
ENV PYTHONFAULTHANDLER=1

FROM base AS python-deps

RUN pip install pipenv
RUN apt-get update && apt-get install -y --no-install-recommends gcc

WORKDIR /app

COPY backend /app/backend
COPY frontend /app/frontend
COPY Pipfile .
COPY Pipfile.lock .

RUN PIPENV_VENV_IN_PROJECT=1 pipenv install --deploy

FROM base AS runtime

ENV PYTHONPATH="/app:/app/backend"
WORKDIR /app

COPY --from=python-deps /app/.venv /app/.venv
ENV PATH="/app/.venv/bin:$PATH"

COPY backend /app/backend
COPY frontend /app/frontend
COPY plugins /app/plugins
COPY bin /app/bin
COPY VERSION .
COPY entrypoint.sh /app/entrypoint.sh

RUN chmod +x /app/entrypoint.sh

EXPOSE 8501

ENTRYPOINT ["/app/entrypoint.sh"]
CMD []