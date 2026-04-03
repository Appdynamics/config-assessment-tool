# Dockerfile

FROM python:3.12 AS base

# Setup env
ENV LANG=C.UTF-8
ENV LC_ALL=C.UTF-8
ENV PYTHONDONTWRITEBYTECODE=1
ENV PYTHONFAULTHANDLER=1

FROM base AS python-deps

RUN pip install pipenv
RUN apt-get update && apt-get install -y --no-install-recommends gcc

WORKDIR ./backend
COPY backend ./backend
COPY Pipfile .
COPY Pipfile.lock .
RUN PIPENV_VENV_IN_PROJECT=1 pipenv install --deploy

FROM base AS runtime

# Copy virtual env from python-deps stage
COPY --from=python-deps /backend/.venv /backend/.venv
ENV PATH="/backend/.venv/bin:$PATH"

## Install application into container
COPY backend ./backend
COPY VERSION .

ENTRYPOINT [ "python", "backend/backend.py" ]