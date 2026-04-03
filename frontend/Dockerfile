# Dockerfile

FROM python:3.9 AS base

# Setup env
ENV LANG=C.UTF-8
ENV LC_ALL=C.UTF-8
ENV PYTHONDONTWRITEBYTECODE=1
ENV PYTHONFAULTHANDLER=1

FROM base AS python-deps

RUN pip install pipenv
RUN apt-get update && apt-get install -y --no-install-recommends gcc

WORKDIR /frontend
COPY frontend /frontend
COPY Pipfile .
COPY Pipfile.lock .
RUN PIPENV_VENV_IN_PROJECT=1 pipenv install --deploy

FROM base AS runtime
EXPOSE 8501

COPY --from=python-deps /frontend/.venv /frontend/.venv
ENV PATH="/frontend/.venv/bin:$PATH"

COPY frontend /frontend
COPY VERSION .

ENTRYPOINT ["streamlit", "run"]
CMD ["frontend/frontend.py"]