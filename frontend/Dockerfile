FROM python:3.9 as base

# Setup env
ENV LANG C.UTF-8
ENV LC_ALL C.UTF-8
ENV PYTHONDONTWRITEBYTECODE 1
ENV PYTHONFAULTHANDLER 1

FROM base AS python-deps

# Install pipenv and compilation dependencies
RUN pip install pipenv
RUN apt-get update && apt-get install -y --no-install-recommends gcc

# Install python dependencies in /frontend/.venv
WORKDIR ./frontend
COPY frontend ./frontend
COPY Pipfile .
COPY Pipfile.lock .
RUN PIPENV_VENV_IN_PROJECT=1 pipenv install --deploy

FROM base AS runtime
EXPOSE 8501

# Copy virtual env from python-deps stage
COPY --from=python-deps /frontend/.venv /frontend/.venv
ENV PATH="/frontend/.venv/bin:$PATH"

# Install application into container
COPY frontend ./frontend
COPY VERSION .

ENTRYPOINT ["streamlit", "run"]
CMD ["frontend/frontend.py"]