# Contributing

Thanks for contributing.

## Development setup

```bash
python3 -m pip install pipenv
pipenv install --dev
```

## Common commands

```bash
make test
make lint
./config-assessment-tool.sh --help
```

You can also run tests directly:

```bash
python3 -m unittest discover -s tests -p "*.py"
```

## Coding guidelines

- Keep changes focused and small
- Preserve existing file layout and naming conventions where possible
- Update documentation when user-facing behavior changes
- Avoid committing secrets, tokens, controller URLs, or customer data

## Pull requests

Include:

- a short summary of the change
- testing notes
- any docs updates required
- screenshots only when UI behavior changed

