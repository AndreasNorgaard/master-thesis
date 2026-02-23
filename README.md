# Master Thesis

## Setup

**Prerequisites:** [uv](https://docs.astral.sh/uv/getting-started/installation/)

```bash
# Create virtual environment and install dependencies
uv sync

# Activate the virtual environment
source .venv/bin/activate
```

Copy `.env.example` to `.env` and fill in your API keys:

```bash
cp .env.example .env
```

## Common Commands

```bash
# Add a new package
uv add <package>

# Remove a package
uv remove <package>

# Deactivate virtual environment
deactivate
```
