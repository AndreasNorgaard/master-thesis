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

## Running Models

Ensure the virtual environment is activated first (see [Setup](#setup)), then:

```bash
# Run Model 1 (profit maximization)
python -m models.model_1

# Run Model 2 (multi-objective: profit + CO2)
python -m models.model_2
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
