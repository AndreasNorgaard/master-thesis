# Master Thesis

## Setup

**Prerequisites:** [uv](https://docs.astral.sh/uv/getting-started/installation/)

```bash
# Create virtual environment and install dependencies
uv sync

# Activate the virtual environment
source .venv/bin/activate
```

## Running Models

Ensure the virtual environment is activated first (see [Setup](#setup)), then:

```bash
# Run Model 1: Profit maximization)
python -m analysis.model_1

# Run Model 2: Multi-objective (profit + CO2)
python -m analysis.model_2

# Run Model 3: Extended multi-objective (profit + CO2)
python -m analysis.model_3

# Run Scenario Analysis 1: Asset configuration
python -m analysis.scenario_1

# Run Scenario Analysis 2: Price scenarios
python -m analysis.scenario_2
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
