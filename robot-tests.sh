#!/bin/bash

if [[ "$VIRTUAL_ENV" == "" ]]; then
    source .venv/bin/activate
fi

TARGET=${1:-"."}

echo " Running tests in: tests/robot/$TARGET"
PYTHONPATH=src python -m robot -d results "tests/robot/$TARGET"