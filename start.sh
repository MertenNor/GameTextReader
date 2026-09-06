#!/bin/bash

# This is a start script for linux systems that checks if the virtual environment is set up and runs the main.py script. If the virtual environment is not found, it will run setup_venv.sh to set it up.

if [ ! -d "pyenv" ]; then
    echo "Virtual environment not found. Setting up virtual environment..."
    ./setup_venv.sh
    exit 1
fi

source pyenv/bin/activate
python main.py
