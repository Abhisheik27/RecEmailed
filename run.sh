#!/bin/bash
# -----------------------------------------------
# Quick launcher for RecEmailed
# Just run:  ./run.sh
# Or double-click from Finder
# -----------------------------------------------

# Move to the script's directory (so it works from anywhere)
cd "$(dirname "$0")"

# Activate the virtual environment
source venv/bin/activate

# Launch the Streamlit app
streamlit run app.py
