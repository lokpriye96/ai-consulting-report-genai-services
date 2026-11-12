#!/usr/bin/env bash
# Rebuilds data/derived.json, the deck and its PDF, then runs the tests.
set -euo pipefail
python src/derive.py > /dev/null
node src/build_deck.js
soffice --headless --convert-to pdf --outdir outputs outputs/TCS_AI_Deflation_Strategy.pptx
python -m pytest -q
