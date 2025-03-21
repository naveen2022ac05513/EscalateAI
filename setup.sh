#!/bin/bash
pip install -r requirements.txt
python -m spacy download en_core_web_sm
git add setup.sh
git commit -m "Added setup script to fix dependencies"
git push origin main
