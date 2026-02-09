#!/usr/bin/env bash
# Render build script
set -o errexit

pip install --upgrade pip
pip install -r requirements.txt

# Create required directories
mkdir -p output
mkdir -p assets/ndr_cache
mkdir -p Credentials
mkdir -p templates
mkdir -p manifests
