# Security Policy

## Supported versions

This repository is currently maintained as a research tool. Security reports should target the latest public version in the default branch.

## Reporting a vulnerability

Please do not open a public issue for vulnerabilities involving data exposure, unsafe file handling, or credential leakage. Report the problem privately to the repository maintainer.

## Local data warning

PaleoN stores uploaded files, SQLite databases, settings, dictionaries, OCR output, and exports under `paleon_data/`. This directory is intentionally ignored by Git and should not be published.
