# Jira Service Management Organization Importer (CSV)

Imports Jira Service Management (JSM) Organizations from a CSV file and optionally attaches them to a Service Desk.

## Features
- CSV import (UTF-8 friendly; handles diacritics like É/ü/â)
- Create organizations (idempotent)
- Attach organizations to a Service Desk (serviceDeskId)
- Retry/backoff on transient errors (429/5xx)
- End-of-run error summary + optional error export CSV
- CI-friendly exit codes (0 success, 1 errors)

## Requirements
- PowerShell 5.1+ or PowerShell 7+
- Jira Service Management Cloud
- Atlassian API token

## CSV format
Single column CSV (default header: `Organization`).

Example: `orgs.csv`
```csv
Organization
Google
Microsoft
