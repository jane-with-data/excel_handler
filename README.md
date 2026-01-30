# PROJECT STRUCTURE
```
project_root/
│
├── src/                          # Source code (Python package)
│   ├── __init__.py
|
│   ├── services/                 # Utility services (reusable)
│   │   ├── excel_handler/
│   │   │   ├── excel_reader.py   # Read Excel files
│   │   │   ├── excel_writer.py   # Write Excel files
│   │   │   ├── set_sheet_formatter.py   # Format Excel files
│   │   │   └── __init__.py
│   │   │
│   │   └── logger_service/
│   │       ├── logger.py         # Logging service
│   │       └── __init__.py
│   │
│   └── shared/                   # Shared utilities & constants
│       ├── configs.py            # Configuration management
│       ├── constants.py          # Application constants
│       ├── exceptions.py         # Custom exceptions
│       ├── settings.py           # (Legacy)
│       └── __init__.py
│
├── data/                         # Data directory
│   ├── input/                    # Input data
│   ├── output/                   # Output results
│   ├── temp/                     # Temporary result files
│   └── logs/                     # Application logs
│
├── docs/                         # Documentation
│   ├── project_overview.md
│
├── pyproject.toml                # Python project metadata
├── .env.example                  # Environment variables template
├── README.md                     # Project readme
```
# SETUP
1. Create env-Virtual Enviroment (Windows OS)
```
Create env: python -m venv .venv
Activate: .venv\Scripts\Activate.ps1
```
2. Install Dependencies ```pip install -e .```
...