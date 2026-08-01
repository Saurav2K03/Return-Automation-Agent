# Return Automation Agent

This project is a simple Python automation tool for handling return tasks from an Excel file. It is built to help process return requests in a more organized and repeatable way.

## What this project does

- Reads return-related data from an Excel sheet
- Organizes each row as a task
- Tracks task status and return progress
- Supports simple command-line options for running the script
- Writes logs and stores screenshots when needed

## Current status

The basic structure of the project is ready. The script includes:

- task and result models
- Excel file handling setup
- logging support
- a simple CLI interface

The browser automation part is still being developed and improved.

## Requirements

This project uses Python and the following package:

- openpyxl

Install it with:

```bash
pip install openpyxl
```

## How to run

Run the script in dry-run mode to test the flow without opening a browser:

```bash
python3 return_agent.py --dry-run
```

Other useful options:

```bash
python3 return_agent.py --headless
python3 return_agent.py --order-id OD123456
```

## Project files

- return_agent.py: main Python script
- Faym Status Test Orders.xlsx: sample Excel input data
- logs/: folder for generated log files
- screenshots/: folder for screenshots if needed

## Notes

This is an early version of the project. The workflow is being built step by step, and more automation features will be added over time.
