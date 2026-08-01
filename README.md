# Return Automation Agent

This project is a Python-based automation tool for handling return tasks from an Excel file. It is designed to make return processing more structured, easier to track, and simpler to run from the command line.

## Overview

The script reads return-related data from an Excel sheet, prepares each row as a task, and provides a foundation for automating return actions. It also includes support for logging, status tracking, and basic command-line options.

## What the project includes

- A simple task model for each return request
- A result model to track what happened after a task is processed
- Excel column mapping for the input sheet
- Logging support for debugging and monitoring
- A command-line interface for running the script

## Current status

The project is still in an early stage. The basic structure and workflow are in place, while the browser automation part is being developed further.

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

- [return_agent.py](return_agent.py): main Python script
- [Faym Status Test Orders.xlsx](Faym%20Status%20Test%20Orders.xlsx): sample Excel input file
- logs/: folder for generated log files
- screenshots/: folder for screenshots if needed

## Notes

This is a working foundation for the automation project. More features will be added as the workflow becomes more complete.
