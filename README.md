# Return Automation Agent

This project is a Python-based automation tool for processing return tasks from an Excel file. It is designed to make return handling more structured, easier to track, and simple to run from the command line.

## What this project does

The script reads pending return tasks from an Excel sheet, creates task objects for each row, and prepares the workflow for browser-based return automation. It also supports:

- reading and writing results back to Excel
- logging activity during execution
- saving screenshots when the flow hits an issue
- running in dry-run mode for testing without opening a browser

## Current status

The project is now in a working prototype stage. The core structure, Excel handling, logging, and browser automation flow are in place.

During testing, the login and order-search flow were verified, but the final return action could not be fully verified because the assessment account could not complete the required pickup flow and respond to the OTP request.

## Setup

1. Create a virtual environment

```bash
python3 -m venv .venv
source .venv/bin/activate
```

2. Install the required packages

```bash
pip install -r requirements.txt
```

3. Make sure Google Chrome or Chromium is installed on your machine.

## Run the script

Run in dry-run mode to test the flow without opening a browser:

```bash
python3 return_agent.py --dry-run
```

Run normally:

```bash
python3 return_agent.py
```

Useful options:

```bash
python3 return_agent.py --headless
python3 return_agent.py --order-id OD123456
```

## Project files

- return_agent.py: main Python script
- sql_questions.py: answers to the assignment questions
- Faym Status Test Orders.xlsx: sample Excel file used as input
- README.md: project overview and usage instructions
- requirements.txt: Python dependencies
- logs/: folder for log files
- screenshots/: folder for screenshots generated during execution

## Dependencies

The project uses:

- openpyxl
- selenium
- undetected-chromedriver

## Notes

This is a strong starting point for the project. The workflow can be improved further with better browser handling, more reliable return-step detection, and additional testing.
