# File path manipulation, directory creation, env access
import os

# System-level access, here to add stdout as logging handler
import sys

# Provides time.sleep() for delays -> bot-detection avoidance
import time

# Random numbers for variable delays using random.uniform(min, max)
import random

# For logging. Better than print here b'coz it supports:
#   – log levels (DEBUG, INFO, WARNING, ERROR)
#   – simultaneous o/p to file & console
#   – timestamps
#   – named loggers for filtering
import logging

# For generating ISO 8601 timestamps
from datetime import datetime

# Decorator that auto-generates __init__, __repr__, __eq__ methods for data-holding classes.
# Reduces boilerplate significantly. 'field' allows setting default values for mutable types.
from dataclasses import dataclass, field

# Defines a set of named constants. Used for return status values to prevent typos and enable IDE autocomplete
from enum import Enum

# Library for reading/writing excel (.xlsx) files.
# Chosen over pandas b'coz:
#   – openpyxl preserves existing formatting
#   – allows cell-level writes
#   – safer for concurrent access (read-modify-write pattern)
import openpyxl

import argparse


# –––––––––––––
# CONFIGURATION
# –––––––––––––
# All config is defined as module-level constants.
# This makes it easy to find and modify settings without searching through the code.

# Path to the excel file containing return tasks.
# os.path.abspath(__file__) gets the absolute path to the script itself
# os.path.dirname() strips the filename, leaving the directory path
# os.path.join() constructs the full path in an OS-independent way
EXCEL_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'Faym Status Test Orders.xlsx')

# Directory where screenshots are saved when the agent encounters issues.
SCREENSHOTS_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'screenshots')

# Directory for log files. Each run creates timestamped log file.
LOG_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'logs')

FLIPKART_LOGIN_URL = "https://www.flipkart.com/account/login"
FLIPKART_ORDERS_URL = "https://www.flipkart.com/account/orders"

# Test Flipkart Account
FLIPKART_PHONE = "9205359199"

# Bot-detection avoidance timing config
MIN_ACTION_DELAY = 1.0
MAX_ACTION_DELAY = 3.5

MIN_TYPING_DELAY = 0.5
MAX_TYPING_DELAY = 0.15

# How long to wait for a page to load before timing out.
PAGE_LOAD_TIMEOUT = 30

MAX_RETRIES = 3

# Column mappings for the Excel file (1-indexed to match openpyxl convention).
# This class documents which column contains which data, preventing magic numbers scattered throughout the code.
# If the Excel structure changes, only this class needs to be updated.
class ExcelColumns:
    ADDRESS = 1         # Column A: Shipping address
    CONTACT = 2         # Column B: Contact phone number
    PRODUCT_LINK = 3    # Column C: Product URL(s) on Flipkart
    AMOUNT = 4          # Column D: Total order amount in ₹
    NUM_PRODUCTS = 5    # Column E: Number of products in the order
    ORDER_DATE = 6      # Column F: Date the order was placed
    ORDER_ID = 7        # Column G: Flipkart order ID (e.g., OD337915012166989100)
    DELIVERY_DATE = 8   # Column H: Expected/actual delivery date
    RETURN_WINDOW = 9   # Column I: Return eligibility window (e.g., "10 Days")
    STATUS = 10         # Column J: Task status (Pending/Done/Needs Review)
    PLATFORM = 11       # Column K: E-commerce platform (Flipkart/Amazon)
    REFUND_ID = 12      # Column L: Refund/Return ID assigned by the platform
    RETURN_STATUS = 13  # Column M: Detailed return status
    REFUND_AMOUNT = 14  # Column N: Amount refunded by the platform
    TIMESTAMP = 15      # Column O: When the agent last processed this row
    LOG = 16            # Column P: Detailed log/error message


# –––––––––––
# DATA MODELS
# –––––––––––
# These classes define the data structures used throughout the agent.
# Using @dataclass and Enum provides type safety and self-documentation.

class ReturnStatus(Enum):
    """
    Enumeration of all possible return statuses.
    Using an Enum prevents typos (e.g. "Pendingg") that would cause bugs,
    and enables IDE autocomplete
    """
    PENDING = "Pending"                           # Not yet processed
    PLACED = "Return Placed"                      # Return successfully initiated
    FAILED = "Failed"                             # Return attempt failed
    OUT_OF_WINDOW = "Out of Return Window"        # Past the return deadline
    NOT_DELIVERED = "Not yet delivered"            # Order hasn't arrived yet
    ALREADY_CANCELLED = "Already Cancelled & Refunded"  # Already handled
    NEEDS_REVIEW = "Support Needed"               # Requires manual intervention
    DONE = "Done"

@dataclass
class ReturnTask:
    """
    Represents a single return task (one line item) read from the Excel file.
    
    Each row in the Excel sheet becomes one ReturnTask instance.
    The row_number field is crucial - it tells us which excel row to write
    results back to after processing.
    
    Fields marked 'datatype | None' can be None when the column is empty in excel.
    This happens for new/unprocessed tasks where result columns haven't been filled yet.
    """
    row_number: int                        # Excel row number (2-indexed, since row 1 is headers)
    address: str                           # Shipping address of the customer
    contact: str                           # Contact phone number
    product_link: str                      # Flipkart product URL(s)
    amount: float                          # Order amount in ₹
    num_products: int                      # Number of products in the order
    order_date: str                        # When the order was placed
    order_id: str                          # Flipkart order ID (unique identifier)
    delivery_date: str                     # When the order was delivered
    return_window: str                     # Return eligibility window (e.g., "10 Days")
    status: str                            # Current task status
    platform: str                          # E-commerce platform (e.g., "Flipkart")
    refund_id: str | None = None        # Refund ID (filled by agent or pre-existing)
    return_status: str | None = None    # Detailed return status
    refund_amount: float | None = None  # Refund amount in ₹
    timestamp: str | None = None        # Last processing timestamp
    log: str | None = None              # Log/error message from previous runs


@dataclass
class ReturnResult:
    """
    The outcome of processing a single return task.
    
    This is what the agent produces after attempting a return.
    It's separate from ReturnTask because a task (input) and its result (output)
    are conceptually different - the task describes WHAT to do, the result
    describes WHAT HAPPENDED.
    """
    success: bool                           # True if return was successfully placed
    return_id: str | None = None            # Return/Refund ID from the platform
    return_status: str = ""                 # Human-readable status description
    refund_amount: float | None = None      # Refund amount in ₹ (from platform)
    log_message: str = ""                   # Detailed log for debugging
    screenshot_path: str | None = None      # Path to screenshot (for failed cases)


# –––––––––––––
# LOGGING SETUP
# –––––––––––––

def setup_logging():
    """
    Configure the logging system with dula ouput (file + console).
    
    Each run creates a new timestamped log file in the logs/ directory.
    This ensures:
        1. Console output for real-time monitoring during execution
        2. Persistent log files for post-mortem debugging
        3. Consistent timestamp format across all log entries
    
    Returns:
        logging.Logger: Configured logger instance named 'ReturnAgent'
    """
    os.makedirs(LOG_DIR, exist_ok=True)

    log_file = os.path.join(LOG_DIR, f"return_agent_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log")

    # Configure root logger with:
    #   - INFO level: logs INFO, WARNING, ERROR, CRITICAL (not DEBUG)
    #   - Format: timestamp [LEVEL] logger_name: message
    #   - Two handlers: file (persistent) and stdout (real-time)
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s [%(levelname)s] %(name)s: %(message)s',
        handlers=[
            logging.FileHandler(log_file),
            logging.StreamHandler(sys.stdout)
        ]
    )
    return logging.getLogger('ReturnAgent')


# ------------------------
# HUMAN BEHAVIOR SIMULATOR
# ------------------------

class HumanSimulator:
    """
    Simulates human-like behavior to avoid bot detection by e-commerce platforms.

    E-commerce sites like Flipkart use sophisticated bot detection systems
    (e.g., PerimeterX, Akamai Bot Manager, DataDome) that analyze:
      - Click timing patterns (bots click at machine speed)
      - Mouse movement trajectories (bots move in straight lines)
      - Typing speed (bots paste entire strings instantly)
      - Scroll behavior (bots jump directly to elements)
      - Session patterns (bots don't pause to "think")
    
    This class implements five strategies to mimic human behavior:
      1. Random delays between actions (1-3.5 seconds)
      2. Character-by-character typing with variable speed
      3. Natural scroll patterns (small increments, not instant jumps)
      4. Random mouse movements to avoid straight-line patterns
      5. Session-level rate limiting (longer pauses every N actions)
    """
    # TODO
    pass


# -------------
# EXCEL MANAGER
# -------------

def _safe_float(value):
    if value is None:
        return None
    try:
        return float(value)
    except (ValueError, TypeError):
        return None
    

class ExcelManager:
    """
    Handles reading pending tasks from and writing results to the Excel file.
    
    Design decisions:
    1. Uses openpyxl (not pandas) because openpyxl supports cell-level writes
       without rewriting the entire sheet. Pandas would destroy formatting.
    2. Opens and closes the workbook for each operation to prevent file
       corruption if the agent crashes. This is slower but much safer.
    3. Results are written immediately after each task completes, so partial
       progress is preserved even if the agent crashes mid-execution.
    """

    def __init__(self, filepath, logger):
        """
        Args:
        filepath: Absolute path to the Excel file
        logger: Logger instance for recording read/write operations
        """
        self.filepath = filepath
        self.logger = logger

    def read_pending_tasks(self):
        """
        Read all rows with status 'Pending' from the excel file.
        
        Workflow:
        1. Open the workbook in read mode
        2. Iterate through all rows starting from row 2 (row 1 = headers)
        3. Filter rows where the Status column (J) is "Pending"
        4. Convert each matching row into a ReturnTask dataclass instance
        5. Close the workbook and return the list of tasks
        
        Returns:
            list[ReturnTask]: List of pending tasks to process
        """
        # Opent the workbook.
        wb = openpyxl.load_workbook(self.filepath)
        ws = wb.active      # Get's the first active sheet (we only have one sheet)
        tasks = []

        # Iterate through rows starting from row 2 (skip header row 1).
        # enumerate(..., start=2) makes row_idx match the actual Excel row number,
        # which we need later for writing results back to the correct row.
        # values_only=False returns Cell objects (not just values) — we use this
        # to access cell.value for each column.
        for row_idx, row in enumerate(ws.iter_rows(min_row=2, values_only=False), start=2):
            # Extract the value from each cell in the row
            values = [cell.value for cell in row]
            # Skip rows with fewer than 11 columns (malformed/empty rows)
            if len(values) < 11:
                continue

            # Check if this row's status is "Pending" (case-insensitive).
            # str(values[9] or '') handles None values safely.
            status = str(values[9] or '').strip()
            if status.lower() != 'pending':
                continue    # Skip non-pending tasks

            # Create a ReturnTask instance from the row data.
            # Each values[N] corresponds to a column (0-indexed):
            #   0=Address, 1=Contact, 2=Product Link, 3=Amount, etc.
            # str(values[N] or '') safely converts None to empty string.
            # float(values[N] or 0) safely converts None to 0.0.
            task = ReturnTask(
                row_number=row_idx,                                       # Excel row for write-back
                address=str(values[0] or ''),                             # Column A
                contact=str(values[1] or ''),                             # Column B
                product_link=str(values[2] or ''),                        # Column C
                amount=float(values[3] or 0),                             # Column D (₹)
                num_products=int(values[4] or 0),                         # Column E
                order_date=str(values[5] or ''),                          # Column F
                order_id=str(values[6] or ''),                            # Column G
                delivery_date=str(values[7] or ''),                       # Column H
                return_window=str(values[8] or ''),                       # Column I
                status=status,                                            # Column J
                platform=str(values[10] or ''),                           # Column K
                refund_id=str(values[11] or '') if len(values) > 11 else None,      # Column L
                return_status=str(values[12] or '') if len(values) > 12 else None,  # Column M
                refund_amount=_safe_float(values[13]) if len(values) > 13 else None, # Column N
                timestamp=str(values[14] or '') if len(values) > 14 else None,      # Column O
                log=str(values[15] or '') if len(values) > 15 else None,            # Column P
            )
            tasks.append(task)

        wb.close()  # Always close to release the file handle
        self.logger.info(f"Found {len(tasks)} pending tasks in Excel")
        return tasks

    def write_result(self, task: ReturnTask, result: ReturnResult):
        """
        Write the return result back to Excel for a specific line item.
        
        This method implements the "per-line-item write-back" requirement:
        each SKU/product gets its own result, not just the order as a whole.
        
        The method opens the workbook, writes to the specific row, saves,
        and closes. This per-operation I/O pattern is intentional:
          - If the agent crashes after processing 3 of 7 tasks, the first 3
            results are safely persisted in the file.
          - A batch approach (write all at end) would lose ALL results on crash.
        
        Args:
            task: The original task (contains row_number for write-back)
            result: The processing result to write
        """
        # Open the workbook fresh for each write (not cached — safer)
        wb = openpyxl.load_workbook(self.filepath)
        ws = wb.active
        row = task.row_number  # The exact excel row to update

        # Write result columns (L through P):
        ws.cell(row=row, column=12, value=result.return_id or '')         # Column L: Refund ID
        ws.cell(row=row, column=13, value=result.return_status)           # Column M: Return Status
        ws.cell(row=row, column=14, value=result.refund_amount)           # Column N: Refund Amount
        ws.cell(row=row, column=15,                                       # Column O: Timestamp
                value=datetime.now().isoformat())  # ISO 8601 format for consistency
        ws.cell(row=row, column=16, value=result.log_message)             # Column P: Log/Error

        # Update the Status column (J) based on the outcome:
        if result.success:
            ws.cell(row=row, column=10, value='Done')        # Task completed successfully
        elif 'human review' in result.return_status.lower():
            ws.cell(row=row, column=10, value='Needs Review') # Needs manual intervention

        # Save and close. This persists the changes to disk immediately.
        wb.save(self.filepath)
        wb.close()
        self.logger.info(f"Row {row}: Written result — Status: {result.return_status}, "
                         f"Refund ID: {result.return_id}")


# –––––––––––––––––––––––
# MAIN AGENT ORCHESTRATOR
# –––––––––––––––––––––––

class ReturnAgent:
    """
    Main orchestrator that coordinates the entire return automation workflow.
    
    This is the entry point for the agent. It:
        1. Reads pending tasks from excel
        2. Groups tasks by platform (e.g., Flipkart, Amazon)
        3. For each platform group:
            a. Starts a browser with stealth settings
            b. Logs into the platform
            c. Processes each line item independently (partial-success)
            d. Writes results back to excel after each item
        4. Prints a summary of all results
        
    The "independence" of line items is a critical design choice:
    if Item A fails (e.g., out of return window), the agent continues
    processing Items B and C. It never abandons an entire order because
    one item couldn't be returned.
    """
    
    def __init__(self, dry_run=False, headless=False, target_order_id=None):
        """
        Args:
            dry_run: If True, simulate everything without opening a browser.
                     Useful for testing the excel reading/writing pipeline.
            headless: If True, run the browser without a visible window.
            target_order_id: If set, only process this specific order ID.
        """
        self.logger = setup_logging()       # Initialize logging first
        self.dry_run = dry_run
        self.headless = headless
        self.target_order_id = target_order_id
        self.excel_manager = ExcelManager(EXCEL_FILE, self.logger)
        # Accumulates (task, result) tuples for the final summary report
        self.results_summary = []

    def run(self):
        """
        Execute the main return automation workflow.
        
        This is the top-level orchestration method that coordinates all
        components. The workflow is:
        
        Step 1: Read pending tasks from excel
        Step 2: Group tasks by platform
        Step 3: Process each platform group (browser automation)
        Step 4: Print Summary of all results
        """
        # Print startup banner with config
        self.logger.info("=" * 60)
        self.logger.info("RETURN AUTOMATION AGENT - STARTED")
        self.logger.info(f"Dry run: {self.dry_run}")
        self.logger.info(f"Target order: {self.target_order_id or 'ALL'}")
        self.logger.info("=" * 60)

        # Step 1: Read all pending tasks from excel spreadsheet
        tasks = self.excel_manager.read_pending_tasks()     # returns a ReturnTask instance
        if not tasks:
            self.logger.info("No pending tasks found. Exiting.")

        # If a specific order ID was requested, filter to just that order
        if self.target_order_id:
            tasks = [t for t in tasks if t.order_id == self.target_order_id]
            self.logger.info(f"Filtered to {len(tasks)} tasks for order {self.target_order_id}")

        # Step 2: Group tasks by platform
        platform_groups = {}
        for task in tasks:
            platform = task.platform.strip()
            if platform not in platform_groups:
                platform_groups[platform] = []
            platform_groups[platform].append(task)

        self.logger.info(f"Task groups: {', '.join(f'{k}: {len(v)}' for k, v in platform_groups.items())}")

        # Step 3: Process each platform group
        for platform, platform_tasks in platform_groups.items():
            self.logger.info(f"\n{'=' * 40}")
            self.logger.info(f"Processing platform: {platform}")
            self.logger.info(f"Tasks: {len(platform_tasks)}")
            self.logger.info(f"{'=' * 40}")

            if platform.lower() == 'flipkart':
                self._process_flipkart_tasks(platform_tasks)
            else:
                # Unsupported platform - flag all tasks for manual handling.
                # This is extensible: add elif blocks for Amazon, Meesho, etc.
                self.logger.warning(f"Platform '{platform}' not yet supported. "
                                    f"Flagging {len(platform_tasks)} tasks for review.")
                for task in platform_tasks:
                    result = ReturnResult(
                        success=False,
                        return_status="Platform not supported",
                        log_message=f"Platform '{platform}' automation not implemented"
                    )
                    self._record_result(task, result)

        # Step 4: Print final summary
        self._print_summary()

    def _process_flipkart_tasks(self, tasks):
        """
        Process all Flipkart return tasks in a single browser session.
        
        In dry-run mode, simulates processing without opening a browser.
        In live mode: starts browser -> logs in -> processes each task -> closes.
        
        The try/finally pattern ensures the browser is ALWAYS closed,
        even if the agent crashes during task processing.
        
        Args:
            tasks: list of ReturnTask instances for Flipkart orders
        """
        if self.dry_run:
            self.logger.info("Dry run mode - simulating Flipkart returns")
            for task in tasks:
                self._dry_run_task(task)
            return

        # Initialize browser and helper objects
        browser = BrowserController(self.logger, headless=self.headless)
        driver = browser.start()
        human = HumanSimulator(self.logger)
        handler = FlipkartHandler(driver, human, self.logger)

        try:
            # Authenticate with Flipkart
            login_success = handler.login(FLIPKART_PHONE)
            if not login_success:
                # If login fails after all retries, flag ALL tasks and abort.
                self.logger.error("Could not login to Flipkart. Flagging all tasks.")
                for task in tasks:
                    result = ReturnResult(
                        success=False,
                        return_status="Login failed",
                        log_message="Could not login to Flipkart"
                    )
                    self._record_result(task, result)
                return

            # Process each task INDEPENDENTLY - this is the partial-success handling.
            for i, task in enumerate(tasks):
                self.logger.info(f"\n--- Task {i + 1}/{len(tasks)} ---")
                self.logger.info(f"Order: {task.order_id} | Amount: ₹{task.amount}")

                # Pre-check: some tasks already have a status from previous runs.
                if task.return_status and task.return_status.lower() not in ['', 'none', 'pending']:
                    if 'cancelled' in task.return_status.lower():
                        result = ReturnResult(
                            success=True,
                            return_id=task.refund_id,
                            return_status="Already Cancelled & Refunded",
                            refund_amount=task.refund_amount,
                            log_message=f"Skipped: {task.return_status}"
                        )
                    elif 'not yet delivered' in task.return_status.lower():
                        result = ReturnResult(
                            success=False,
                            return_status="Not yet delivered",
                            log_message="Order not yet delivered"
                        )
                    elif 'support' in task.return_status.lower():
                        result = ReturnResult(
                            success=False,
                            return_status="Support Needed",
                            log_message=task.log or "Manual support needed"
                        )
                    else:
                        # Unknown status - try the return flow
                        result = handler.initiate_return(task)
                else:
                    # No pre-existing status - initiate the full return flow
                    result = handler.initiate_return(task)

                # Write result to excel immediately (crash-safe)
                self._record_result(task, result)
                # Pause between tasks to avoid rate limiting
                human.random_delay(2, 5)

        except Exception as e:
            self.logger.error(f"Critical error: {str(e)}")
        finally:
            # ALWAYS close the browser, even if an exception occured.
            browser.stop()

    def _dry_run_task(self, task: ReturnTask):
        # TODO
        pass

    def _record_result(self):
        """
        Record a result and write it to excel (unless in dry-run mode).
        
        This is the single point where results are persisted. Having one
        method for this ensures consistent behavior across all code paths.
        
        Args:
            task: The original task
            results: The processing result
        """
        self.results_summary.append((task, result))

        if not self.dry_run:
            try:
                self.excel_manager.write_result(task, result)
            except Exception as e:
                self.logger.error(f"Failed to write result to excel: {e}")
    
    def _print_summary(self):
        """
        Print a formatted summary of all processed tasks.
        
        Shows total/successful/failed counts and individual task outcomes.
        Uses emoji icons (✅,❌) for quick visual scanning."""
        self.logger.info("\n" + "=" * 60)
        self.logger.info("RETURN AUTOMATION - SUMMARY")
        self.logger.info("=" * 60)

        total = len(self.results_summary)
        successful = sum(1 for _, r in self.results_summary if r.success)
        failed = sum(1 for _, r in self.results_summary if not r.success)

        self.logger.info(f"Total tasks processed: {total}")
        self.logger.info(f"Successful: {successful}")
        self.logger.info(f"Failed/Needs Review: {failed}")
        self.logger.info("")

        # Print each task result with status icon and details
        for task, result in self.results_summary:
            status_icon = "✅" if result.success else "❌"
            self.logger.info(f"  {status_icon} Order {task.order_id} | "
                             f"₹{task.amount} | {result.return_status}")
            if result.log_message:
                self.logger.info(f"     └─ {result.log_message}")

        self.logger.info("=" * 60)


# –––––––––––––––
# CLI ENTRY POINT
# –––––––––––––––

def main():
    """
    Parse command-line arguments and launch the return automation agent.
    
    Supported flags:
        --dry-run: Simulate everyting without opening a browser or modifying excel
        --headless: Run the browser without a visible window
        --order-id: Process only a specific order (useful for debugging)

    Examples:
        python3 return_agent.py --dry-run
        python3 return_agent.py
        python3 return_agent.py --order-id OD123456
        python3 return_agent.py --headless
    """
    # argparse creates a cli with help text, type checking, and default values
    parser = argparse.ArgumentParser(
        description='Return Automation Agent - Automates product on Flipkart'
    )
    parser.add_argument('--dry-run', action='store_true',
                        help='Run without browser interaction (simulation mode)')
    parser.add_argument('--headless', action='store_true',
                        help='Run browser in headless mode')
    parser.add_argument('--order-id', type=str, default=None,
                        help='Process a specific order ID only')

    # Parse the arguments from sys.argv
    args = parser.parse_args()

    # Create and run the agent with the parsed config
    agent = ReturnAgent(
        dry_run=args.dry_run,
        headless=args.headless,
        target_order_id=args.order_id
    )
    agent.run()


if __name__ == '__main__':
    main()