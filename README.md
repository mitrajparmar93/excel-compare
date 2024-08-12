# excel-compare
This Python project automatically compares a newly downloaded Excel file with a local Excel file, updates the local file with any new data, and highlights the changes. This script is set up to run easily on Windows using a shortcut.
## Requirements

Python 3.10.13+
Poetry for dependency management
Excel files in `.xlsx` format

## Install python 3.10 on windows

1. **Download Python**:
    * Go to the [Python Downloads](https://www.python.org/downloads/windows/) page.
    * Click the "Latest Python 3 Release" link.
    * Scroll down and select "Windows installer (64-bit)" under "Files"
    * Download the executable installer and run it.
2. **Install Python**:
    * Make sure to check the box that says "Add Python 3.x to PATH" before clicking "Install Now."
    * Click "Disable path length limit" if you see the option.
3. **Verify Installation**:
    * Open a command prompt and type `python --version`.
    * You should see the version of Python you installed.

## Setup

1. **Clone the repository** (or create a new directory for your project):
   ```bash
   git clone git@github.com:mitrajparmar93/excel-compare.git
   cd excel-compare
   ```

1. **Install Poetry** (if not already installed):
   ```bash
   curl -sSL https://install.python-poetry.org | python3 -
   ```

1. **Install the dependencies**:
   ```bash
   poetry install
   ```

2. **Edit the script**:
    * Modify the `DIRECTORY` and `LOCAL_FILE` variables in `excel_comparison/compare.py` to point to the correct paths where your files are located.

## Usage

### Running the Script with a Shortcut

1. **Create a `.bat` file** in the project directory:
    * Create a `run_comparison.bat` file with the following content:
     ```batch
     @echo off
     cd path\to\your\project
     poetry run python excel_comparison/compare.py
     pause
     ```
    * Replace `path\to\your\project` with the actual path to your project directory.

1. **Create a Windows shortcut**:
    * Right-click the `.bat` file and select "Create shortcut."
    * Rename and place the shortcut wherever it’s convenient (e.g., Desktop).

1. **Run the script**:
    * Double-click the shortcut to execute the script. The script will automatically find the latest Excel file in the specified directory, compare it with your local file, update the local file with any new data, and highlight the changes.

## Customization

* **File Extension**: The script is set to find files with the `.xlsx` extension. Modify the `extension` variable in the `get_latest_file` function if you need to handle other file types.
* **Highlight Color**: The script highlights new rows in yellow by default. You can change the `PatternFill` color in the script if needed.
