# excel-compare

This Python project automatically compares a newly downloaded Excel file with a local Excel file, updates the local file with any new data, and highlights the changes. This script is set up to run easily on Windows using a shortcut.

## Requirements

Python 3.10.13+
Poetry for dependency management
Excel files in `.xlsx` format

## Install python 3.10 on windows

1. **Download Python**:
   - Go to the [Python Downloads](https://www.python.org/downloads/windows/) page.
   - Click the "Latest Python 3 Release" link.
   - Scroll down and select "Windows installer (64-bit)" under "Files"
   - Download the executable installer and run it.
2. **Install Python**:
   - Make sure to check the box that says "Add Python 3.x to PATH" before clicking "Install Now."
   - Click "Disable path length limit" if you see the option.
3. **Verify Installation**:
   - Open a command prompt and type `python --version`.
   - You should see the version of Python you installed.

## Setup

1. **Open Command Prompt**:

   - On Windows, press `Win + R`, type `cmd`, and press `Enter`.
   - On macOS, press `Cmd + Space`, type `Terminal`, and press `Enter`.
   - On Linux, press `Ctrl + Alt + T`.

2. **Clone the repository** (or create a new directory for your project):

   - In the terminal or command prompt, type the following commands and press `Enter` after each line:
     ```bash
     git clone git@github.com:mitrajparmar93/excel-compare.git
     cd excel-compare
     ```

3. **Install Python** (if not already installed):

   - Download and install Python from the official website: [Python Downloads](https://www.python.org/downloads/).
   - During installation, make sure to check the box that says "Add Python to PATH".

4. **Verify Python Installation**:

   - In the terminal or command prompt, type the following command and press `Enter`:
     ```bash
     python --version
     ```
   - You should see the Python version number.

5. **Install Poetry** (if not already installed):

   - In the terminal or command prompt, type the following command and press `Enter`:
     ```bash
     curl -sSL https://install.python-poetry.org | python3 -
     ```

6. **Verify Poetry Installation**:

   - In the terminal or command prompt, type the following command and press `Enter`:
     ```bash
     poetry --version
     ```
   - You should see the Poetry version number.

7. **Install the dependencies**:

   - In the terminal or command prompt, make sure you are in the `excel-compare` directory by typing:
     ```bash
     cd excel-compare
     ```
   - Then, type the following command and press `Enter` to install the project dependencies:
     ```bash
     poetry install
     ```

8. **Run the Project**:
   - In the terminal or command prompt, type the following command and press `Enter` to run the project:
     ```bash
     poetry run python src/main.py
     ```

By following these steps, you will have set up the project, installed all necessary dependencies, and run the project using the command line interface.

2. **Edit the script**:
   - Modify the `DIRECTORY` and `LOCAL_FILE` variables in `excel_comparison/compare.py` to point to the correct paths where your files are located.

## Usage

### Running the Script with a Shortcut

1. **Create a `.bat` file** in the project directory:

   - Create a `run_comparison.bat` file with the following content:

   ```batch
   @echo off
   cd path\to\your\project
   poetry run python excel_comparison/compare.py
   pause
   ```

   - Replace `path\to\your\project` with the actual path to your project directory.

1. **Create a Windows shortcut**:

   - Right-click the `.bat` file and select "Create shortcut."
   - Rename and place the shortcut wherever it’s convenient (e.g., Desktop).

1. **Run the script**:
   - Double-click the shortcut to execute the script. The script will automatically find the latest Excel file in the specified directory, compare it with your local file, update the local file with any new data, and highlight the changes.

## Customization

- **File Extension**: The script is set to find files with the `.xlsx` extension. Modify the `extension` variable in the `get_latest_file` function if you need to handle other file types.
- **Highlight Color**: The script highlights new rows in yellow by default. You can change the `PatternFill` color in the script if needed.
