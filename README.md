<!-- TABLE OF CONTENTS -->
<details>
    <summary>Table of Contents</summary>
    <ol>
        <li><a href="#excel-column-based-blank-cell-row-cleaner">Excel Column-Based Blank Cell Row Cleaner</a></li>
        <li><a href="#overview">Overview</a></li>
        <li><a href="#getting-started">Getting Started</a>
            <ul>
                <li><a href="#prerequisites">Prerequisites</a></li>
                <li><a href="#installation">Installation</a></li>
            </ul>
        </li>
        <li><a href="#usage">Usage</a>
            <ul>
                <li><a href="#pandas_cleanerpy">pandas_cleaner.py</a></li>
                <li><a href="#openpyxl_cleanerpy">openpyxl_cleaner.py</a></li>
            </ul>
        </li>
        <li><a href="#license">License</a></li>
    </ol>
</details>

<!-- EXCEL COLUMN-BASED BLANK CELL ROW CLEANER -->
# Excel Column-Based Blank Cell Row Cleaner

![Demo](assets/demo.gif)

Scripts for removing rows with blank cells in specified columns from Excel files. The scripts use two different libraries: `pandas` and `openpyxl`. Each script serves a similar purpose but employs a different approach to handling Excel files.

<!-- OVERVIEW -->
## Overview

- **[`pandas_cleaner.py`](pandas_cleaner.py)**: Utilizes the `pandas` library for efficient data processing, ideal for large datasets. This script does not preserve original formatting (e.g., fonts, colors).
- **[`openpyxl_cleaner.py`](openpyxl_cleaner.py)**: Utilizes the `openpyxl` library, which preserves original formatting and ensures the visual presentation of the Excel file remains unchanged. This is suitable when formatting needs to be retained.

<!-- GETTING STARTED -->
## Getting Started

<!-- PREREQUISITES -->
### Prerequisites

- **Python**: You can download the latest version of Python [here](https://www.python.org/downloads/).

- **Required Libraries**: Install the necessary Python libraries using pip:

    ```bash
    pip install pandas openpyxl
    ```

<!-- INSTALLATION -->
### Installation

1. Clone the repository:

    ```bash
    git clone https://github.com/23Shade/excel-column-based-blank-cell-row-cleaner.git
    ```

2. Navigate to the software directory:

    ```bash
    cd excel-column-based-blank-cell-row-cleaner
    ```

<!-- USAGE -->
## Usage

### `pandas_cleaner.py`

1. Open the `pandas_cleaner.py` file in a text editor.
2. Update the following parameters:
    - **`input_file_path`**: Path to the input Excel file (e.g., `C:\path\to\your_file.xlsx`).
    - **`sheet_name`**: Name of the sheet to process (e.g., `Sheet1`).
    - **`target_columns`**: List of columns to check for blank cells (e.g., `['Email', 'Phone']`). You can add more columns as needed.
    - **`output_file_path`**: Path where the cleaned Excel file will be saved (e.g., `C:\path\to\output_file.xlsx`).
3. Save your changes.
4. Run the script:

    ```bash
    python pandas_cleaner.py
    ```

### `openpyxl_cleaner.py`

1. Open the `openpyxl_cleaner.py` file in a text editor.
2. Update the following parameters:
    - **`input_file_path`**: Path to the input Excel file (e.g., `C:\path\to\your_file.xlsx`).
    - **`sheet_name`**: Name of the sheet to process (e.g., `Sheet1`).
    - **`target_columns`**: List of columns to check for blank cells (e.g., `['Email', 'Phone']`). You can add more columns as needed.
    - **`output_file_path`**: Path where the cleaned Excel file will be saved (e.g., `C:\path\to\output_file.xlsx`).
3. Save your changes.
4. Run the script:

    ```bash
    python openpyxl_cleaner.py
    ```

<!-- LICENSE -->
## License

This software is licensed under the MIT [LICENSE](LICENSE).
