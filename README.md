# Excel Data Cleaner

## Overview
Excel Data Cleaner is a Python-based tool designed to streamline the process of cleaning and organizing data within Excel spreadsheets. This tool automates common data cleaning tasks, making it easier for users to prepare their data for analysis.

## Features
- **Remove Duplicates**: Automatically identifies and removes duplicate rows based on specified columns.
- **Fill Missing Values**: Offers options to fill missing values with user-defined constants or statistical measures (mean, median).
- **Format Data**: Standardizes data formats (e.g., date formats, text casing).
- **Filter Data**: Allows users to filter data based on specific criteria.
- **Export Cleaned Data**: Saves the cleaned data back to an Excel file or CSV format.

## Requirements
- Python 3.6 or higher
- Pandas library
- OpenPyXL library (for Excel file handling)
- NumPy library (for numerical operations)

## Installation
1. Clone the repository:
   ```bash
   git clone https://github.com/yourusername/excel-data-cleaner.git
   ```
2. Navigate to the project directory:
   ```bash
   cd excel-data-cleaner
   ```
3. Install the required packages:
   ```bash
   pip install -r requirements.txt
   ```

## Usage
1. Import the necessary modules:
   ```python
   from excel_data_cleaner import DataCleaner
   ```
2. Create an instance of the `DataCleaner` class:
   ```python
   cleaner = DataCleaner('path/to/your/excel/file.xlsx')
   ```
3. Call the desired cleaning methods:
   ```python
   cleaner.remove_duplicates()
   cleaner.fill_missing_values(method='mean')
   cleaner.format_data()
   ```
4. Export the cleaned data:
   ```python
   cleaner.export_cleaned_data('path/to/save/cleaned_data.xlsx')
   ```

## Contributing
Contributions are welcome! Please fork the repository and submit a pull request with your changes.

## License
This project is licensed under the MIT License. See the LICENSE file for details.

## Contact
For any inquiries, please contact [your.email@example.com](mailto:your.email@example.com).