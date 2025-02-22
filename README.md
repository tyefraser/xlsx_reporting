# xlsx_reporting
Automate Reporting using Excel (xlsx) as your reporting tool.


Proposed file structure:
your_project/
│── src/                          # Source code
│   ├── __init__.py               # Makes src a package
│   ├── main.py                   # Main script to run the program
│   ├── file_loader.py            # Handles reading various file types
│   ├── data_validator.py         # Checks and validates data formats
│   ├── report_generator.py       # Creates the Excel report
│   ├── utils.py                  # Utility functions (e.g., logging, formatting)
│
│── config/                        # Configuration files
│   ├── settings.yaml              # Defines acceptable formats, column names, etc.
│
│── tests/                         # Unit tests
│   ├── test_file_loader.py
│   ├── test_data_validator.py
│   ├── test_report_generator.py
│
│── examples/                      # Example input files
│   ├── sample.csv
│   ├── sample.xlsx
│
│── output/                        # Output reports
│   ├── generated_report.xlsx
│
│── requirements.txt               # Dependencies
│── README.md                      # Project documentation
│── .gitignore                      # Ignore unnecessary files
│── setup.py                        # If you want to make it installable


Notes for the data laoded
- CSV: Can only have one set of data loaded
- xlsx:
    - sheet_<sheet_name>: can only have one set of data
    - table_<table_name>: can only have one set of data
