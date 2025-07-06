# BOQ Tools - Bill of Quantities (BOQ) Excel Processor

A powerful and intelligent desktop application designed to streamline the processing and analysis of Bill of Quantities (BOQ) Excel files. It automates tedious tasks like data extraction, classification, and validation, providing a comprehensive suite of tools for professionals dealing with BOQs.

## Key Features

- **Intelligent File Processing**: Automatically reads and understands complex BOQ Excel files, including multi-sheet documents.
- **Interactive User Interface**: A user-friendly graphical interface to manage, process, and review BOQ files.
- **Advanced Categorization**: Sophisticated automatic and manual categorization of BOQ items.
- **Powerful Validation**: A robust validation engine to ensure data integrity and consistency.
- **High Configurability**: Easily customize the application's behavior to fit specific BOQ formats and project requirements.
- **Flexible Exporting**: Export processed data into well-formatted Excel files.

## Functionalities in Detail

### File Processing Engine

The core of the application is a sophisticated engine that can handle a wide variety of BOQ file formats.

- **Broad File Support**: Natively processes both modern (`.xlsx`) and legacy (`.xls`) Excel formats.
- **Memory Efficient**: Designed to handle large files without consuming excessive memory.
- **Sheet Classification**: Automatically identifies and classifies different types of sheets within a workbook (e.g., Main BOQ, Summary, Preliminaries, Notes).
- **Intelligent Column Mapping**: Automatically detects and maps columns to standard BOQ fields like `Description`, `Quantity`, `Unit`, `Rate`, and `Total`, even with non-standard column headers.
- **Row Classification**: Distinguishes between header rows, data rows, sub-total rows, and ignored rows.

### User Interface

The application provides a rich, interactive user interface that makes processing BOQs intuitive and efficient.

- **Main Dashboard**: A central window to add, remove, and manage BOQ files for processing.
- **Data Preview**: Preview the contents of an Excel file before committing to a full processing run.
- **Live Progress Tracking**: Real-time progress dialogs for long-running operations.
- **Interactive Categorization**:
    - **Sheet Categorization**: If the app is unsure about a sheet's type, it will prompt the user to classify it manually.
    - **Row Categorization**: For items that cannot be categorized automatically, a dialog allows for quick manual categorization.
- **Review and Correction**:
    - **Category Review**: A dedicated dialog to review, edit, and approve the automatic categorization results.
    - **Row Review**: Inspect and correct the classification of individual rows.
- **Categorization Statistics**: View detailed statistics about the categorization process, including categorized and uncategorized items.
- **Centralized Settings**: A comprehensive settings dialog to manage all application configurations.

### Advanced Categorization

- **Automatic Categorization**: Leverages a customizable keyword-based dictionary to automatically categorize BOQ items.
- **Manual Override**: Full control to manually categorize items that are ambiguous or require specific classification.
- **Category Dictionary Management**: Easily manage and extend the category dictionary used for automatic categorization.

### Validation

- **Data Integrity Checks**: Validates data against a set of configurable rules.
- **Confidence Scoring**: Provides confidence scores for automatic classifications to help identify potential errors.

### Exporting

- **Styled Excel Exports**: Export the processed and cleaned data into a new, well-formatted, and styled Excel file.
- **Inclusion of Reports**: Option to include validation and summary reports in the exported file.

### Command-Line Interface (CLI)

For automation and power users, a CLI provides access to the application's core functionalities.

- **File Processing**: Process one or more files directly from the command line.
- **Batch Operations**: Script batch processing of multiple BOQ files.
- **Interactive Mode**: An interactive CLI mode for guided processing.

## Saving and Resuming Work

The application allows you to save your analysis and mappings to a file. This is useful for resuming your work later or for reusing a set of mappings on a new BOQ file that has a similar structure.

- **Save Analysis**: Saves the entire state of your current analysis, including the processed data and all mappings, to a `.pkl` file. This allows you to close the application and perfectly restore your session later.
- **Save Mappings**: Saves only the sheet, column, and row mappings to a `.pkl` file. This is ideal for creating a reusable template for BOQs with a consistent layout.
- **Load Analysis/Mappings**: You can load a previously saved `.pkl` file to either resume a session or apply a saved mapping to a new file.



## Configuration

The application is highly configurable via the `config/boq_settings.json` file. Through the UI, settings can be modified in the **Settings Dialog**.

Key configurable areas include:
- **Column Mappings**: Keywords used to identify columns.
- **Sheet Classifications**: Keywords to classify Excel sheets.
- **Validation Thresholds**: Rules for the data validation engine.
- **Processing Limits**: Settings for file size, memory usage, etc.

## Dependencies

- **pandas**: For data manipulation and analysis.
- **openpyxl**: For reading and writing Excel files.
- **xlrd**: For reading legacy `.xls` Excel files.
- **PyInstaller**: For packaging the application into a standalone executable.
- **PyQt5 / PySide6**: The application requires a Qt binding for its user interface. Please install one of them manually (`pip install PyQt5` or `pip install PySide6`).

