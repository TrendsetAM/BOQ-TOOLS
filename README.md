# BOQ Tools - Bill of Quantities Excel Processor

A comprehensive Python application for processing and analyzing Bill of Quantities (BoQ) Excel files with intelligent column mapping, sheet classification, and validation.

## Features

- **Intelligent Column Mapping**: Automatically identifies and maps columns based on common BoQ terminology
- **Sheet Classification**: Classifies different types of BoQ sheets (main, summary, preliminaries, etc.)
- **Validation System**: Comprehensive validation with configurable thresholds
- **Modular Architecture**: Clean separation of concerns with core, UI, and utility modules
- **Configuration System**: Flexible configuration for different BoQ formats and requirements

## Project Structure

```
BOQ-TOOLS/
├── main.py                 # Main application entry point
├── requirements.txt        # Python dependencies
├── setup.py               # Package configuration
├── run.bat                # Windows launcher script
├── README.md              # This file
├── .gitignore             # Git ignore rules
├── core/                  # Core business logic
│   ├── __init__.py
│   └── boq_processor.py   # Main BOQ processor class
├── ui/                    # User interface
│   ├── __init__.py
│   └── main_window.py     # Tkinter UI
├── utils/                 # Utilities
│   ├── __init__.py
│   ├── config.py          # Configuration system
│   └── logger.py          # Logging setup
├── resources/             # Resources folder
│   └── .gitkeep
└── examples/              # Example scripts
    └── config_demo.py     # Configuration system demo
```

## Installation

1. **Clone the repository**:
   ```bash
   git clone <repository-url>
   cd BOQ-TOOLS
   ```

2. **Create virtual environment**:
   ```bash
   python -m venv venv
   ```

3. **Activate virtual environment**:
   - Windows: `venv\Scripts\activate`
   - Linux/Mac: `source venv/bin/activate`

4. **Install dependencies**:
   ```bash
   pip install -r requirements.txt
   ```

## Usage

### Running the Application

**Option 1: Using the launcher script**
```bash
run.bat
```

**Option 2: Direct Python execution**
```bash
python main.py
```

**Option 3: Using virtual environment**
```bash
venv\Scripts\python.exe main.py
```

### Configuration System Demo

Run the configuration system demo to see how the intelligent mapping works:

```bash
python examples/config_demo.py
```

## Configuration System

The BOQ Tools configuration system (`utils/config.py`) provides comprehensive settings for:

### Column Mappings

The system automatically maps Excel columns to BoQ data types using keyword matching:

- **Description**: `["description", "item", "work", "activity", "task", "detail", ...]`
- **Quantity**: `["qty", "quantity", "no", "number", "count", ...]`
- **Unit Price**: `["unit price", "rate", "price per unit", "unit cost", ...]`
- **Total Price**: `["total", "total price", "total cost", "value", ...]`
- **Classification**: `["type", "category", "class", "mandatory", "optional", ...]`
- **Unit**: `["unit", "measurement", "measure", "uom", "unit of measure", ...]`
- **Code**: `["code", "item code", "reference", "ref", "item no", ...]`
- **Remarks**: `["remarks", "notes", "comments", "observation", "note", ...]`

### Sheet Classifications

Automatically classifies different types of BoQ sheets:

- **BOQ Main**: Main bill of quantities
- **Summary**: Summary and total sheets
- **Preliminaries**: General items and requirements
- **Substructure**: Foundation and substructure works
- **Superstructure**: Structural and framing works
- **Finishes**: Interior and exterior finishes
- **Services**: MEP and building services
- **External Works**: Site works and landscaping

### Validation Thresholds

Configurable validation settings:

- Minimum column confidence: 0.7
- Minimum sheet confidence: 0.6
- Maximum empty rows percentage: 30%
- Minimum data rows: 5
- Maximum header rows: 10

### Processing Limits

Performance and resource limits:

- Maximum file size: 50 MB
- Maximum sheets per file: 20
- Maximum rows per sheet: 10,000
- Maximum columns per sheet: 50
- Timeout: 300 seconds
- Memory limit: 512 MB

### Export Settings

Default export configuration:

- Output format: XLSX
- Include summary: Yes
- Include validation report: Yes
- Backup original: Yes
- Compression level: 6

## Using the Configuration System

```python
from utils.config import get_config, ColumnType

# Get configuration instance
config = get_config()

# Get column mapping for a specific type
description_mapping = config.get_column_mapping(ColumnType.DESCRIPTION)

# Get all required columns
required_columns = config.get_required_columns()

# Classify a sheet
sheet_type, confidence = config.get_sheet_classification("BOQ Main", ["description", "quantity"])

# Access validation thresholds
min_confidence = config.validation_thresholds.min_column_confidence

# Access processing limits
max_file_size = config.processing_limits.max_file_size_mb
```

## Development

### Adding New Column Types

1. Add new enum value to `ColumnType`
2. Add mapping configuration in `_setup_column_mappings()`
3. Update validation patterns if needed

### Adding New Sheet Classifications

1. Add new `SheetClassification` in `_setup_sheet_classifications()`
2. Define keywords and confidence thresholds
3. Test with `config_demo.py`

### Customizing Validation

Modify the `ValidationThresholds` dataclass to adjust:
- Confidence thresholds
- Row limits
- Data requirements

## Dependencies

- **pandas**: Data manipulation and analysis
- **openpyxl**: Excel file reading and writing
- **PyInstaller**: Creating standalone executables
- **pathlib2**: Path manipulation (for older Python versions)

## License

[Add your license information here]

## Contributing

[Add contribution guidelines here]

## Support

[Add support information here] 