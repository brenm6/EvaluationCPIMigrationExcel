# Excel Evaluation Tool

A Python-based tool for processing and evaluating SAP CPI (Cloud Platform Integration) migration Excel files. This tool automates the analysis of integration scenarios, generating comprehensive evaluation reports with detailed metrics and recommendations.

## Overview

The Excel Evaluation Tool processes SAP CPI evaluation data and generates a consolidated report with integration scenario analysis, including adapter types, mappings, quality of service metrics, effort estimations, and modernization recommendations.

## Features

- **Automated Data Processing**: Extracts and analyzes integration scenarios from source Excel files
- **Comprehensive Metrics**: Calculates various metrics including:
  - Adapter types (Sender/Receiver)
  - Module presence detection
  - Mapping types (XSLT, Java, Message Mapping)
  - Quality of Service (QoS) analysis
  - UDF and Function Library usage
  - Interface counts (FTP, SFTP, FTPS)
  - Effort estimations (Min/Max/Average hours)
- **Visual Formatting**: Applies color-coding and styling to the output Excel file
- **Empty Cell Protection**: Uses whitespace in empty cells to prevent text overflow from adjacent cells
- **GUI Interface**: User-friendly Tkinter-based interface for file selection and processing
- **Command-Line Support**: Can be run with or without the GUI

## Project Structure

```
Excel Evaluation Tool/
├── Excel_Manager.py       # Core processing logic and data extraction
├── Columns_Manager.py     # Excel formatting and styling utilities
├── Headers.py            # Column header definitions
├── Frontend.py           # GUI interface (Tkinter)
├── README.md            # This file
└── evaluation_run_results_input_PA3_2025-07-18.xlsx  # Sample input file
```

## Components

### Excel_Manager.py
The main processing engine that:
- Loads and parses the source Excel file
- Builds lookup tables for various data points
- Extracts integration scenarios and their properties
- Calculates metrics and statistics
- Generates the output "Evaluation" sheet
- Applies borders and formatting to cells

Key methods:
- `__init__(filename)`: Initializes the manager and loads the workbook
- `create_sheet(title, index)`: Creates a new worksheet
- `set_columns(sheet)`: Sets up column headers
- `fill_sheet(sheet)`: Main processing method that populates the evaluation sheet
- `save()`: Saves the processed workbook to "Test_____File.xlsx"

### Columns_Manager.py
Handles Excel cell formatting and styling:
- Sets headers with custom fonts and colors
- Applies color coding (green, orange, light blue)
- Manages column widths
- Creates bold text formatting

Key methods:
- `set_headers(headers, worksheet)`: Sets header row with yellow background
- `set_colour_green(worksheet, columnnumber)`: Applies green fill to column
- `set_clour_orange(worksheet, columnnumber)`: Applies orange fill to column
- `set_colour_light_blue(worksheet, columnnumber)`: Applies light blue fill to column
- `set_column_width(worksheet, columnnumber, width)`: Sets column width
- `first_line_bold(worksheet)`: Makes first row bold with large font

### Headers.py
Defines the 48 column headers for the evaluation report, including:
- Basic information (Scenario, Type, Party, Components)
- Adapter information (Sender/Receiver types and modules)
- Module flags (MTB, MLB, PGPE, etc.)
- Metrics (Interface counts, QoS, UDF usage)
- Mapping details (XSLT, Java, ABAP, MM)
- Effort estimations
- Modernization recommendations

### Frontend.py
Provides a graphical user interface using Tkinter:
- File selection dialog
- Processing status display
- Error handling with user-friendly messages
- Black-themed modern UI with Microsoft-style blue buttons

## Requirements

```python
openpyxl>=3.0.0
tkinter (usually comes with Python)
line_profiler (optional, for performance profiling)
```

## Installation

1. Clone or download this repository
2. Install required dependencies:
```bash
pip install openpyxl line_profiler
```

## Usage

### GUI Mode (Recommended)

Run the frontend application:
```bash
python Frontend.py
```

1. Click "Datei auswählen" (Select File)
2. Choose your CPI evaluation Excel file (e.g., `evaluation_run_results_PA3_...xlsx`)
3. Wait for processing to complete
4. The output file "Test_____File.xlsx" will be created in the same directory

### Command-Line Mode

Run the Excel Manager directly:
```bash
python Excel_Manager.py
```

The script will process the hardcoded input file and generate the output.

## Input File Requirements

The input Excel file must contain the following sheets:
- **Full Evaluation Results**: Main data source with integration scenarios
- **Eval by Integration Scenario**: T-shirt sizing and effort data
- **Recommendations**: Modernization recommendations

Expected columns in "Full Evaluation Results":
1. Integration Scenario (Column 1)
2. Rule (Column 2)
3. Value (Column 4)

## Output

The tool generates a single "Evaluation" sheet with 48 columns containing:
- Numbered scenarios
- Integration scenario details (Party, Components, Interfaces)
- Adapter types and module information
- Module presence flags (MTB, MLB, PGPE, etc.)
- Mapping information
- Interface counts and metrics
- Quality of Service data
- Effort estimations
- Modernization recommendations

### Color Coding
- **Yellow**: Header row
- **Green**: Module flags (columns 13-20) and core metrics
- **Orange**: Selected analysis columns
- **Light Blue**: Effort estimation columns

## Data Processing Logic

1. **Sorting**: Integration scenarios are sorted alphabetically
2. **Deduplication**: Only the first occurrence of each scenario is processed
3. **Aggregation**: Multiple rules per scenario are aggregated into counts or flags
4. **Empty Cell Handling**: All empty cells contain a whitespace to prevent text overflow
5. **Formula Generation**: Summary formulas (SUM, COUNTIF) are added to row 1

## Special Features

### Empty Cell Protection
All empty cells are filled with a single whitespace character (" ") to prevent Excel from allowing text from adjacent cells to overflow visually. This ensures clean, professional-looking reports.

### Lookup Tables
The tool builds multiple lookup tables for efficient data retrieval:
- Adapter types (Sender/Receiver)
- Module presence
- Quality of Service
- Mapping types
- UDF and Function Library usage
- T-shirt sizing
- Effort estimations

### Rule Detection
Special rules are detected for:
- UDF usage (Dynamic Configuration, Lookup Service, File OS)
- Function Library usage
- EOIO (Exactly Once In Order) processing

## Error Handling

- **Permission Errors**: Detects if output file is open and prompts user to close it
- **Missing Sheets**: Validates required sheets exist in input file
- **Data Validation**: Handles missing or malformed data gracefully

## Performance

The tool uses the `line_profiler` decorator for performance monitoring. Processing time varies based on:
- Number of integration scenarios
- Complexity of rules
- File size

Typical processing time: 5-30 seconds for standard files.

## Limitations

- Input file path is currently hardcoded in command-line mode
- Output filename is fixed as "Test_____File.xlsx"
- Requires specific sheet names and column structure
- Some columns are hardcoded to "n/a" (ABAP-related fields)

## Future Enhancements

- Configurable input/output file paths
- Support for additional data sources
- Export to multiple formats (CSV, JSON)
- Enhanced error reporting
- Progress bar for long operations
- Configuration file support

## Troubleshooting

**Issue**: "Permission denied" error
- **Solution**: Close the output Excel file if it's currently open

**Issue**: Missing sheets error
- **Solution**: Ensure input file contains required sheets: "Full Evaluation Results", "Eval by Integration Scenario", "Recommendations"

**Issue**: Incorrect data in output
- **Solution**: Verify input file follows expected column structure

## Contributing

For bugs, feature requests, or contributions, please contact the repository owner.

## License

Internal tool for ATOS use.

## Author

Created for SAP CPI migration evaluation projects.

## Version History

- **Latest**: Added whitespace protection for empty cells
- Previous versions included basic evaluation functionality

---

*Last Updated: October 2025*
