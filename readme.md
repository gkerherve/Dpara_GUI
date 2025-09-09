Here's the README.md with proper markdown formatting:

```markdown
# D-Parameter Measurement Tool

## Overview

The D-Parameter Measurement Tool is a wxPython-based desktop application designed for analyzing spectroscopic data from Excel files. It provides advanced smoothing and differentiation capabilities to calculate D-parameter values, commonly used in X-ray photoelectron spectroscopy (XPS) and other analytical techniques.

## Features

### Core Functionality
- **Excel File Support**: Read data from Excel files (.xlsx, .xls) with multiple sheet support
- **Data Visualization**: Interactive matplotlib plots with automatic axis inversion for spectroscopic data
- **D-Parameter Analysis**: Advanced smoothing and differentiation algorithms for peak analysis
- **Data Export**: Save results back to Excel with derivative data included

### Analysis Parameters
- **Smooth Width**: Adjustable smoothing strength (0.1-10.0)
- **Pre-Smooth Passes**: Number of initial smoothing iterations (0-10)
- **Differentiation Width**: Width parameter for derivative calculation (0.1-10.0 eV)
- **Post-Smooth Passes**: Number of final smoothing iterations (0-10)
- **Multiple Algorithms**: Choice of smoothing methods including:
  - Gaussian filtering
  - Savitzky-Golay filtering
  - Moving Average
  - Wiener filtering
  - No smoothing option

## Installation Requirements

### Python Dependencies
```bash
pip install wxpython numpy pandas matplotlib scipy openpyxl
```

Required packages:
- wxPython
- numpy
- pandas
- matplotlib
- scipy
- openpyxl (for Excel file handling)

## Usage Instructions

### Getting Started
1. **Launch the Application**: Run `python Main.py` from the command line
2. **Open Excel File**: Use File → Open (Ctrl+O) to select your Excel file
3. **Select Data Sheet**: Choose the appropriate sheet from the dropdown menu
4. **Configure Analysis**: Adjust D-parameter analysis settings as needed
5. **Calculate**: Click the "Calculate" button to perform analysis
6. **Save Results**: Use File → Save (Ctrl+S) to export results

### Data Format Requirements
- **Column A**: X-axis values (typically binding energy in eV)
- **Column B**: Y-axis values (typically intensity counts)
- **Headers**: First row should contain column headers
- **Data Quality**: Minimum 2 data points required, NaN values automatically removed

### Menu Options
- **File Menu**:
  - Open (Ctrl+O): Load Excel file
  - Save (Ctrl+S): Save results to Excel
  - Exit (Ctrl+Q): Close application
- **Help Menu**:
  - Help (F1): Display usage instructions

## Technical Details

### Data Processing Pipeline
1. **Data Import**: Excel file parsing with pandas
2. **Pre-Processing**: NaN removal and data validation
3. **Smoothing**: Configurable algorithm application
4. **Differentiation**: Gradient calculation with normalization
5. **Peak Detection**: Minimum/maximum identification
6. **D-Parameter Calculation**: Distance measurement between extrema

### Smoothing Algorithms
- **Gaussian**: Uses `scipy.ndimage.gaussian_filter` for optimal noise reduction
- **Savitzky-Golay**: Polynomial-based smoothing preserving peak shapes
- **Moving Average**: Simple windowed averaging
- **Wiener**: Adaptive noise reduction filtering

### Output Data Structure
The application maintains a hierarchical data structure:
```python
Data['Core levels'][sheet_name]['Fitting']['Peaks']['D-parameter']
```

Contains:
- Position (center point)
- FWHM (D-parameter value)
- Analysis parameters
- Derivative data array

## User Interface Components

### Main Window Layout
- **Left Panel**: Control parameters and settings
- **Right Panel**: Interactive matplotlib plot canvas
- **Menu Bar**: File operations and help access
- **Status Controls**: Sheet selection and parameter adjustment

### Control Parameters
- **Numerical Spinners**: Precise parameter adjustment
- **Dropdown Menus**: Algorithm and sheet selection
- **Action Buttons**: Calculate and clear operations
- **Read-only Display**: D-parameter result presentation

## File Operations

### Opening Files
- Supports standard Excel formats (.xlsx, .xls)
- Automatic sheet detection and population
- Error handling for invalid files or formats

### Saving Results
- Preserves original data structure
- Adds derivative column to current sheet
- Maintains other sheets unchanged
- Overwrites existing files with confirmation

## Error Handling

### Common Issues
- **File Format Errors**: Comprehensive error messages for unsupported formats
- **Data Validation**: Automatic handling of insufficient data points
- **Calculation Errors**: Graceful handling of numerical computation issues
- **Memory Management**: Efficient handling of large datasets

### User Feedback
- **Status Messages**: Clear success/error notifications
- **Progress Indication**: Real-time calculation feedback
- **Input Validation**: Parameter range enforcement

## Advanced Features

### Plot Customization
- **Automatic Scaling**: Optimal axis range selection
- **Legend Management**: Clear data series identification
- **Grid Display**: Enhanced data visualization
- **Color Coding**: Distinct colors for original and derivative data

### Data Analysis
- **Peak Detection**: Automatic extrema identification
- **Normalization**: Range-based derivative scaling
- **Center Calculation**: Accurate peak center determination
- **Quality Metrics**: Statistical analysis parameters

## Troubleshooting

### Common Problems
- **Import Errors**: Verify all dependencies are installed
- **File Access Issues**: Check file permissions and format compatibility
- **Calculation Problems**: Ensure data contains valid numerical values
- **Display Issues**: Update matplotlib and wxPython to latest versions

### Performance Optimization
- **Large Files**: Application handles datasets up to several thousand points
- **Memory Usage**: Efficient numpy array operations
- **Response Time**: Real-time parameter adjustment and visualization

## Screenshots

*Add screenshots of the application interface here*

## License and Support

This tool is designed for scientific and research applications. For technical support or feature requests, refer to the help menu within the application.

## Version Information

Current version includes full Excel integration, multiple smoothing algorithms, and comprehensive data export capabilities suitable for professional spectroscopic analysis workflows.

## Contributing

To contribute to this project:
1. Fork the repository
2. Create a feature branch
3. Submit a pull request with detailed description

## Changelog

### Version 1.0
- Initial release with core D-parameter analysis functionality
- Excel file integration
- Multiple smoothing algorithms
- Interactive plotting capabilities
```
