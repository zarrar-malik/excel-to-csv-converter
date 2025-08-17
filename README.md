# Excel to CSV Converter

![Python Version](https://img.shields.io/badge/python-3.7%2B-blue)
![License](https://img.shields.io/badge/license-MIT-green)

A command-line tool to convert Excel files (.xlsx/.xls) to CSV format with automatic email column detection.

## Features ✨

- Automatically detects columns containing email addresses
- Supports both single files and batch directory processing
- Preserves all columns or extracts only email addresses
- Verbose mode for debugging
- Clean, type-annotated Python code

## Installation 📦

1. Clone the repository:
   ```bash
   git clone https://github.com/yourusername/excel-to-csv-converter.git
   cd excel-to-csv-converter
   ```

2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

## Usage 🚀

### Basic Conversion
```bash
python converter.py -i input.xlsx -o output.csv
```

### Batch Process a Directory
```bash
python converter.py -i input_folder -o output_folder
```

### All Available Options
```bash
python converter.py --help
```

## Examples 📋

1. Convert single file (email columns only):
   ```bash
   python converter.py -i data/contacts.xlsx -o converted/emails.csv
   ```

2. Convert all Excel files in a folder (keep all columns):
   ```bash
   python converter.py -i data/ -o converted/ --all-columns
   ```

## File Structure 📂
```
excel-to-csv-converter/
├── converter.py       # Main conversion script
├── requirements.txt   # Dependency list
├── LICENSE            # MIT License
└── README.md          # This documentation
```

## Contributing 🤝
Pull requests are welcome! Please ensure:
1. Your code passes `flake8` linting
2. Add tests for new features
3. Update documentation as needed

## License 📜
This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.