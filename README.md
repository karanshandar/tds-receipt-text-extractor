# PDF Data Extractor

A simple, reliable tool for extracting structured data from TDS (Tax Deducted at Source) PDF documents.

## Features

- Extracts key fields from TDS statements
- Multiple extraction strategies for reliability
- Debug mode for troubleshooting
- Excel output for easy data analysis
- Simple command-line interface

## Quick Start

1. Place your PDF files in the `input` folder
2. Run the batch file:
   ```
   run_extractor.bat
   ```
3. Find the extracted data in `tax_invoice_data.xlsx`

## Installation

### Prerequisites

- Python 3.7 or newer
- pip (Python package installer)

### Installing Dependencies

```bash
pip install -r requirements.txt
```

Or run the included batch file which will check and install required dependencies:

```bash
run_extractor.bat
```

## Usage

### Basic Usage

Simply run the batch file:
```
run_extractor.bat
```

### Debug Mode

For troubleshooting, run in debug mode:
```
run_extractor.bat --debug
```

This will:
- Create text files with the extracted content from each PDF
- Generate a detailed log file `pdf_extraction.log`

### Command Line Options

The Python script supports the following command line options:

- `--input PATH` or `-i PATH`: Input directory containing PDF files
- `--output PATH` or `-o PATH`: Output Excel file
- `--debug` or `-d`: Enable debug mode

Example:
```
python pdf_extractor.py --input "path/to/pdfs" --output "results.xlsx" --debug
```

## How It Works

The tool uses a multi-strategy approach to extract data:

1. First extracts the "Tax Invoice cum Token Number" (which is needed for other extractions)
2. For the "Name of Deductor":
   - First tries to extract from table rows containing the token number
   - Then tries pattern matching on the entire document
   - Finally falls back to extracting from the filename if needed
3. For all other fields:
   - Uses pattern matching with multiple regex patterns
   - Falls back to default values for certain fields if needed

This provides maximum reliability when working with various PDF formats.

## Extracted Fields

The tool extracts the following fields:

- Tax Invoice cum Token Number
- Name of Deductor
- Date
- TAN
- Form No
- Receipt no.(to be quoted on TDS)
- Type of Statement
- Financial Year
- Periodicity
- Total (Rounded off)

## Troubleshooting

If extraction isn't working correctly:

1. Run in debug mode: `run_extractor.bat --debug`
2. Check the generated text files to see what text was extracted from each PDF
3. Review the log file `pdf_extraction.log` for details on what happened
4. If needed, update the regex patterns in the `PATTERNS` dictionary in the script

## Project Structure

```
pdf-extractor/
├── pdf_extractor.py     # Main Python script
├── run_extractor.bat    # Windows batch file for easy running
├── requirements.txt     # Python dependencies
├── README.md            # This documentation
├── LICENSE              # MIT License
├── input/               # Place PDF files here
└── examples/            # Example PDF and output
```

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add some amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Acknowledgements

- [pdfplumber](https://github.com/jsvine/pdfplumber) for PDF text extraction
- [pandas](https://pandas.pydata.org/) for data manipulation
- [openpyxl](https://openpyxl.readthedocs.io/) for Excel file generation
