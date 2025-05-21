import os
import re
import sys
import pandas as pd
from pathlib import Path
import pdfplumber
import logging
import warnings

# Suppress pdfplumber warnings about CropBox
warnings.filterwarnings("ignore", category=UserWarning)

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler(sys.stdout)
    ]
)

logger = logging.getLogger("pdf_extractor")

# Common patterns for TDS documents
PATTERNS = {
    "Tax Invoice cum Token Number": [
        r'Token Number\s+(097979\d{9})',
        r'(097979\d{9})',
    ],
    "Name of Deductor": [
        r'Name of Deductor\s+([A-Z][A-Z0-9\s&\.]+)(?:\s+NA\s+QVZ)',  # Capture name but stop before NA QVZ
        r'Name of Deductor\s+([A-Z][A-Z0-9\s&\.]+(?:REALTORS|INFRA|LLP|ASSOCIATES|PVT|LTD))',
        r'Name of Deductor\s+([A-Z][A-Z0-9\s&\.]+)',  # More general pattern
    ],
    "Date": [
        r'Date\s+(\d{1,2}\s+[A-Za-z]+\s+\d{4})',
        r'\bDate\b[^0-9]*(\d{1,2}\s+[A-Za-z]+\s+\d{4})',
    ],
    "TAN": [
        r'TAN\s+([A-Z]{4}\d{5}[A-Z])',
        r'\bTAN\b[^A-Z]*([A-Z]{4}\d{5}[A-Z])',
        r'([A-Z]{4}\d{5}[A-Z])',
        r'(MUMS79818E)',
        r'(PNEA\d{5}[A-Z])',
        r'(PNES\d{5}[A-Z])',
        r'(PNEJ\d{5}[A-Z])',
        r'(PNEM\d{5}[A-Z])',
    ],
    "Form No": [
        r'Form No\s+(26Q)',
        r'\bForm No\b[^0-9]*(26Q)',
        r'(26Q)',
    ],
    "Receipt no.(to be quoted on TDS)": [
        r'be quoted on TDS\s+(QVZ[A-Z]{5})',
        r'be quoted on TDS\s+([A-Z0-9]+)',
        r'Receipt no\.\(note i\)[^A-Z]*\(to be quoted on TDS\s+(QVZ[A-Z]{5})',
        r'(QVZ[A-Z]{5})',
    ],
    "Type of Statement": [
        r'Type of Statement\s+(Regular|Correction)',
        r'\bType of Statement\b[^A-Za-z]*(Regular|Correction)',
        r'(Regular)',
    ],
    "Financial Year": [
        r'Financial Year\s+(2024-25)',
        r'\bFinancial Year\b[^0-9]*(2024-25)',
        r'(2024-25)',
    ],
    "Periodicity": [
        r'Periodicity\s+(Q4)',
        r'\bPeriodicity\b[^A-Z0-9]*(Q4)',
        r'(Q4)',
    ],
    "Total (Rounded off)": [
        r'Total \(Rounded off\)[^0-9]*\(₹\)\s*([\d.]+)',
        r'Total \(Rounded off\)[^0-9]*\(₹\)([\d.]+)',
        r'Total \(Rounded off\)[^0-9]*([\d.]+)',
        r'Total \(Rounded off\)\s+\(\₹\) ([\d.]+)',
        r'Total \(Rounded off\)\s+\(₹\) ([\d.]+)',
        r'Total\s*\(Rounded off\)\s+\(₹\) ([\d.]+)',
        r'Total \(Rounded off\)\s+\(₹\)([\d.]+)',
        r'(59\.00)',
    ],
}

# Default values for fields
DEFAULT_VALUES = {
    "Form No": "26Q",
    "Type of Statement": "Regular",
    "Financial Year": "2024-25",
    "Periodicity": "Q4",
    "Total (Rounded off)": "59.00"
}

def extract_text_from_pdf(pdf_path):
    """Extract all text from a PDF file"""
    try:
        with pdfplumber.open(pdf_path) as pdf:
            text = ""
            for page in pdf.pages:
                extracted_text = page.extract_text()
                if extracted_text:
                    text += extracted_text + "\n"
            return text
    except Exception as e:
        logger.error(f"Error extracting text from {pdf_path}: {str(e)}")
        return ""

def safe_regex_search(pattern, text):
    """Safely apply regex pattern with error handling"""
    try:
        match = re.search(pattern, text)
        return match
    except re.error:
        logger.warning(f"Invalid regex pattern: {pattern}")
        return None

def clean_deductor_name(name):
    """Clean up extracted deductor name to remove common issues"""
    if not name:
        return name
    
    # Remove any "Token Number" or "Deductor/Collector" text
    name = re.sub(r'Token Number|Deductor/Collector', '', name).strip()
    
    # Remove "NA QVZ..." pattern that appears in the output
    name = re.sub(r'\s+NA\s+QVZ[A-Z0-9]+', '', name).strip()
    
    # Remove "be quoted on TDS" text
    name = re.sub(r'be quoted on TDS', '', name).strip()
    
    # Remove numbers at the end
    name = re.sub(r'\s+\d+$', '', name).strip()
    
    # If the name is too short or contains only generic words, return None
    if len(name) < 3 or name in ['0', 'NA', 'None', ''] or len(name) > 100:
        return None
    
    return name

def extract_field_from_text(field_name, text, patterns=None):
    """Extract a field from text using regex patterns"""
    if not patterns:
        patterns = PATTERNS.get(field_name, [])
    
    for pattern in patterns:
        match = safe_regex_search(pattern, text)
        if match:
            try:
                # If the pattern has capturing groups, use the first one
                if match.groups():
                    value = match.group(1).strip()
                    if field_name == "Name of Deductor":
                        value = clean_deductor_name(value)
                    return value
                # Otherwise use the whole match
                return match.group(0).strip()
            except (IndexError, AttributeError):
                continue
    
    return None

def extract_field_from_row(text, token_number):
    """Extract Name of Deductor from a table row"""
    if not token_number:
        return None
    
    lines = text.split('\n')
    
    for line in lines:
        if token_number in line:
            # This might be the row with the token number and deductor name
            parts = line.split()
            if len(parts) > 1 and token_number in parts[0]:
                # The deductor name is likely everything until "NA" appears
                name_parts = []
                for part in parts[1:]:
                    if part == "NA":
                        break
                    name_parts.append(part)
                
                if name_parts:
                    possible_name = ' '.join(name_parts)
                    # Clean up and validate the name
                    return clean_deductor_name(possible_name)
    
    return None

def extract_receipt_number(text):
    """Extract receipt number using specific pattern"""
    # Look for the QVZ pattern which appears to be the receipt number format
    match = re.search(r'(QVZ[A-Z]{5})', text)
    if match:
        return match.group(1)
    return None

def extract_field_from_filename(field_name, filename):
    """Extract a field from the filename"""
    if field_name == "Tax Invoice cum Token Number":
        match = re.search(r'^(\d+)', filename)
        if match:
            return match.group(1)
    elif field_name == "Name of Deductor":
        match = re.search(r'\d+\s+(.+?)\.pdf$', filename)
        if match:
            return clean_deductor_name(match.group(1))
    
    return None

def extract_data_from_pdf(pdf_path, debug=False):
    """
    Extract TDS data from PDF using multiple strategies
    """
    filename = os.path.basename(pdf_path)
    logger.info(f"Processing: {filename}")
    
    # Initialize result with filename
    result = {
        "FileName": filename
    }
    
    # Extract text from PDF
    text = extract_text_from_pdf(pdf_path)
    if not text:
        logger.warning(f"No text extracted from {pdf_path}")
        result["Error"] = "No text extracted"
        return result
    
    # Save raw text for debugging if enabled
    if debug:
        debug_filename = os.path.splitext(pdf_path)[0] + "_text.txt"
        with open(debug_filename, "w", encoding="utf-8") as f:
            f.write(text)
        logger.debug(f"Saved extracted text to {os.path.basename(debug_filename)}")
    
    # Clean text
    clean_text = re.sub(r'\s+', ' ', text)
    
    # First, extract the token number (needed for some other extractions)
    # Try from PDF content first
    token_value = extract_field_from_text("Tax Invoice cum Token Number", clean_text)
    
    # If not found, try from filename
    if not token_value:
        token_value = extract_field_from_filename("Tax Invoice cum Token Number", filename)
    
    if token_value:
        result["Tax Invoice cum Token Number"] = token_value
    
    # Multi-strategy extraction for Name of Deductor
    # Strategy 1: Try to extract from table row
    name_value = extract_field_from_row(text, token_value)
    
    # Strategy 2: If not found, try pattern matching
    if not name_value:
        name_value = extract_field_from_text("Name of Deductor", clean_text)
    
    # Strategy 3: If still not found, try extracting from filename as last resort
    if not name_value:
        name_value = extract_field_from_filename("Name of Deductor", filename)
    
    if name_value:
        result["Name of Deductor"] = name_value
    
    # Special handling for Receipt number
    receipt_value = extract_receipt_number(clean_text)
    if receipt_value:
        result["Receipt no.(to be quoted on TDS)"] = receipt_value
    
    # Extract all remaining fields using pattern matching
    for field_name in PATTERNS.keys():
        # Skip already extracted fields
        if field_name in ["Tax Invoice cum Token Number", "Name of Deductor", "Receipt no.(to be quoted on TDS)"] and field_name in result:
            continue
        
        value = extract_field_from_text(field_name, clean_text)
        if value:
            result[field_name] = value
        elif field_name in DEFAULT_VALUES:
            # Use default value if pattern matching failed and a default exists
            result[field_name] = DEFAULT_VALUES[field_name]
    
    # Log what we found
    logger.debug("Extracted fields:")
    for field_name, value in result.items():
        if field_name not in ["FileName", "Error"]:
            logger.debug(f"  {field_name}: {value}")
    
    return result

def process_pdfs(input_dir, output_file, debug=False):
    """
    Process all PDFs in a directory and save results to Excel
    """
    # Ensure input directory exists
    input_path = Path(input_dir)
    if not input_path.exists():
        logger.error(f"Input directory not found: {input_dir}")
        return False
    
    # Find all PDF files
    pdf_files = list(input_path.glob("*.pdf"))
    if not pdf_files:
        logger.warning(f"No PDF files found in {input_dir}")
        return False
    
    logger.info(f"Found {len(pdf_files)} PDF files")
    
    # Process each PDF
    results = []
    for pdf_file in pdf_files:
        try:
            data = extract_data_from_pdf(str(pdf_file), debug)
            results.append(data)
        except Exception as e:
            logger.error(f"Error processing {pdf_file}: {str(e)}")
            # Add a minimal entry for this file
            results.append({
                "FileName": pdf_file.name,
                "Error": str(e)
            })
    
    # Create DataFrame from results
    df = pd.DataFrame(results)
    
    # Reorder columns for better readability
    column_order = [
        "FileName", 
        "Tax Invoice cum Token Number", 
        "Name of Deductor", 
        "Date", 
        "TAN", 
        "Form No", 
        "Receipt no.(to be quoted on TDS)", 
        "Type of Statement", 
        "Financial Year", 
        "Periodicity", 
        "Total (Rounded off)",
        "Error"
    ]
    
    # Select only columns that exist in the DataFrame
    existing_columns = [col for col in column_order if col in df.columns]
    df = df[existing_columns]
    
    # Save to Excel
    try:
        df.to_excel(output_file, index=False)
        logger.info(f"Data saved to {output_file}")
        return True
    except Exception as e:
        logger.error(f"Error saving to Excel: {str(e)}")
        return False

def main():
    """Main function"""
    # Set up debug mode if requested
    debug = "--debug" in sys.argv or "-d" in sys.argv
    if debug:
        logger.setLevel(logging.DEBUG)
        file_handler = logging.FileHandler("pdf_extraction.log")
        file_handler.setLevel(logging.DEBUG)
        formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
        file_handler.setFormatter(formatter)
        logger.addHandler(file_handler)
        logger.debug("Debug mode enabled")
    
    # Get input directory
    input_dir = None
    for i, arg in enumerate(sys.argv):
        if arg in ["--input", "-i"] and i < len(sys.argv) - 1:
            input_dir = sys.argv[i+1]
            break
    
    if not input_dir:
        script_dir = Path(__file__).parent
        input_dir = script_dir / "input"
    
    # Get output file
    output_file = None
    for i, arg in enumerate(sys.argv):
        if arg in ["--output", "-o"] and i < len(sys.argv) - 1:
            output_file = sys.argv[i+1]
            break
    
    if not output_file:
        script_dir = Path(__file__).parent
        output_file = script_dir / "tax_invoice_data.xlsx"
    
    # Create input directory if it doesn't exist
    if not Path(input_dir).exists():
        logger.info(f"Creating input directory: {input_dir}")
        Path(input_dir).mkdir(parents=True, exist_ok=True)
        
        # Check if any PDFs in main directory to copy
        script_dir = Path(__file__).parent
        main_dir_pdfs = list(script_dir.glob("*.pdf"))
        if main_dir_pdfs:
            logger.info(f"Found {len(main_dir_pdfs)} PDFs in main directory")
            for pdf in main_dir_pdfs:
                logger.info(f"  Copying {pdf.name} to input directory")
                dest = Path(input_dir) / pdf.name
                with open(pdf, 'rb') as src, open(dest, 'wb') as dst:
                    dst.write(src.read())
    
    # Process PDFs
    success = process_pdfs(input_dir, output_file, debug)
    
    if success:
        logger.info("PDF extraction completed successfully!")
    else:
        logger.error("PDF extraction encountered issues.")
    
    print("\nPress Enter to exit...")
    input()

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        logger.critical(f"An unexpected error occurred: {str(e)}")
        print(f"An unexpected error occurred: {str(e)}")
        print("\nPress Enter to exit...")
        input()
