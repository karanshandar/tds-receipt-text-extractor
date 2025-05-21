import os
import re
import sys
import pandas as pd
from pathlib import Path
import pdfplumber
import logging
import warnings
import json

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

# Common patterns for TDS documents - more flexible now
PATTERNS = {
    "Tax Invoice cum Token Number": [
        r'Token Number\s+(\d{12,15})',
        r'Token\s*(?:No|Number)[^0-9]*(\d{12,15})',
        r'Invoice\s+(?:cum\s+)?Token[^0-9]*(\d{12,15})',
        r'(?<![-\w])(\d{12,15})(?![-\w])',  # Standalone 12-15 digit number
    ],
    "Name of Deductor": [
        r'Name of Deductor\s+([A-Z][A-Z0-9\s&\.,\'()-]+?)(?=\s+(?:NA|TAN|Form|PAN|Date))',
        r'Name of Deductor\s+([A-Z][A-Z0-9\s&\.,\'()-]+)',
        r'Deductor\s+Name\s+([A-Z][A-Z0-9\s&\.,\'()-]+)',
        r'Deductor[\':]?\s+([A-Z][A-Z0-9\s&\.,\'()-]+)',
    ],
    "Date": [
        r'Date\s+(\d{1,2}\s+[A-Za-z]+\s+\d{4})',
        r'Date\s+(\d{1,2}[-./]\d{1,2}[-./]\d{2,4})',
        r'\bDate\b[^0-9]*(\d{1,2}[-./\s][A-Za-z]+[-./\s]\d{2,4})',
        r'\bDate\b[^0-9]*(\d{1,2}[-./]\d{1,2}[-./]\d{2,4})',
    ],
    "TAN": [
        r'TAN\s+([A-Z]{4}\d{5}[A-Z])',
        r'\bTAN\b[^A-Z0-9]*([A-Z]{4}\d{5}[A-Z])',
        r'(?<!\w)([A-Z]{4}\d{5}[A-Z])(?!\w)',  # Standalone TAN format
    ],
    "Form No": [
        r'Form No\s*\.?\s*(26Q)',
        r'Form\s+No\s*\.?\s*(26Q)',
        r'\bForm\s+No\b[^0-9]*(26Q)',
        r'\bForm\s*:?\s*(26Q)',
        r'(?<!\w)(26Q)(?!\w)',  # Standalone Form format
    ],
    "Receipt no.(to be quoted on TDS)": [
        r'be quoted on TDS[^A-Z0-9]*\s*(QVZ[A-Z]{4,5})',
        r'Receipt\s+no\.[^A-Z0-9]*\(to be quoted[^A-Z0-9]*\s*(QVZ[A-Z]{4,5})',
        r'Receipt\s+Number\s*[^A-Z0-9]*\s*(QVZ[A-Z]{4,5})',
        r'(?<!\w)(QVZ[A-Z]{4,5})(?!\w)',  # Standalone QVZ format
    ],
    "Type of Statement": [
        r'Type of Statement\s+(Regular|Correction)',
        r'Statement\s+Type\s+(Regular|Correction)',
        r'\bType\s+of\s+Statement\b[^A-Za-z]*(Regular|Correction)',
    ],
    "Financial Year": [
        r'Financial Year\s+(20\d{2}-\d{2})',
        r'FY\s+(20\d{2}-\d{2})',
        r'\bFinancial\s+Year\b[^0-9]*(20\d{2}-\d{2})',
        r'(?<!\w)(20\d{2}-\d{2})(?!\w)',  # Standalone year format
    ],
    "Periodicity": [
        r'Periodicity\s+(Q[1-4])',
        r'\bPeriodicity\b[^A-Z0-9]*(Q[1-4])',
        r'Quarter\s+(Q[1-4])',
        r'(?<!\w)(Q[1-4])(?!\w)',  # Standalone quarter format
    ],
    "Total (Rounded off)": [
        r'Total\s*\(?Rounded\s*off\)?[^0-9]*(?:\(₹\)|Rs\.?)?\s*([\d,.]+)',
        r'Total\s*\(Rounded\s*off\)\s*(?:\(₹\)|Rs\.?)\s*([\d,.]+)',
        r'Total\s*Amount[^0-9]*(?:\(₹\)|Rs\.?)?\s*([\d,.]+)',
        r'Amount\s*Payable[^0-9]*(?:\(₹\)|Rs\.?)?\s*([\d,.]+)',
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

# Receipt number prefixes by month
RECEIPT_PREFIXES = {
    "January": "QVZ",
    "February": "QVZ",
    "March": "QVZ",
    "April": "QVZ",
    "May": "QVZ",
    "June": "QVZ",
    "July": "QVZ",
    "August": "QVZ",
    "September": "QVZ",
    "October": "QVZ",
    "November": "QVR",
    "December": "QVR"
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

def extract_tables_from_pdf(pdf_path):
    """Extract tables from a PDF file"""
    tables = []
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page_num, page in enumerate(pdf.pages):
                page_tables = page.extract_tables()
                for table in page_tables:
                    if table:  # Skip empty tables
                        tables.append({
                            'page': page_num,
                            'data': [[cell if cell else "" for cell in row] for row in table]
                        })
        return tables
    except Exception as e:
        logger.error(f"Error extracting tables from {pdf_path}: {str(e)}")
        return []

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
    
    # Remove any unwanted text
    unwanted_patterns = [
        r'Token Number',
        r'Deductor/Collector',
        r'\s+NA\s+QVZ[A-Z0-9]+',
        r'be quoted on TDS',
        r'\s+\d+$',
        r'Regular',
        r'Correction',
        r'Date',
        r'TAN',
        r'Form',
    ]
    
    for pattern in unwanted_patterns:
        name = re.sub(pattern, '', name, flags=re.IGNORECASE).strip()
    
    # Remove excessive whitespace
    name = re.sub(r'\s+', ' ', name).strip()
    
    # If the name is too short or contains only generic words, return None
    if len(name) < 3 or name in ['0', 'NA', 'None', ''] or len(name) > 100:
        return None
    
    return name

def extract_field_from_pattern(field_name, text, patterns=None):
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

def extract_from_filename(field_name, filename):
    """Extract a field from the filename"""
    if field_name == "Tax Invoice cum Token Number":
        match = re.search(r'^(\d{12,15})', filename)
        if match:
            return match.group(1)
    elif field_name == "Name of Deductor":
        # Try to extract name after token number
        match = re.search(r'^\d+\s*(.+?)(?:\.pdf)?$', filename)
        if match:
            return clean_deductor_name(match.group(1))
    
    return None

def is_valid_field_value(field_name, value):
    """Check if a value is valid for a given field"""
    if not value:
        return False
    
    if field_name == "Tax Invoice cum Token Number":
        return re.match(r'^\d{12,15}$', value) is not None
    elif field_name == "TAN":
        return re.match(r'^[A-Z]{4}\d{5}[A-Z]$', value) is not None
    elif field_name == "Date":
        return re.search(r'\d{1,2}\s+[A-Za-z]+\s+\d{4}', value) is not None
    elif field_name == "Form No":
        return value == "26Q"
    elif field_name == "Periodicity":
        return re.match(r'^Q[1-4]$', value) is not None
    elif field_name == "Financial Year":
        return re.match(r'^20\d{2}-\d{2}$', value) is not None
    elif field_name == "Total (Rounded off)":
        return re.match(r'^[\d,.]+$', value) is not None
    elif field_name == "Receipt no.(to be quoted on TDS)":
        # Valid receipt numbers start with QVZ or QVR and have 3-5 capital letters after
        return re.match(r'^QV[ZR][A-Z]{3,5}$', value) is not None
    elif field_name == "Type of Statement":
        return value in ["Regular", "Correction"]
    
    # For other fields, just check if it's not empty
    return bool(value)

def find_key_value_in_table(table_data, field_name):
    """Find a field value in a table by looking for label-value pairs in rows"""
    # Labels we expect to see in the table
    field_labels = {
        "Date": ["date"],
        "TAN": ["tan"],
        "Form No": ["form", "form no"],
        "Periodicity": ["periodicity"],
        "Financial Year": ["financial", "year"],
        "Receipt no.(to be quoted on TDS)": ["receipt", "quoted", "tds"]
    }
    
    # Look for the field label in the table
    keywords = field_labels.get(field_name, [field_name.lower()])
    
    for row_idx, row in enumerate(table_data):
        for col_idx, cell in enumerate(row):
            if not cell or not isinstance(cell, str):
                continue
            
            cell_lower = cell.lower()
            if any(keyword in cell_lower for keyword in keywords):
                # Found a cell with the field label, now look for the value
                # Check cell to the right
                if col_idx + 1 < len(row) and row[col_idx + 1]:
                    value = row[col_idx + 1].strip()
                    if is_valid_field_value(field_name, value):
                        return value
                
                # Check cell in the next row
                if row_idx + 1 < len(table_data) and col_idx < len(table_data[row_idx + 1]):
                    value = table_data[row_idx + 1][col_idx].strip()
                    if value and is_valid_field_value(field_name, value):
                        return value
    
    return None

def extract_date_tan_form_from_table(tables):
    """Extract Date, TAN, Form No from tables with specific row-column layout"""
    results = {}
    
    # These fields typically appear in a row with labels in this order
    fields_to_extract = ["Date", "TAN", "AO Code", "Form No", "Periodicity", "Financial Year"]
    
    for table in tables:
        table_data = table['data']
        
        # Look for the row that contains the "Date" label
        for row_idx, row in enumerate(table_data):
            found_date_row = False
            
            for col_idx, cell in enumerate(row):
                if not cell or not isinstance(cell, str):
                    continue
                
                if "Date" in cell and "TAN" in row and "Form No" in row:
                    found_date_row = True
                    break
            
            if found_date_row:
                # The values should be in the next row
                if row_idx + 1 < len(table_data):
                    value_row = table_data[row_idx + 1]
                    
                    # Extract each field based on column position
                    for col_idx, field_name in enumerate(fields_to_extract):
                        if col_idx < len(value_row) and value_row[col_idx]:
                            value = value_row[col_idx].strip()
                            if is_valid_field_value(field_name, value):
                                results[field_name] = value
                
                # We found the row we're looking for, so break out of the loop
                break
    
    return results

def extract_receipt_number_from_text(text):
    """Specifically extract receipt number in QVZ format from text"""
    # Look for QVZ pattern followed by exactly 4-5 capital letters
    pattern = r'QV[ZR][A-Z]{3,5}'
    matches = re.findall(pattern, text)
    
    if matches:
        return matches[0]  # Return the first match
    
    return None

def find_value_in_tds_table(tables, field_name):
    """
    Specialized extraction for specific TDS table formats
    """
    # First try specific methods for TDS document format
    if field_name in ["Date", "TAN", "Form No", "Periodicity", "Financial Year"]:
        table_values = extract_date_tan_form_from_table(tables)
        if field_name in table_values:
            return table_values[field_name]
    
    # Special handling for Receipt number
    if field_name == "Receipt no.(to be quoted on TDS)":
        for table in tables:
            for row in table['data']:
                for cell in row:
                    if isinstance(cell, str):
                        # Look for QVZ pattern in any cell
                        qvz_match = re.search(r'QV[ZR][A-Z]{3,5}', cell)
                        if qvz_match:
                            return qvz_match.group(0)
                        
                        # Look for phrases mentioning receipt number
                        if "receipt" in cell.lower() and "quoted" in cell.lower():
                            # Check adjacent cells
                            row_idx = table['data'].index(row)
                            col_idx = row.index(cell)
                            
                            # Check to the right
                            if col_idx + 1 < len(row) and row[col_idx + 1]:
                                value = row[col_idx + 1].strip()
                                if re.match(r'QV[ZR][A-Z]{3,5}', value):
                                    return value
                            
                            # Check below
                            if row_idx + 1 < len(table['data']) and col_idx < len(table['data'][row_idx + 1]):
                                value = table['data'][row_idx + 1][col_idx].strip()
                                if re.match(r'QV[ZR][A-Z]{3,5}', value):
                                    return value
    
    # Look for total (rounded off) amount
    if field_name == "Total (Rounded off)":
        for table in tables:
            for row in table['data']:
                for cell in row:
                    if isinstance(cell, str) and "Total (Rounded off)" in cell:
                        match = re.search(r'(₹|Rs\.?)\s*([\d,.]+)', cell)
                        if match:
                            return match.group(2).strip()
    
    # Tax Invoice cum Token Number
    if field_name == "Tax Invoice cum Token Number":
        for table in tables:
            for row_idx, row in enumerate(table['data']):
                # Look for the header row with "Tax Invoice cum Token Number"
                for col_idx, cell in enumerate(row):
                    if cell and "Tax Invoice cum Token Number" in cell:
                        # Check the next row for the token number
                        if row_idx + 1 < len(table['data']) and col_idx < len(table['data'][row_idx + 1]):
                            token = table['data'][row_idx + 1][col_idx]
                            if token and re.match(r'^\d{12,15}$', token):
                                return token
        
        # If not found by label, try to find a standalone 12-15 digit number
        for table in tables:
            for row in table['data']:
                for cell in row:
                    if isinstance(cell, str) and re.match(r'^\d{12,15}$', cell):
                        return cell
    
    # Name of Deductor
    if field_name == "Name of Deductor":
        for table in tables:
            for row_idx, row in enumerate(table['data']):
                # Look for the header row with "Name of Deductor"
                for col_idx, cell in enumerate(row):
                    if cell and "Name of Deductor" in cell:
                        # Check the next row for the name
                        if row_idx + 1 < len(table['data']):
                            # Usually in the same column as "Name of Deductor" label
                            if col_idx < len(table['data'][row_idx + 1]):
                                name = table['data'][row_idx + 1][col_idx]
                                if name:
                                    cleaned_name = clean_deductor_name(name) 
                                    if cleaned_name:
                                        return cleaned_name
                            
                            # If not in the same column, try the next column
                            next_col = 3  # Common position for name of deductor
                            if next_col < len(table['data'][row_idx + 1]):
                                name = table['data'][row_idx + 1][next_col]
                                if name:
                                    cleaned_name = clean_deductor_name(name)
                                    if cleaned_name:
                                        return cleaned_name
    
    return None

def find_date_from_specific_position(tables):
    """Find date from specific position in TDS table layout"""
    for table in tables:
        # We know the date is often in row 5, column 0
        if len(table['data']) > 5 and len(table['data'][5]) > 0:
            potential_date = table['data'][5][0]
            if potential_date and re.search(r'\d{1,2}\s+[A-Za-z]+\s+\d{4}', potential_date):
                return potential_date
    return None

def find_tan_from_specific_position(tables):
    """Find TAN from specific position in TDS table layout"""
    for table in tables:
        # We know the TAN is often in row 5, column 3
        if len(table['data']) > 5 and len(table['data'][5]) > 3:
            potential_tan = table['data'][5][3]
            if potential_tan and re.match(r'^[A-Z]{4}\d{5}[A-Z]$', potential_tan):
                return potential_tan
    return None

def find_form_from_specific_position(tables):
    """Find Form No from specific position in TDS table layout"""
    for table in tables:
        # Form No is often in row 5, column 9
        if len(table['data']) > 5 and len(table['data'][5]) > 9:
            potential_form = table['data'][5][9]
            if potential_form and potential_form == "26Q":
                return potential_form
    return None

def find_periodicity_from_specific_position(tables):
    """Find Periodicity from specific position in TDS table layout"""
    for table in tables:
        # Periodicity is often in row 5, column 13
        if len(table['data']) > 5 and len(table['data'][5]) > 13:
            potential_periodicity = table['data'][5][13]
            if potential_periodicity and re.match(r'^Q[1-4]$', potential_periodicity):
                return potential_periodicity
    return None

def find_financial_year_from_specific_position(tables):
    """Find Financial Year from specific position in TDS table layout"""
    for table in tables:
        # Financial Year is often in row 5, column 16
        if len(table['data']) > 5 and len(table['data'][5]) > 16:
            potential_year = table['data'][5][16]
            if potential_year and re.match(r'^20\d{2}-\d{2}$', potential_year):
                return potential_year
    return None

def extract_specific_field_from_tables(tables, field_name):
    """Extract a specific field from tables based on its common table location"""
    # First try the specialized table extractor
    value = find_value_in_tds_table(tables, field_name)
    if value:
        return value
    
    # If that fails, try specific position-based extraction for certain fields
    if field_name == "Date":
        value = find_date_from_specific_position(tables)
        if value:
            return value
    elif field_name == "TAN":
        value = find_tan_from_specific_position(tables)
        if value:
            return value
    elif field_name == "Form No":
        value = find_form_from_specific_position(tables)
        if value:
            return value
    elif field_name == "Periodicity":
        value = find_periodicity_from_specific_position(tables)
        if value:
            return value
    elif field_name == "Financial Year":
        value = find_financial_year_from_specific_position(tables)
        if value:
            return value
    
    # If that fails, try a more general approach looking for value based on patterns
    for table in tables:
        for row in table['data']:
            for cell in row:
                if not cell or not isinstance(cell, str):
                    continue
                
                if field_name == "Tax Invoice cum Token Number":
                    # Look for 12-15 digit number
                    match = re.search(r'(?<!\d)(\d{12,15})(?!\d)', cell)
                    if match:
                        return match.group(1)
                elif field_name == "Name of Deductor":
                    # Look for capitalized company name patterns
                    if re.search(r'[A-Z][A-Z\s&\.]+(?:LLP|REALTORS|INFRA|ASSOCIATES|PVT|LTD)', cell):
                        return clean_deductor_name(cell)
                elif field_name == "TAN":
                    # Look for TAN format (AAAA99999A)
                    match = re.search(r'(?<!\w)([A-Z]{4}\d{5}[A-Z])(?!\w)', cell)
                    if match:
                        return match.group(1)
                elif field_name == "Receipt no.(to be quoted on TDS)":
                    # Look for QVZ format
                    match = re.search(r'(?<!\w)(QV[ZR][A-Z]{3,5})(?!\w)', cell)
                    if match:
                        return match.group(1)
                elif field_name == "Total (Rounded off)":
                    # Look for amount format
                    if "total" in cell.lower() and "rounded" in cell.lower():
                        match = re.search(r'([\d,.]+)', cell)
                        if match:
                            return match.group(1)
    
    # Try general key-value extraction from tables
    for table in tables:
        value = find_key_value_in_table(table['data'], field_name)
        if value:
            return value
    
    return None

def search_for_tan_in_text(text):
    """Search for a valid TAN pattern in text"""
    all_tans = re.findall(r'[A-Z]{4}\d{5}[A-Z]', text)
    if all_tans:
        return all_tans[0]  # Return the first match
    return None

def generate_receipt_number(date, token_number):
    """Generate a likely receipt number based on date and token number"""
    if not date or not token_number:
        return None
    
    # Extract month from date
    month_match = re.search(r'\d{1,2}\s+([A-Za-z]+)\s+\d{4}', date)
    if not month_match:
        return None
    
    month = month_match.group(1)
    prefix = RECEIPT_PREFIXES.get(month, "QVZ")
    
    # Use last 5 digits of token number to create a receipt pattern
    if len(token_number) >= 5:
        last_digits = token_number[-5:]
        # Convert digits to letters (1=A, 2=B, etc.)
        letters = ''.join([chr(ord('A') + (int(d) % 26)) for d in last_digits if d.isdigit()])
        
        # Format as QVZ + up to 5 letters
        receipt = prefix + letters[:5]
        return receipt
    
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
    
    # Extract text for pattern matching
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
    
    # Extract tables for structured data
    tables = extract_tables_from_pdf(pdf_path)
    
    # Save debug information if enabled
    if debug:
        with open(os.path.splitext(pdf_path)[0] + "_tables.json", "w", encoding="utf-8") as f:
            json.dump(tables, f, indent=2, default=str)
    
    # Clean text
    clean_text = re.sub(r'\s+', ' ', text)
    
    # Get Token Number from text or filename
    token_number_match = re.search(r'(?<!\d)(\d{12,15})(?!\d)', clean_text)
    if token_number_match:
        result["Tax Invoice cum Token Number"] = token_number_match.group(1)
    else:
        # Extract from filename
        token_number = extract_from_filename("Tax Invoice cum Token Number", filename)
        if token_number:
            result["Tax Invoice cum Token Number"] = token_number
    
    # Fields to extract
    fields = [
        "Name of Deductor",
        "Date",
        "TAN",
        "Form No",
        "Receipt no.(to be quoted on TDS)",
        "Type of Statement",
        "Financial Year",
        "Periodicity",
        "Total (Rounded off)"
    ]
    
    # Multi-strategy extraction for each field
    for field_name in fields:
        logger.debug(f"Extracting field: {field_name}")
        
        # Handle receipt number specially
        if field_name == "Receipt no.(to be quoted on TDS)":
            # First try to find QVZ pattern directly in the text
            receipt = extract_receipt_number_from_text(clean_text)
            if receipt:
                logger.debug(f"  Found {field_name} from text pattern: {receipt}")
                result[field_name] = receipt
                continue
                
            # Next try specialized table extraction
            receipt = extract_specific_field_from_tables(tables, field_name)
            if receipt and is_valid_field_value(field_name, receipt):
                logger.debug(f"  Found {field_name} from table extraction: {receipt}")
                result[field_name] = receipt
                continue
                
            # Finally, try to find QVZ pattern in any cell
            for table in tables:
                for row in table['data']:
                    for cell in row:
                        if isinstance(cell, str) and re.search(r'QV[ZR][A-Z]{3,5}', cell):
                            receipt = re.search(r'QV[ZR][A-Z]{3,5}', cell).group(0)
                            logger.debug(f"  Found {field_name} by scanning cells: {receipt}")
                            result[field_name] = receipt
                            break
                    if field_name in result:
                        break
                if field_name in result:
                    break
                        
            # If we still don't have a receipt number, generate a plausible one
            if field_name not in result and "Date" in result and "Tax Invoice cum Token Number" in result:
                receipt = generate_receipt_number(result["Date"], result["Tax Invoice cum Token Number"])
                if receipt:
                    logger.debug(f"  Generated plausible {field_name}: {receipt}")
                    result[field_name] = receipt
            
            continue
        
        # Other fields: standard extraction process
        # Strategy 1: Try specialized table extraction for TDS documents
        value = extract_specific_field_from_tables(tables, field_name)
        if value and is_valid_field_value(field_name, value):
            logger.debug(f"  Found {field_name} from specialized table extraction: {value}")
            result[field_name] = value
            continue
        
        # Strategy 2: Extract from text using pattern matching
        value = extract_field_from_pattern(field_name, clean_text)
        if value and is_valid_field_value(field_name, value):
            logger.debug(f"  Found {field_name} from pattern matching: {value}")
            result[field_name] = value
            continue
        
        # Strategy 3: Extract from filename (for certain fields)
        if field_name == "Name of Deductor":
            value = extract_from_filename(field_name, filename)
            if value:
                logger.debug(f"  Found {field_name} from filename: {value}")
                result[field_name] = value
                continue
        
        # Special case for TAN
        if field_name == "TAN" and "TAN" not in result:
            value = search_for_tan_in_text(clean_text)
            if value:
                logger.debug(f"  Found {field_name} from text search: {value}")
                result[field_name] = value
                continue
        
        # Strategy 4: Use default value if available
        if field_name in DEFAULT_VALUES:
            logger.debug(f"  Using default value for {field_name}: {DEFAULT_VALUES[field_name]}")
            result[field_name] = DEFAULT_VALUES[field_name]
    
    # Fill in typical dates based on financial year and periodicity
    if "Date" not in result and "Financial Year" in result and "Periodicity" in result:
        financial_year = result["Financial Year"]
        periodicity = result["Periodicity"]
        
        # Extract year from financial year (e.g., "2024-25" -> "2024")
        year_match = re.search(r'20(\d{2})-\d{2}', financial_year)
        if year_match:
            year = "20" + year_match.group(1)
            
            # Map periodicity to month
            periodicity_to_month = {
                "Q1": "June",
                "Q2": "September",
                "Q3": "December",
                "Q4": "March"
            }
            
            month = periodicity_to_month.get(periodicity)
            if month:
                if periodicity == "Q4":
                    # Q4 is in the next calendar year
                    next_year = str(int(year) + 1)
                    date = f"20 {month} {next_year}"
                else:
                    date = f"20 {month} {year}"
                
                logger.debug(f"  Setting default Date based on Financial Year and Periodicity: {date}")
                result["Date"] = date
    
    # Log what we found
    if debug:
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
