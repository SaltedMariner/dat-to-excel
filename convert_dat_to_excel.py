#!/usr/bin/env python3
"""
Interactive Sybase .dat File to Excel Converter
Works with double-click or drag-and-drop
"""

import pandas as pd
import numpy as np
import chardet
import os
import sys
import logging
from datetime import datetime
import struct
import re

# Set up logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)


class SybaseDatReader:
    """Class to handle reading various Sybase .dat file formats"""
    
    def __init__(self, filepath, encoding=None):
        self.filepath = filepath
        self.encoding = encoding
        self.data = None
        self.detected_format = None
        
    def detect_encoding(self, sample_size=10000):
        """Detect file encoding"""
        try:
            with open(self.filepath, 'rb') as f:
                raw_data = f.read(sample_size)
                result = chardet.detect(raw_data)
                detected = result['encoding']
                confidence = result['confidence']
                
                logger.info(f"Detected encoding: {detected} (confidence: {confidence:.2f})")
                
                if detected and confidence > 0.7:
                    return detected
                else:
                    return 'latin-1'
        except Exception as e:
            logger.warning(f"Could not detect encoding: {e}")
            return 'latin-1'
    
    def detect_delimiter(self, sample_lines=10):
        """Detect delimiter in the file"""
        try:
            delimiters = ['\t', '|', ',', ';', '\x01']
            delimiter_counts = {d: 0 for d in delimiters}
            
            with open(self.filepath, 'r', encoding=self.encoding) as f:
                for i, line in enumerate(f):
                    if i >= sample_lines:
                        break
                    for delimiter in delimiters:
                        delimiter_counts[delimiter] += line.count(delimiter)
            
            best_delimiter = max(delimiter_counts, key=delimiter_counts.get)
            if delimiter_counts[best_delimiter] > 0:
                logger.info(f"Detected delimiter: {repr(best_delimiter)}")
                return best_delimiter
            else:
                return None
        except Exception as e:
            logger.warning(f"Could not detect delimiter: {e}")
            return None
    
    def read_delimited(self, delimiter=None, **kwargs):
        """Read delimited file (CSV, TSV, etc.)"""
        try:
            if delimiter is None:
                delimiter = self.detect_delimiter()
                if delimiter is None:
                    delimiter = '\t'
            
            default_params = {
                'delimiter': delimiter,
                'encoding': self.encoding,
                'quoting': 3,
                'on_bad_lines': 'skip',
                'low_memory': False
            }
            
            default_params.update(kwargs)
            
            logger.info(f"Reading delimited file with delimiter: {repr(delimiter)}")
            self.data = pd.read_csv(self.filepath, **default_params)
            self.detected_format = 'delimited'
            
            return self.data
            
        except Exception as e:
            logger.error(f"Failed to read as delimited file: {e}")
            raise
    
    def read_fixed_width(self, widths=None, columns=None):
        """Read fixed-width file"""
        try:
            logger.info("Reading as fixed-width file")
            
            if widths is None:
                widths = self._detect_fixed_widths()
            
            self.data = pd.read_fwf(
                self.filepath,
                widths=widths,
                names=columns,
                encoding=self.encoding
            )
            self.detected_format = 'fixed_width'
            
            return self.data
            
        except Exception as e:
            logger.error(f"Failed to read as fixed-width file: {e}")
            raise
    
    def auto_read(self):
        """Automatically detect and read the file"""
        logger.info(f"Auto-detecting format for: {self.filepath}")
        
        if self.encoding is None:
            self.encoding = self.detect_encoding()
        
        methods = [
            (self.read_delimited, {}),
            (self.read_fixed_width, {}),
        ]
        
        for method, params in methods:
            try:
                return method(**params)
            except Exception as e:
                logger.debug(f"Method {method.__name__} failed: {e}")
                continue
        
        raise ValueError("Could not read file with any known format")
    
    def _detect_fixed_widths(self, sample_lines=100):
        """Try to detect fixed widths by analyzing spaces"""
        try:
            lines = []
            with open(self.filepath, 'r', encoding=self.encoding) as f:
                for i, line in enumerate(f):
                    if i >= sample_lines:
                        break
                    lines.append(line.rstrip('\n'))
            
            if not lines:
                return None
            
            space_positions = []
            for pos in range(len(max(lines, key=len))):
                if all(pos < len(line) and line[pos] == ' ' for line in lines):
                    space_positions.append(pos)
            
            if space_positions:
                widths = []
                prev = 0
                for pos in space_positions:
                    if pos - prev > 1:
                        widths.append(pos - prev)
                        prev = pos
                widths.append(None)
                return widths
            
            return None
            
        except Exception:
            return None


def clean_data(df):
    """Clean and prepare data for Excel export - basic cleaning"""
    # Remove leading/trailing whitespace
    for col in df.columns:
        if df[col].dtype == 'object':
            df[col] = df[col].str.strip()
    
    # Replace empty strings with NaN
    df = df.replace('', np.nan)
    
    # Try to convert numeric columns
    for col in df.columns:
        try:
            if pd.api.types.is_numeric_dtype(df[col]):
                continue
                
            numeric_series = pd.to_numeric(df[col], errors='coerce')
            
            if numeric_series.notna().sum() / len(df) > 0.5:
                df[col] = numeric_series
        except:
            pass
    
    return df


def clean_for_excel(df):
    """Remove all special characters that Excel doesn't like"""
    
    # Clean column names first
    new_columns = []
    for col in df.columns:
        # Convert to string and clean
        clean_col = str(col)
        # Replace special characters with underscore
        clean_col = re.sub(r'[^a-zA-Z0-9_]', '_', clean_col)
        # Remove leading/trailing underscores
        clean_col = clean_col.strip('_')
        # Handle empty column names
        if not clean_col:
            clean_col = 'Column'
        new_columns.append(clean_col)
    
    df.columns = new_columns
    
    # Clean data in each column
    for col in df.columns:
        if df[col].dtype == 'object':
            # Convert to string to ensure consistent handling
            df[col] = df[col].astype(str)
            
            # Remove null bytes (most common issue)
            df[col] = df[col].str.replace('\x00', '', regex=False)
            
            # Remove control characters
            df[col] = df[col].str.replace('[\x01-\x1f\x7f-\x9f]', '', regex=True)
            
            # Fix various quote types
            df[col] = df[col].str.replace('"', '"', regex=False)
            df[col] = df[col].str.replace('"', '"', regex=False)
            df[col] = df[col].str.replace(''', "'", regex=False)
            df[col] = df[col].str.replace(''', "'", regex=False)
            
            # Replace other problematic characters
            df[col] = df[col].str.replace('–', '-', regex=False)  # En dash
            df[col] = df[col].str.replace('—', '-', regex=False)  # Em dash
            df[col] = df[col].str.replace('…', '...', regex=False)  # Ellipsis
            
            # Clean up whitespace
            df[col] = df[col].str.strip()
            
            # Replace 'nan' strings with actual NaN
            df[col] = df[col].replace('nan', np.nan)
    
    return df


def export_to_excel(df, output_path, sheet_name='Data'):
    """Export DataFrame to Excel with comprehensive error handling"""
    try:
        # First apply special character cleaning
        print("Cleaning special characters...")
        df = clean_for_excel(df)
        
        # Try to export with openpyxl
        print(f"Exporting to Excel: {output_path}")
        df.to_excel(output_path, index=False, sheet_name=sheet_name)
        
        print(f"✅ Successfully exported to: {output_path}")
        print(f"📊 Rows: {len(df)}, Columns: {len(df.columns)}")
        
    except Exception as e:
        print(f"⚠️  Excel export error: {e}")
        
        # Try without sheet name
        try:
            print("Trying simpler export...")
            df.to_excel(output_path, index=False)
            print(f"✅ Export successful (basic format)")
        except:
            # Fall back to CSV
            csv_path = output_path.replace('.xlsx', '.csv')
            print(f"Excel export failed. Saving as CSV: {csv_path}")
            df.to_csv(csv_path, index=False, encoding='utf-8-sig')
            print(f"✅ Saved as CSV instead: {csv_path}")
            print("You can open this CSV in Excel and save as .xlsx")


def interactive_mode():
    """Interactive mode for double-click execution"""
    print("=" * 60)
    print("     SYBASE DAT TO EXCEL CONVERTER - INTERACTIVE MODE")
    print("=" * 60)
    print()
    
    # Check if file was dragged onto script
    if len(sys.argv) > 1 and os.path.exists(sys.argv[1]):
        input_file = sys.argv[1]
        print(f"Processing dropped file: {input_file}")
    else:
        # Ask for input file
        print("Enter the path to your .dat file:")
        print("TIP: You can drag and drop the file here from Windows Explorer")
        print()
        
        while True:
            input_file = input("Input file path: ").strip().strip('"').strip("'")
            
            # Handle Windows paths
            input_file = input_file.replace('\\', '/')
            
            if os.path.exists(input_file):
                print(f"✅ File found!")
                break
            else:
                print(f"❌ File not found: {input_file}")
                print("Please check the path and try again.")
                print()
    
    # Generate output filename
    output_file = os.path.splitext(input_file)[0] + '.xlsx'
    
    # Ask if user wants to change output name
    print()
    print(f"Output will be saved as: {output_file}")
    change = input("Press Enter to accept, or type a new name: ").strip()
    
    if change:
        if not change.endswith('.xlsx'):
            change += '.xlsx'
        output_file = os.path.join(os.path.dirname(input_file), change)
    
    # Process options
    print()
    print("Conversion Options:")
    print("1. Auto-detect format (recommended)")
    print("2. Specify delimiter")
    print("3. Fixed-width format")
    
    choice = input("Select option (1-3) [default: 1]: ").strip() or '1'
    
    # Create reader
    reader = SybaseDatReader(input_file)
    
    try:
        print()
        print("Processing file...")
        print("-" * 40)
        
        if choice == '1':
            df = reader.auto_read()
        elif choice == '2':
            delimiter = input("Enter delimiter (tab, pipe, comma): ").strip().lower()
            delimiter_map = {
                'tab': '\t',
                'pipe': '|',
                'comma': ',',
                '\\t': '\t'
            }
            delimiter = delimiter_map.get(delimiter, delimiter)
            df = reader.read_delimited(delimiter=delimiter)
        elif choice == '3':
            widths_str = input("Enter column widths (comma-separated): ").strip()
            widths = [int(w.strip()) for w in widths_str.split(',')]
            df = reader.read_fixed_width(widths=widths)
        
        # Show data info
        print(f"\n📊 Data loaded: {len(df)} rows × {len(df.columns)} columns")
        print(f"📝 Column names: {', '.join(df.columns[:5])}" + 
              ("..." if len(df.columns) > 5 else ""))
        
        # Clean data (basic cleaning)
        print("\nCleaning data...")
        df = clean_data(df)
        
        # Export to Excel (includes special character cleaning)
        print("\nExporting to Excel...")
        export_to_excel(df, output_file)
        
        print()
        print("=" * 60)
        print(f"✅ SUCCESS! File converted successfully")
        print(f"📁 Output saved to: {output_file}")
        print("=" * 60)
        
        # Open file location option
        if sys.platform == 'win32':
            open_folder = input("\nOpen output folder? (y/n): ").strip().lower()
            if open_folder == 'y':
                os.startfile(os.path.dirname(os.path.abspath(output_file)))
        
    except Exception as e:
        print()
        print("=" * 60)
        print(f"❌ ERROR: {str(e)}")
        print("=" * 60)
        print()
        print("Troubleshooting tips:")
        print("1. Make sure the file is not open in another program")
        print("2. Check if you have write permissions to the folder")
        print("3. Try copying the file to a simpler path (e.g., C:\\temp\\)")
        print("4. For OneDrive files, ensure they're downloaded locally")
        print("\nTechnical details:")
        import traceback
        traceback.print_exc()


def main():
    """Main entry point"""
    try:
        # Always run in interactive mode for this version
        interactive_mode()
    except Exception as e:
        print(f"\n❌ Unexpected error: {str(e)}")
    finally:
        print()
        input("Press Enter to exit...")


if __name__ == "__main__":
    main()