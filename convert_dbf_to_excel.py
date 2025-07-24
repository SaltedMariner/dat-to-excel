import pandas as pd
import os
import sys
from datetime import datetime

def convert_dbf_properly(dbf_file, excel_file=None):
    """Convert DBF with proper data type handling"""
    
    try:
        from dbfread import DBF
    except ImportError:
        print("❌ dbfread not installed. Please install it:")
        print("   pip install dbfread")
        return None
    
    if excel_file is None:
        excel_file = dbf_file.replace('.dat', '.xlsx').replace('.dbf', '.xlsx')
    
    try:
        print(f"Opening DBF file: {dbf_file}")
        print("=" * 60)
        
        # Open DBF file with various compatibility options
        table = DBF(dbf_file, 
                    encoding='latin-1',  # Try different encodings if needed
                    lowernames=False,    # Keep original case
                    ignore_missing_memofile=True,
                    char_decode_errors='ignore')  # Ignore decode errors
        
        # Get info about the file
        print(f"DBF Information:")
        print(f"- Encoding: {table.encoding}")
        print(f"- Field count: {len(table.field_names)}")
        print(f"- Field names: {', '.join(table.field_names[:10])}...")
        
        # Read all records
        print(f"\nReading records...")
        records = []
        errors = []
        
        for i, record in enumerate(table):
            if i % 1000 == 0:
                print(f"  Processing record {i}...")
            
            try:
                # Convert record to dict, handling any conversion issues
                record_dict = {}
                for field in table.field_names:
                    try:
                        value = record.get(field)
                        # Handle None values
                        if value is None:
                            record_dict[field] = None
                        else:
                            record_dict[field] = value
                    except Exception as e:
                        record_dict[field] = None
                        if len(errors) < 5:
                            errors.append(f"Field '{field}' in record {i}: {str(e)}")
                
                records.append(record_dict)
                
            except Exception as e:
                if len(errors) < 5:
                    errors.append(f"Record {i}: {str(e)}")
        
        print(f"\n✅ Successfully read {len(records)} records")
        
        if errors:
            print(f"\n⚠️  Encountered {len(errors)} errors (showing first 5):")
            for error in errors[:5]:
                print(f"  - {error}")
        
        # Convert to DataFrame
        print(f"\nCreating DataFrame...")
        df = pd.DataFrame(records)
        
        # Show data types before cleaning
        print(f"\nData types summary:")
        type_counts = df.dtypes.value_counts()
        for dtype, count in type_counts.items():
            print(f"  - {dtype}: {count} columns")
        
        # Clean data based on actual data types
        print(f"\nCleaning data...")
        for col in df.columns:
            # Only apply string operations to object (string) columns
            if df[col].dtype == 'object':
                try:
                    # Check if column actually contains strings
                    sample = df[col].dropna().head(10)
                    if len(sample) > 0 and all(isinstance(x, str) for x in sample):
                        # Safe to apply string operations
                        df[col] = df[col].apply(lambda x: x.strip() if isinstance(x, str) else x)
                except:
                    pass  # Skip if any error
            
            # Convert date columns if needed
            elif 'date' in col.lower() or 'dt' in col.lower():
                try:
                    df[col] = pd.to_datetime(df[col], errors='coerce')
                except:
                    pass
        
        # Display sample data
        print(f"\nSample data (first 5 rows, selected columns):")
        display_cols = ['UNI_NO', 'UIDITEM', 'DESC', 'SIZE', 'TYPE', 'ONHAND', 'CPERUNIT']
        available_cols = [col for col in display_cols if col in df.columns]
        if available_cols:
            print(df[available_cols].head())
        else:
            print(df.iloc[:5, :5])  # First 5 rows, first 5 columns
        
        # Check for data quality issues
        print(f"\nData quality check:")
        print(f"  - Total rows: {len(df)}")
        print(f"  - Total columns: {len(df.columns)}")
        print(f"  - Columns with all nulls: {sum(df[col].isna().all() for col in df.columns)}")
        print(f"  - Rows with all nulls: {sum(df.isna().all(axis=1))}")
        
        # Save to Excel with error handling
        print(f"\nSaving to Excel: {excel_file}")
        try:
            # For large files, use xlsxwriter with constant memory
            if len(df) > 10000:
                with pd.ExcelWriter(excel_file, 
                                  engine='xlsxwriter',
                                  options={'constant_memory': True}) as writer:
                    df.to_excel(writer, index=False, sheet_name='Data')
            else:
                df.to_excel(excel_file, index=False)
            
            print(f"\n✅ Successfully saved to: {excel_file}")
            
        except Exception as e:
            print(f"\n⚠️  Excel save failed: {e}")
            # Try CSV as fallback
            csv_file = excel_file.replace('.xlsx', '.csv')
            df.to_csv(csv_file, index=False)
            print(f"✅ Saved as CSV instead: {csv_file}")
        
        return df
        
    except Exception as e:
        print(f"\n❌ Error: {str(e)}")
        import traceback
        traceback.print_exc()
        return None


def analyze_extracted_csv(csv_file):
    """Analyze the extracted CSV to understand parsing issues"""
    print(f"\nAnalyzing extracted CSV: {csv_file}")
    print("=" * 60)
    
    try:
        # Read first few lines to understand structure
        with open(csv_file, 'r', encoding='utf-8') as f:
            lines = [f.readline().strip() for _ in range(10)]
        
        print("First 10 lines of extracted CSV:")
        for i, line in enumerate(lines):
            print(f"Line {i}: {line[:100]}{'...' if len(line) > 100 else ''}")
        
        # Try reading with pandas
        print("\nTrying to read with pandas...")
        df = pd.read_csv(csv_file, nrows=100)  # Just first 100 rows
        
        print(f"\nDataFrame info:")
        print(f"  - Shape: {df.shape}")
        print(f"  - Columns: {list(df.columns)}")
        print(f"\nFirst 5 rows:")
        print(df.head())
        
        # Check if headers are split
        if len(df.columns) < 10 and len(df) > 0:
            print("\n⚠️  Possible issues detected:")
            print("  - Headers might be split across rows")
            print("  - Data might be improperly delimited")
            
            # Try reading without header
            df_no_header = pd.read_csv(csv_file, header=None, nrows=10)
            print("\nFirst 10 rows without header assumption:")
            print(df_no_header)
        
    except Exception as e:
        print(f"Error analyzing CSV: {e}")


def convert_with_field_mapping(dbf_file, excel_file=None):
    """Alternative approach using field mapping"""
    from dbfread import DBF
    
    print("\nTrying field mapping approach...")
    
    if excel_file is None:
        excel_file = dbf_file.replace('.dat', '_mapped.xlsx')
    
    try:
        # Open DBF and get field information
        table = DBF(dbf_file, encoding='latin-1', ignore_missing_memofile=True)
        
        # Get detailed field information
        print("\nField Information:")
        print("-" * 80)
        print(f"{'Field Name':<15} {'Type':<10} {'Length':<10} {'Decimals':<10}")
        print("-" * 80)
        
        fields_info = []
        for field in table.fields:
            print(f"{field.name:<15} {field.type:<10} {field.length:<10} {field.decimal_count:<10}")
            fields_info.append({
                'name': field.name,
                'type': field.type,
                'length': field.length,
                'decimals': field.decimal_count
            })
        
        # Read data with type conversion
        records = []
        print(f"\nReading and converting data...")
        
        for i, record in enumerate(table):
            if i % 1000 == 0:
                print(f"  Record {i}...")
            
            converted_record = {}
            for field_info in fields_info:
                field_name = field_info['name']
                field_type = field_info['type']
                
                try:
                    value = record.get(field_name)
                    
                    # Type-specific handling
                    if value is None:
                        converted_record[field_name] = None
                    elif field_type == 'C':  # Character
                        converted_record[field_name] = str(value).strip() if value else ''
                    elif field_type in ['N', 'F']:  # Numeric/Float
                        converted_record[field_name] = float(value) if value else 0.0
                    elif field_type == 'L':  # Logical
                        converted_record[field_name] = bool(value) if value else False
                    elif field_type == 'D':  # Date
                        converted_record[field_name] = value  # Keep as is, pandas will handle
                    else:
                        converted_record[field_name] = value
                        
                except Exception:
                    converted_record[field_name] = None
            
            records.append(converted_record)
        
        # Create DataFrame
        df = pd.DataFrame(records)
        
        # Save to Excel
        df.to_excel(excel_file, index=False)
        print(f"\n✅ Saved to: {excel_file}")
        
        return df
        
    except Exception as e:
        print(f"Error: {e}")
        return None


def main():
    print("Fixed DBF Converter with Type Handling")
    print("=" * 60)
    
    if len(sys.argv) > 1:
        input_file = sys.argv[1]
    else:
        input_file = input("Enter .dat file path: ").strip('"')
    
    if not os.path.exists(input_file):
        print(f"❌ File not found: {input_file}")
        input("\nPress Enter to exit...")
        return
    
    # Try main conversion
    df = convert_dbf_properly(input_file)
    
    if df is None:
        print("\n" + "="*60)
        # Try field mapping approach
        df = convert_with_field_mapping(input_file)
    
    # Analyze the extracted CSV if it exists
    extracted_csv = input_file.replace('.dat', '_extracted.csv')
    if os.path.exists(extracted_csv):
        print("\n" + "="*60)
        analyze_extracted_csv(extracted_csv)
    
    print("\n" + "="*60)
    print("Conversion complete!")
    input("\nPress Enter to exit...")


if __name__ == "__main__":
    main()