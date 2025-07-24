#!/usr/bin/env python3
"""
DAT File Diagnostic Tool
Helps identify the correct format and encoding for .dat files
"""

import os
import sys
from collections import Counter

def diagnose_dat_file(filepath):
    """Comprehensive diagnostic of a DAT file"""
    print("=" * 60)
    print("DAT FILE DIAGNOSTIC REPORT")
    print("=" * 60)
    print(f"File: {filepath}")
    print(f"Size: {os.path.getsize(filepath):,} bytes")
    print()
    
    # Read raw bytes
    with open(filepath, 'rb') as f:
        raw_sample = f.read(1000)  # First 1KB
    
    print("1. RAW BYTES (first 200 bytes):")
    print("-" * 40)
    print(raw_sample[:200])
    print()
    
    print("2. HEX VIEW (first 100 bytes):")
    print("-" * 40)
    hex_view = ' '.join(f'{b:02x}' for b in raw_sample[:100])
    print(hex_view)
    print()
    
    # Try different encodings
    print("3. ENCODING TESTS:")
    print("-" * 40)
    encodings = ['utf-8', 'latin-1', 'cp1252', 'iso-8859-1', 'utf-16', 'ascii']
    
    successful_decodings = []
    for enc in encodings:
        try:
            decoded = raw_sample.decode(enc)
            successful_decodings.append((enc, decoded))
            print(f"✅ {enc}: Success")
            # Show first line
            first_line = decoded.split('\n')[0][:100]
            print(f"   First line: {repr(first_line)}")
        except:
            print(f"❌ {enc}: Failed")
    print()
    
    # Delimiter detection
    print("4. DELIMITER DETECTION:")
    print("-" * 40)
    
    if successful_decodings:
        # Use the first successful encoding
        enc, decoded = successful_decodings[0]
        lines = decoded.split('\n')[:5]  # First 5 lines
        
        delimiters = {
            'Tab (\\t)': '\t',
            'Pipe (|)': '|',
            'Comma (,)': ',',
            'Semicolon (;)': ';',
            'Unit Sep (\\x01)': '\x01',
            'Field Sep (\\x1f)': '\x1f'
        }
        
        delimiter_counts = {}
        for name, delim in delimiters.items():
            count = sum(line.count(delim) for line in lines)
            delimiter_counts[name] = count
            if count > 0:
                print(f"{name}: {count} occurrences in first 5 lines")
        
        # Find most likely delimiter
        if delimiter_counts:
            best_delimiter = max(delimiter_counts, key=delimiter_counts.get)
            print(f"\n🎯 Most likely delimiter: {best_delimiter}")
    print()
    
    # Show decoded lines with different delimiters
    print("5. SAMPLE DATA WITH DIFFERENT DELIMITERS:")
    print("-" * 40)
    
    if successful_decodings:
        enc, decoded = successful_decodings[0]
        first_line = decoded.split('\n')[0]
        
        # Try tab delimiter
        print("With TAB delimiter:")
        parts = first_line.split('\t')
        print(f"  Columns: {len(parts)}")
        for i, part in enumerate(parts[:5]):
            print(f"  Col {i+1}: {repr(part[:30])}")
        
        print("\nWith PIPE delimiter:")
        parts = first_line.split('|')
        print(f"  Columns: {len(parts)}")
        for i, part in enumerate(parts[:5]):
            print(f"  Col {i+1}: {repr(part[:30])}")
        
        print("\nWith COMMA delimiter:")
        parts = first_line.split(',')
        print(f"  Columns: {len(parts)}")
        for i, part in enumerate(parts[:5]):
            print(f"  Col {i+1}: {repr(part[:30])}")
    
    print()
    print("6. FILE STRUCTURE ANALYSIS:")
    print("-" * 40)
    
    # Check for common Sybase patterns
    if b'\x00' in raw_sample:
        print("⚠️  Contains NULL bytes - might be binary format")
    if b'\r\n' in raw_sample:
        print("📝 Windows line endings (\\r\\n)")
    elif b'\n' in raw_sample:
        print("📝 Unix line endings (\\n)")
    elif b'\r' in raw_sample:
        print("📝 Mac line endings (\\r)")
    
    # Check for BOM
    if raw_sample.startswith(b'\xff\xfe'):
        print("📄 UTF-16 LE BOM detected")
    elif raw_sample.startswith(b'\xfe\xff'):
        print("📄 UTF-16 BE BOM detected")
    elif raw_sample.startswith(b'\xef\xbb\xbf'):
        print("📄 UTF-8 BOM detected")
    
    print("\n" + "=" * 60)
    print("RECOMMENDATIONS:")
    print("=" * 60)
    
    # Give recommendations based on findings
    if successful_decodings:
        print(f"1. Use encoding: {successful_decodings[0][0]}")
        
        if delimiter_counts and max(delimiter_counts.values()) > 0:
            best = max(delimiter_counts, key=delimiter_counts.get)
            delim_map = {
                'Tab (\\t)': '\\t',
                'Pipe (|)': '|',
                'Comma (,)': ',',
                'Semicolon (;)': ';'
            }
            if best in delim_map:
                print(f"2. Use delimiter: {delim_map[best]}")
    else:
        print("⚠️  File might be in binary format or use special encoding")
        print("Consider using a hex editor to inspect the file")


def main():
    print("DAT File Diagnostic Tool")
    print("=" * 60)
    
    if len(sys.argv) > 1:
        filepath = sys.argv[1]
    else:
        filepath = input("Enter path to .dat file: ").strip().strip('"')
    
    if not os.path.exists(filepath):
        print(f"❌ File not found: {filepath}")
        input("\nPress Enter to exit...")
        return
    
    try:
        diagnose_dat_file(filepath)
    except Exception as e:
        print(f"\n❌ Error: {str(e)}")
        import traceback
        traceback.print_exc()
    
    input("\nPress Enter to exit...")


if __name__ == "__main__":
    main()