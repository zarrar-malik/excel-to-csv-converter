#!/usr/bin/env python3
"""
Excel to CSV Converter with Email Detection
==========================================

A command-line tool to convert Excel (.xlsx) files to CSV format while automatically
detecting and extracting columns containing email addresses.
"""

import os
import sys
import argparse
import pandas as pd
from typing import Optional, List

VERSION = "1.0.0"
SUPPORTED_FORMATS = ['.xlsx', '.xls']

def find_email_column(df: pd.DataFrame) -> Optional[str]:
    """
    Identifies columns containing email addresses in a DataFrame.
    
    Args:
        df: Input pandas DataFrame
        
    Returns:
        Name of the first column containing email addresses, or None if none found.
    """
    email_pattern = r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b'
    
    for column in df.columns:
        if pd.api.types.is_string_dtype(df[column]):
            if df[column].str.contains(email_pattern, na=False, regex=True).any():
                return column
    return None

def convert_file(
    input_path: str,
    output_path: str,
    email_only: bool = True,
    verbose: bool = False
) -> bool:
    """
    Converts an Excel file to CSV format.
    
    Args:
        input_path: Path to input Excel file
        output_path: Path to output CSV file
        email_only: Whether to extract only email columns
        verbose: Whether to print detailed progress
        
    Returns:
        True if conversion succeeded, False otherwise
    """
    try:
        if verbose:
            print(f"Reading {os.path.basename(input_path)}...")
            
        df = pd.read_excel(input_path)
        
        if email_only:
            email_col = find_email_column(df)
            if email_col:
                if verbose:
                    print(f"Found email column: {email_col}")
                df = df[[email_col]]
            elif verbose:
                print("No email column found - exporting all columns")
        
        df.to_csv(output_path, index=False, encoding='utf-8')
        return True
        
    except Exception as e:
        print(f"Error converting {input_path}: {str(e)}", file=sys.stderr)
        return False

def process_directory(
    input_dir: str,
    output_dir: str,
    email_only: bool = True,
    verbose: bool = False
) -> int:
    """
    Processes all Excel files in a directory.
    
    Args:
        input_dir: Input directory path
        output_dir: Output directory path
        email_only: Whether to extract only email columns
        verbose: Whether to print detailed progress
        
    Returns:
        Number of files successfully converted
    """
    if not os.path.exists(input_dir):
        print(f"Error: Input directory not found: {input_dir}", file=sys.stderr)
        return 0
        
    os.makedirs(output_dir, exist_ok=True)
    success_count = 0
    
    for filename in os.listdir(input_dir):
        if any(filename.lower().endswith(ext) for ext in SUPPORTED_FORMATS):
            input_path = os.path.join(input_dir, filename)
            output_filename = os.path.splitext(filename)[0] + '.csv'
            output_path = os.path.join(output_dir, output_filename)
            
            if convert_file(input_path, output_path, email_only, verbose):
                success_count += 1
                if verbose:
                    print(f"Successfully converted to {output_filename}")
    
    return success_count

def main():
    parser = argparse.ArgumentParser(
        description="Convert Excel files to CSV with email detection",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter
    )
    parser.add_argument(
        '-i', '--input',
        required=True,
        help="Input file or directory path"
    )
    parser.add_argument(
        '-o', '--output',
        required=True,
        help="Output file or directory path"
    )
    parser.add_argument(
        '--all-columns',
        action='store_false',
        dest='email_only',
        help="Export all columns instead of just email columns"
    )
    parser.add_argument(
        '-v', '--verbose',
        action='store_true',
        help="Show detailed conversion progress"
    )
    parser.add_argument(
        '--version',
        action='version',
        version=f'%(prog)s {VERSION}'
    )
    
    args = parser.parse_args()
    
    # Determine if input is file or directory
    if os.path.isfile(args.input):
        # Single file conversion
        if not any(args.input.lower().endswith(ext) for ext in SUPPORTED_FORMATS):
            print(f"Error: Unsupported file format. Supported formats: {', '.join(SUPPORTED_FORMATS)}", file=sys.stderr)
            sys.exit(1)
            
        if convert_file(args.input, args.output, args.email_only, args.verbose):
            print(f"Successfully converted to {args.output}")
        else:
            sys.exit(1)
    else:
        # Directory processing
        success_count = process_directory(
            args.input,
            args.output,
            args.email_only,
            args.verbose
        )
        
        if args.verbose or success_count == 0:
            print(f"\nConverted {success_count} file(s)")
            
        if success_count == 0:
            sys.exit(1)

if __name__ == "__main__":
    main()