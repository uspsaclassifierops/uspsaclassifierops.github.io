#!/usr/bin/env python3
"""
USPSA Classifier Library - Excel to JavaScript Converter

Usage:
    python classifier_converter.py input.xlsx output.js

This script reads an Excel file containing USPSA classifier data and converts it
to a JavaScript file with a standardized JSON array for use in the Classifier Library
web application.

Features:
- Standardizes property names to camelCase JavaScript conventions
- Validates data integrity and provides warnings for potential issues
- Handles all classifier attributes consistently
- Provides statistics about the conversion process
"""

import sys
import json
import pandas as pd
import re
from datetime import datetime

def standardize_property_name(name):
    """Convert column names to standardized camelCase property names."""
    # Special case handling for known column names
    special_cases = {
        "Indoor & No Steel": "indoorNoSteel",
        "10 Rounds or Less": "tenRoundsOrLess",
        "Has SHO / WHO": "hasShoWho",
        "Up Range Start": "upRangeStart",
        "Seated Start": "seatedStart",
        "Has Barricade": "hasBarricade",
        "Has Steel": "hasSteel",
        "String Count": "stringCount",
        "Scoring Type": "scoringType",
        "Wall Count": "wallCount",
        "Back Berm Only": "backBermOnly",
        "Ban State": "banState",
        "Mandatory Reload": "mandatoryReload",
        "Stand and Deliver": "standAndDeliver",
        "Stage Style": "stageStyle",
        "Round Count": "roundCount",
        "Stage Identifier": "stageIdentifier",
        "Stage Name": "stageName"
    }
    
    # Check if it's a special case
    if name in special_cases:
        return special_cases[name]
    
    # Regular camelCase conversion for other names
    # Replace spaces and other non-alphanumeric characters with underscores
    s = re.sub(r'[^a-zA-Z0-9]', '_', name)
    # Split by underscore
    words = s.split('_')
    # First word lowercase, rest capitalized
    return words[0].lower() + ''.join(word.capitalize() for word in words[1:] if word)

def validate_data(df):
    """Perform data validation and cleanup on the DataFrame."""
    validation_issues = []
    
    # Check for missing values in important columns
    important_columns = ['Stage Name', 'Stage Identifier', 'Round Count', 'Scoring Type']
    for col in important_columns:
        if col in df.columns and df[col].isna().any():
            missing_count = df[col].isna().sum()
            validation_issues.append(f"Warning: {missing_count} missing values found in '{col}' column")
    
    # Convert YES/NO columns to consistent format
    boolean_columns = [
        'Indoor', 'Indoor & No Steel', 'Back Berm Only', '10 Rounds or Less',
        'Ban State', 'Mandatory Reload', 'Stand and Deliver', 'Box to Box',
        'Stage Style', 'Has SHO / WHO', 'Up Range Start', 'Seated Start',
        'Has Barricade', 'Has Steel'
    ]
    
    for col in boolean_columns:
        if col in df.columns:
            # Count NA values
            na_count = df[col].isna().sum()
            if na_count > 0:
                validation_issues.append(f"Warning: {na_count} missing values in '{col}' set to 'NO'")
            
            # Fill NA values with "NO"
            df[col] = df[col].fillna("NO")
            
            # Convert to uppercase
            df[col] = df[col].str.upper()
            
            # Ensure only YES/NO values
            non_standard = df[col].apply(lambda x: x not in ["YES", "NO"]).sum()
            if non_standard > 0:
                validation_issues.append(f"Warning: {non_standard} non-standard values in '{col}' normalized")
            
            df[col] = df[col].apply(lambda x: "YES" if (x == "YES" or x == "Y" or x is True) else "NO")
    
    # Convert numeric columns to appropriate types
    numeric_columns = ['Round Count', 'String Count', 'Wall Count', 'Width', 'Depth']
    for col in numeric_columns:
        if col in df.columns:
            # Count NA values
            na_count = df[col].isna().sum()
            if na_count > 0:
                validation_issues.append(f"Warning: {na_count} missing values in '{col}' set to 0")
            
            # Convert to numeric
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0).astype(int)
    
    return df, validation_issues

def convert_excel_to_js(input_file, output_file):
    """Convert Excel file to JavaScript with standardized property names."""
    try:
        print(f"\n=== USPSA Classifier Library Converter ===")
        print(f"Reading Excel file: {input_file}")
        
        # Read Excel file
        df = pd.read_excel(input_file, sheet_name="Classifiers")
        
        print(f"Found {len(df)} classifier records")
        
        # Validate and clean data
        df, validation_issues = validate_data(df)
        
        # Report validation issues
        if validation_issues:
            print("\nData Validation Issues:")
            for issue in validation_issues:
                print(f"  • {issue}")
        
        # Rename columns with standardized property names
        column_mapping = {col: standardize_property_name(col) for col in df.columns}
        mapped_columns = [f"{col} → {column_mapping[col]}" for col in df.columns]
        
        print("\nColumn mapping:")
        for mapping in sorted(mapped_columns):
            print(f"  • {mapping}")
        
        # Apply the column mapping
        df = df.rename(columns=column_mapping)
        
        # Check for required columns and add them if missing
        required_properties = [
            'stageName', 'stageIdentifier', 'indoor', 'indoorNoSteel', 'backBermOnly', 
            'tenRoundsOrLess', 'banState', 'roundCount', 'mandatoryReload', 'standAndDeliver', 
            'boxToBox', 'stageStyle', 'hasShoWho', 'upRangeStart', 'seatedStart', 'wallCount', 
            'hasBarricade', 'hasSteel', 'stringCount', 'scoringType', 'width', 'depth'
        ]
        
        # Add any missing required columns with default values
        missing_columns = []
        for prop in required_properties:
            if prop not in df.columns:
                missing_columns.append(prop)
                if prop in ['roundCount', 'stringCount', 'wallCount', 'width', 'depth']:
                    df[prop] = 0
                elif prop in ['scoringType']:
                    df[prop] = "COMSTOCK"
                elif prop in ['stageName', 'stageIdentifier']:
                    df[prop] = "Unknown"
                else:
                    df[prop] = "NO"
        
        if missing_columns:
            print(f"\nAdded missing columns with default values:")
            for col in missing_columns:
                print(f"  • {col}")
        
        # Convert DataFrame to list of dictionaries
        data = df.to_dict(orient='records')
        
        # Create JavaScript file content
        js_content = "// USPSA Classifier Library Data\n"
        js_content += "// Generated by classifier_converter.py\n"
        js_content += f"// Source: {input_file}\n"
        js_content += f"// Date: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n"
        js_content += "const classifierData = "
        js_content += json.dumps(data, indent=2)
        js_content += ";\n"
        
        # Write to JavaScript file
        with open(output_file, 'w') as f:
            f.write(js_content)
        
        print(f"\nSuccessfully converted to {output_file}")
        
        # Provide a summary of the data
        print("\nData Summary:")
        print(f"  • Total classifiers: {len(data)}")
        print(f"  • Indoor classifiers: {sum(1 for item in data if item['indoor'] == 'YES')}")
        print(f"  • Classifiers with steel: {sum(1 for item in data if item['hasSteel'] == 'YES')}")
        print(f"  • Classifiers with barricade: {sum(1 for item in data if item['hasBarricade'] == 'YES')}")
        print(f"  • Comstock scoring: {sum(1 for item in data if item['scoringType'] == 'COMSTOCK')}")
        print(f"  • Virginia scoring: {sum(1 for item in data if item['scoringType'] == 'VIRGINIA')}")
        print(f"  • Average round count: {sum(item['roundCount'] for item in data) / len(data):.1f}")
        
    except FileNotFoundError:
        print(f"Error: File not found: {input_file}")
        sys.exit(1)
    except Exception as e:
        print(f"Error: {str(e)}")
        sys.exit(1)

if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Usage: python classifier_converter.py input.xlsx output.js")
        sys.exit(1)
    
    input_file = sys.argv[1]
    output_file = sys.argv[2]
    
    convert_excel_to_js(input_file, output_file)
    print("\nConversion complete!")