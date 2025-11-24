#!/usr/bin/env python3
"""
Test script to verify Excel file processing capabilities
"""

import pandas as pd
import os
from pathlib import Path

def test_excel_processing():
    """Test basic Excel file reading and writing."""
    
    print("Testing Excel file processing capabilities...")
    
    # Test file path
    test_file = "/Volumes/Marketing/Simone Morciano/python working folder/bankScheduleSanitizer/data/Leasing Bank Schedule June 2025.xlsx"
    
    try:
        # Test reading the Excel file
        print(f"\n1. Testing file reading: {os.path.basename(test_file)}")
        
        if not os.path.exists(test_file):
            print(f"❌ Test file not found: {test_file}")
            return False
            
        # Read Excel file
        excel_file = pd.ExcelFile(test_file)
        sheet_names = excel_file.sheet_names
        
        print(f"✅ File read successfully")
        print(f"   Sheet names found: {sheet_names}")
        
        # Check for Bank Schedule sheet
        if "Bank Schedule" in sheet_names:
            print("✅ 'Bank Schedule' sheet found")
            
            # Read the specific sheet
            df = pd.read_excel(test_file, sheet_name="Bank Schedule")
            print(f"   Sheet dimensions: {df.shape[0]} rows x {df.shape[1]} columns")
            print(f"   Columns: {list(df.columns)[:5]}{'...' if len(df.columns) > 5 else ''}")
        else:
            print("⚠️  'Bank Schedule' sheet not found")
            print(f"   Available sheets: {sheet_names}")
        
        print(f"\n2. Testing file duplication...")
        
        # Test creating a copy (simulate sanitization)
        output_path = "/tmp/test_sanitized_output.xlsx"
        
        # Simple copy using pandas
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            for sheet_name in sheet_names:
                df = pd.read_excel(test_file, sheet_name=sheet_name)
                df.to_excel(writer, sheet_name=sheet_name, index=False)
        
        print(f"✅ File duplicated successfully to: {output_path}")
        
        # Verify the output file
        output_excel = pd.ExcelFile(output_path)
        output_sheets = output_excel.sheet_names
        print(f"   Output file sheets: {output_sheets}")
        
        # Clean up
        if os.path.exists(output_path):
            os.remove(output_path)
            print("✅ Test file cleaned up")
        
        print("\n🎉 All tests passed! Excel processing is working correctly.")
        return True
        
    except Exception as e:
        print(f"❌ Error during testing: {str(e)}")
        return False

if __name__ == "__main__":
    test_excel_processing()
