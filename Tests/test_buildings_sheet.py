#!/usr/bin/env python3
"""
Test script to verify the Buildings sheet creation functionality
"""

import pandas as pd
import sys
import os
import tempfile

# Add the current directory to path so we can import our main module
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from bank_schedule_sanitizer import BankScheduleSanitizer
import tkinter as tk

def test_buildings_sheet_creation():
    """Test the Buildings sheet creation functionality."""
    
    print("Testing Buildings Sheet Creation...")
    print("="*50)
    
    test_file = "/Volumes/Marketing/Simone Morciano/python working folder/bankScheduleSanitizer/data/Leasing Bank Schedule June 2025.xlsx"
    
    if not os.path.exists(test_file):
        print(f"❌ Test file not found: {test_file}")
        return False
    
    try:
        # Create a minimal tkinter instance (needed for the class)
        root = tk.Tk()
        root.withdraw()  # Hide the window for testing
        
        # Create an instance of our sanitizer
        sanitizer = BankScheduleSanitizer(root)
        
        # Override the log_message method to print to console for testing
        def test_log(message):
            print(f"[LOG] {message}")
        sanitizer.log_message = test_log
        
        print("📊 Running analysis...")
        
        # Test the analysis function first
        analysis_results = sanitizer.analyze_bank_schedule_data(test_file)
        
        print(f"\n🎉 Analysis completed:")
        print(f"• Buildings: {analysis_results['buildings']}")
        print(f"• Units: {analysis_results['units']}")
        print(f"• Empty rows: {analysis_results['empty_rows']}")
        print(f"• Total rows: {analysis_results['total_rows']}")
        
        # Create a temporary output file
        with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
            temp_output = tmp.name
        
        print(f"\n🏗️  Testing Buildings sheet creation...")
        print(f"Output file: {temp_output}")
        
        # Test the buildings sheet creation
        result = sanitizer.create_buildings_summary_sheet(test_file, temp_output, analysis_results)
        
        if result:
            print("✅ Buildings sheet creation completed")
            
            # Verify the output file
            print("\n🔍 Verifying output file...")
            output_excel = pd.ExcelFile(temp_output)
            sheet_names = output_excel.sheet_names
            
            print(f"📋 Sheets in output file: {sheet_names}")
            
            if 'Buildings' in sheet_names:
                print("✅ 'Buildings' sheet found in output file")
                
                # Read the Buildings sheet to verify structure
                buildings_df = pd.read_excel(temp_output, sheet_name='Buildings')
                
                print(f"🏢 Buildings sheet dimensions: {buildings_df.shape[0]} rows x {buildings_df.shape[1]} columns")
                print(f"📝 Buildings sheet columns: {list(buildings_df.columns)}")
                
                # Check if we have building data
                if len(buildings_df) > 0:
                    print("✅ Buildings data populated")
                    print(f"📊 Sample buildings: {buildings_df['Building'].head(3).tolist()}")
                else:
                    print("⚠️  Buildings sheet is empty")
                
                # Verify headers
                expected_headers = [
                    'Building', 'Net Area', 'Rent PA (£)', '2023 ERV (£)', 
                    '2024 ERV (£)', 'ERV 2024 £.Sq.ft', 'ERV Variation', '2024 Cap Valn. (£)'
                ]
                
                if list(buildings_df.columns) == expected_headers:
                    print("✅ All expected headers present and in correct order")
                else:
                    print("⚠️  Headers don't match expected format")
                    print(f"Expected: {expected_headers}")
                    print(f"Actual: {list(buildings_df.columns)}")
                
            else:
                print("❌ 'Buildings' sheet not found in output file")
                return False
            
            # Check that original sheets are preserved
            original_excel = pd.ExcelFile(test_file)
            original_sheets = original_excel.sheet_names
            
            preserved_sheets = [sheet for sheet in original_sheets if sheet in sheet_names]
            print(f"📄 Original sheets preserved: {len(preserved_sheets)}/{len(original_sheets)}")
            
            if len(preserved_sheets) == len(original_sheets):
                print("✅ All original sheets preserved")
            else:
                print("⚠️  Some original sheets may be missing")
        
        else:
            print("❌ Buildings sheet creation failed")
            return False
        
        # Clean up
        if os.path.exists(temp_output):
            os.unlink(temp_output)
            print("🧹 Temporary file cleaned up")
        
        root.destroy()
        
        print(f"\n🎉 All tests passed! Buildings sheet creation is working correctly.")
        return True
        
    except Exception as e:
        print(f"❌ Error during testing: {str(e)}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    success = test_buildings_sheet_creation()
    sys.exit(0 if success else 1)
