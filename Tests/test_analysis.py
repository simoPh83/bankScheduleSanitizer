#!/usr/bin/env python3
"""
Test script to verify the Bank Schedule data analysis functionality
"""

import pandas as pd
import sys
import os

# Add the current directory to path so we can import our main module
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from bank_schedule_sanitizer import BankScheduleSanitizer
import tkinter as tk

def test_bank_schedule_analysis():
    """Test the Bank Schedule data analysis."""
    
    print("Testing Bank Schedule Analysis...")
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
        
        print("📊 Running data analysis...")
        
        # Test the analysis function
        results = sanitizer.analyze_bank_schedule_data(test_file)
        
        print("\n🎉 ANALYSIS RESULTS:")
        print("-" * 30)
        print(f"🏢 Buildings: {results['buildings']}")
        print(f"🏠 Units: {results['units']}")
        print(f"📝 Empty rows: {results['empty_rows']}")
        print(f"📋 Total rows: {results['total_rows']}")
        print("-" * 30)
        
        # Validate results make sense
        total_analyzed = results['buildings'] + results['units'] + results['empty_rows']
        print(f"\n🔍 Validation:")
        print(f"Total analyzed: {total_analyzed}")
        print(f"Total rows: {results['total_rows']}")
        
        if total_analyzed == results['total_rows']:
            print("✅ Row count validation PASSED")
        else:
            print("⚠️  Row count validation: Some rows may not be categorized")
        
        # Check if we found any data
        if results['buildings'] > 0 and results['units'] > 0:
            print("✅ Found both buildings and units - looks good!")
        elif results['buildings'] == 0:
            print("⚠️  No buildings found - check analysis logic")
        elif results['units'] == 0:
            print("⚠️  No units found - check analysis logic")
        
        root.destroy()
        return True
        
    except Exception as e:
        print(f"❌ Error during analysis test: {str(e)}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    success = test_bank_schedule_analysis()
    sys.exit(0 if success else 1)
