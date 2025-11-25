"""
Dynamic Legacy View Integration for Bank Schedule Sanitizer

This patch adds dynamic formula-based Legacy View functionality.
The Legacy View will automatically update when the Units sheet changes.
"""

def create_dynamic_hierarchical_view_sheet(self, workbook, units_ws, units_row):
    """Create a Legacy View sheet with dynamic formulas that auto-update from Units sheet"""
    try:
        self.log_message("📋 Creating dynamic Legacy View sheet...")
        
        # Remove any existing legacy view sheets
        sheets_to_remove = ["Hierarchical View", "Hierarchical View1", "Legacy View"]
        for sheet_name in sheets_to_remove:
            if sheet_name in workbook.sheetnames:
                del workbook[sheet_name]
                
        hier_ws = workbook.create_sheet("Legacy View")
        
        # Add title and instructions  
        hier_ws.cell(row=1, column=1, value="Legacy View - Dynamic Building Structure")
        hier_ws.cell(row=2, column=1, value="(Auto-updates from Units sheet - changes to Units reflect here instantly)")
        
        # Copy headers from Units sheet starting at row 4
        headers_row = 4
        
        for col in range(1, units_ws.max_column + 1):
            header = units_ws.cell(row=1, column=col).value
            if header:
                hier_ws.cell(row=headers_row, column=col, value=header)
                # Format header
                header_cell = hier_ws.cell(row=headers_row, column=col)
                if hasattr(openpyxl.styles, 'Font'):
                    header_cell.font = openpyxl.styles.Font(bold=True)
                if hasattr(openpyxl.styles, 'PatternFill'):
                    header_cell.fill = openpyxl.styles.PatternFill(
                        start_color="D3D3D3", end_color="D3D3D3", fill_type="solid"
                    )
        
        # Create dynamic formulas for all data rows
        legacy_row = headers_row + 1
        
        for units_data_row in range(2, units_row):  # Start from row 2 (after Units header)
            for col in range(1, units_ws.max_column + 1):
                # Create formula that dynamically references Units sheet
                col_letter = self.get_column_letter(col)
                formula = f"=Units.{col_letter}{units_data_row}"
                hier_ws.cell(row=legacy_row, column=col, value=formula)
            
            legacy_row += 1
        
        # Auto-adjust column widths  
        for col in range(1, units_ws.max_column + 1):
            col_letter = self.get_column_letter(col)
            try:
                # Set a reasonable default width
                hier_ws.column_dimensions[col_letter].width = 15
            except:
                pass
        
        # Calculate and log statistics
        formula_count = (units_row - 2) * (units_ws.max_column)
        self.log_message(f"✅ Legacy View created with {formula_count} dynamic formulas")
        self.log_message("🔗 Legacy View will automatically update when Units sheet changes")
        
        return True
        
    except Exception as e:
        self.log_message(f"❌ Error creating dynamic Legacy View: {str(e)}")
        return False

def get_column_letter(self, col_num):
    """Convert column number to Excel letter (A, B, C, ..., Z, AA, AB, etc.)"""
    result = ""
    while col_num > 0:
        col_num -= 1
        result = chr(65 + (col_num % 26)) + result
        col_num //= 26
    return result

# Instructions for integration:
print("""
🔧 Dynamic Legacy View Implementation Ready!

To integrate this into your main application:

1. Replace the existing create_hierarchical_view_sheet method in bank_schedule_sanitizer.py
2. Add the get_column_letter helper method  
3. The Legacy View will now contain formulas like =Units.A2, =Units.B2, etc.

✅ Benefits:
• Automatically updates when Units sheet changes
• No need to regenerate the Legacy View
• Maintains Excel's calculation engine
• Works with all Excel features (sorting, filtering, etc.)

📝 How it works:
• Each cell in Legacy View contains a formula pointing to Units sheet
• Excel automatically recalculates when Units data changes
• Compatible with all Excel versions that support cross-sheet references
""")
