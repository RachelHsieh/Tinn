import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
import os
from openpyxl import load_workbook
import numpy as np

class ExcelProcessor:
    def __init__(self):
        self.root = tk.Tk()
        self.root.withdraw()  # Hide the main window
        self.file_path = None
        self.order_file_path = None
        self.workbook = None
        self.summary_lookup = {}
        self.order_quantity_lookup = {}
        
    def select_file(self):
        """Open file dialog to select Excel file"""
        self.file_path = filedialog.askopenfilename(
            title="Select Main Excel File",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if not self.file_path:
            messagebox.showwarning("Warning", "No main file selected!")
            return False
        return True
    
    def select_order_file(self):
        """Open file dialog to select Order Excel file"""
        self.order_file_path = filedialog.askopenfilename(
            title="Select Order Excel File (with Item and Order Quantity columns)",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if not self.order_file_path:
            messagebox.showwarning("Warning", "No order file selected!")
            return False
        return True
    
    def load_excel_file(self):
        """Load the Excel file and read all sheets"""
        try:
            self.workbook = pd.ExcelFile(self.file_path)
            print(f"Successfully loaded main file: {os.path.basename(self.file_path)}")
            print(f"Available sheets: {self.workbook.sheet_names}")
            return True
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load Excel file: {str(e)}")
            return False
    
    def process_order_file(self):
        """Process the order file to create quantity lookup"""
        try:
            print(f"Processing order file: {os.path.basename(self.order_file_path)}")
            
            # Try to read all sheets to find the data
            try:
                order_workbook = pd.ExcelFile(self.order_file_path)
                print(f"Available sheets in order file: {order_workbook.sheet_names}")
            except:
                pass
            
            # Try to read the first sheet of the order file
            order_df = pd.read_excel(self.order_file_path, header=None)
            
            # Find "Item" and "Order Quantity" columns with more flexible matching
            item_col = None
            order_qty_col = None
            header_row = None
            
            print("Searching for columns...")
            print("First 10 rows of order file:")
            for row_idx in range(min(10, len(order_df))):
                row_values = []
                for col_idx in range(min(10, len(order_df.columns))):
                    if pd.notna(order_df.iloc[row_idx, col_idx]):
                        row_values.append(str(order_df.iloc[row_idx, col_idx]).strip())
                print(f"Row {row_idx}: {row_values}")
            
            # Search for column headers with more flexible matching
            for row_idx in range(min(15, len(order_df))):  # Search in first 15 rows
                for col_idx in range(len(order_df.columns)):
                    if pd.notna(order_df.iloc[row_idx, col_idx]):
                        cell_value = str(order_df.iloc[row_idx, col_idx]).strip().lower()
                        
                        # More flexible item column matching
                        if cell_value in ["item", "item number", "item_number", "itemNumber", "part", "part number", "part_number", "partNumber"]:
                            item_col = col_idx
                            header_row = row_idx
                            print(f"Found item column '{cell_value}' at row {row_idx}, col {col_idx}")
                        
                        # More flexible quantity column matching
                        elif cell_value in ["order quantity", "order qty", "ordered quantity", "quantity", "qty", "order_quantity", "ordered_qty"]:
                            order_qty_col = col_idx
                            header_row = row_idx
                            print(f"Found quantity column '{cell_value}' at row {row_idx}, col {col_idx}")
                
                # If we found both columns, break
                if item_col is not None and order_qty_col is not None:
                    break
            
            if item_col is None or order_qty_col is None:
                print("Warning: Could not find required columns in order file")
                print("Looking for columns containing 'item' or 'quantity':")
                
                # Show all possible column headers for debugging
                for row_idx in range(min(10, len(order_df))):
                    for col_idx in range(len(order_df.columns)):
                        if pd.notna(order_df.iloc[row_idx, col_idx]):
                            cell_value = str(order_df.iloc[row_idx, col_idx]).strip()
                            if 'item' in cell_value.lower() or 'quantity' in cell_value.lower() or 'qty' in cell_value.lower():
                                print(f"  Found potential column: '{cell_value}' at row {row_idx}, col {col_idx}")
                
                return False
            
            print(f"Using Item column at index {item_col} (row {header_row}), Order Quantity at index {order_qty_col}")
            
            # Process the data starting from the row after headers
            item_quantities = {}
            processed_rows = 0
            
            for row_idx in range(header_row + 1, len(order_df)):
                item_value = order_df.iloc[row_idx, item_col]
                qty_value = order_df.iloc[row_idx, order_qty_col]
                
                if pd.notna(item_value) and pd.notna(qty_value):
                    item_str = str(item_value).strip()
                    try:
                        qty_num = float(qty_value)
                        if item_str in item_quantities:
                            item_quantities[item_str] += qty_num
                        else:
                            item_quantities[item_str] = qty_num
                        processed_rows += 1
                        
                        # Show first few items being processed for debugging
                        if processed_rows <= 5:
                            print(f"  Processing: Item '{item_str}' -> Qty {qty_num}")
                            
                    except (ValueError, TypeError):
                        print(f"Warning: Invalid quantity value '{qty_value}' for item '{item_str}'")
                        continue
            
            self.order_quantity_lookup = item_quantities
            print(f"Successfully processed {processed_rows} rows")
            print(f"Created order quantity lookup with {len(self.order_quantity_lookup)} unique items")
            
            # Print some examples
            if len(self.order_quantity_lookup) > 0:
                print("Sample order quantities (first 10):")
                for i, (item, qty) in enumerate(list(self.order_quantity_lookup.items())[:10]):
                    print(f"  '{item}': {qty}")
                if len(self.order_quantity_lookup) > 10:
                    print("  ...")
            else:
                print("WARNING: No items were processed from the order file!")
            
            return True
            
        except Exception as e:
            print(f"Error processing order file: {str(e)}")
            import traceback
            traceback.print_exc()
            return False
    
    def process_summary_sheet(self):
        """Process Summary sheet and create lookup dictionary"""
        try:
            # Read Summary sheet
            summary_df = pd.read_excel(self.file_path, sheet_name='Summary', header=None)
            
            # Find the table with "Issue Key" and "Summary" columns
            issue_key_row = None
            issue_key_col = None
            
            # Search for "Issue Key" in the sheet
            for row_idx in range(len(summary_df)):
                for col_idx in range(len(summary_df.columns)):
                    if pd.notna(summary_df.iloc[row_idx, col_idx]) and \
                       str(summary_df.iloc[row_idx, col_idx]).strip() == "Issue key":
                        issue_key_row = row_idx
                        issue_key_col = col_idx
                        break
                if issue_key_row is not None:
                    break
            
            if issue_key_row is None:
                print("Warning: 'Issue key' column not found in Summary sheet")
                return
            
            # Find Summary column (should be next to Issue Key)
            summary_col = None
            header_row = summary_df.iloc[issue_key_row]
            for col_idx in range(len(header_row)):
                if pd.notna(header_row.iloc[col_idx]) and \
                   str(header_row.iloc[col_idx]).strip() == "Summary":
                    summary_col = col_idx
                    break
            
            if summary_col is None:
                print("Warning: 'Summary' column not found in Summary sheet")
                return
            
            # Create lookup dictionary
            for row_idx in range(issue_key_row + 1, len(summary_df)):
                issue_key = summary_df.iloc[row_idx, issue_key_col]
                summary_value = summary_df.iloc[row_idx, summary_col]
                
                if pd.notna(issue_key) and pd.notna(summary_value):
                    self.summary_lookup[str(issue_key).strip()] = str(summary_value).strip()
            
            print(f"Created lookup dictionary with {len(self.summary_lookup)} entries")
            
        except Exception as e:
            print(f"Error processing Summary sheet: {str(e)}")
    
    def find_table_boundaries(self, df, required_columns):
        """Find the boundaries of a table containing required columns"""
        table_start_row = None
        table_start_col = None
        
        # Search for required columns
        for row_idx in range(len(df)):
            for col_idx in range(len(df.columns)):
                if pd.notna(df.iloc[row_idx, col_idx]):
                    cell_value = str(df.iloc[row_idx, col_idx]).strip()
                    if cell_value in required_columns:
                        # Check if this row contains multiple required columns
                        row_data = df.iloc[row_idx]
                        found_columns = 0
                        for col in required_columns:
                            if any(str(cell).strip() == col for cell in row_data if pd.notna(cell)):
                                found_columns += 1
                        
                        if found_columns >= 2:  # Found at least 2 required columns
                            table_start_row = row_idx
                            table_start_col = col_idx
                            break
            
            if table_start_row is not None:
                break
        
        return table_start_row, table_start_col
    
    def process_other_sheets(self):
        """Process all sheets except Summary sheet"""
        processed_sheets = []
        required_columns = ["Planner", "Published", "Item Number", "Item Description", "Oracle On Hand"]
        
        for sheet_name in self.workbook.sheet_names:
            if sheet_name == 'Summary':
                continue
                
            try:
                print(f"Processing sheet: {sheet_name}")
                
                # Read sheet without header to handle custom positioning
                df = pd.read_excel(self.file_path, sheet_name=sheet_name, header=None)
                
                # Get values from B1 and B3 (0-indexed: B1 = [0,1], B2 = [1,1])
                model_value = ""
                b2c_date_value = ""
                
                if len(df) > 0 and len(df.columns) > 1:
                    if pd.notna(df.iloc[0, 1]):
                        model_value = str(df.iloc[0, 1]).strip()
                
                if len(df) > 2 and len(df.columns) > 1:
                    if pd.notna(df.iloc[1, 1]):
                        b2c_date_value = str(df.iloc[1, 1]).strip()
                
                # Check if A1 contains an Issue Key and do vlookup
                if len(df) > 0 and len(df.columns) > 0:
                    a1_value = df.iloc[0, 0]
                    if pd.notna(a1_value):
                        a1_str = str(a1_value).strip()
                        if a1_str in self.summary_lookup:
                            # Fill B1 with the Summary value
                            model_value = self.summary_lookup[a1_str]
                            print(f"Found {a1_str} in lookup, setting B1 to: {model_value}")
                
                # Find the table containing required columns
                table_start_row, table_start_col = self.find_table_boundaries(df, required_columns)
                
                if table_start_row is None:
                    print(f"Warning: Required table not found in sheet {sheet_name}")
                    continue
                
                # Extract the table data
                # First, get the header row
                header_row = df.iloc[table_start_row]
                
                # Find column positions for required columns
                column_positions = {}
                for col_idx, cell_value in enumerate(header_row):
                    if pd.notna(cell_value):
                        cell_str = str(cell_value).strip()
                        if cell_str in required_columns:
                            column_positions[cell_str] = col_idx
                
                # Create new DataFrame with required columns plus Model, B2C Date, and Ordered Qty
                new_columns = ["Model", "B2C Date"] + required_columns + ["Ordered Qty"]
                processed_data = []
                
                # Process data rows (starting from the row after header)
                for row_idx in range(table_start_row + 1, len(df)):
                    row_data = {}
                    
                    # Add Model and B2C Date
                    row_data["Model"] = model_value
                    row_data["B2C Date"] = b2c_date_value
                    
                    # Add required columns data
                    has_data = False
                    item_number = None
                    
                    for col_name in required_columns:
                        if col_name in column_positions:
                            col_idx = column_positions[col_name]
                            if col_idx < len(df.columns):
                                cell_value = df.iloc[row_idx, col_idx]
                                row_data[col_name] = cell_value
                                if pd.notna(cell_value):
                                    has_data = True
                                    # Store item number for quantity lookup
                                    if col_name == "Item Number":
                                        item_number = str(cell_value).strip()
                            else:
                                row_data[col_name] = ""
                        else:
                            row_data[col_name] = ""
                    
                    # Add Ordered Qty using vlookup
                    ordered_qty = ""
                    if item_number:
                        # Try exact match first
                        if item_number in self.order_quantity_lookup:
                            ordered_qty = self.order_quantity_lookup[item_number]
                        else:
                            # Try to find partial matches or different formats
                            item_number_clean = item_number.strip().upper()
                            for lookup_item, lookup_qty in self.order_quantity_lookup.items():
                                lookup_item_clean = lookup_item.strip().upper()
                                if item_number_clean == lookup_item_clean:
                                    ordered_qty = lookup_qty
                                    break
                    
                    row_data["Ordered Qty"] = ordered_qty
                    
                    # Only add row if it has some data
                    if has_data:
                        processed_data.append(row_data)
                
                if processed_data:
                    sheet_df = pd.DataFrame(processed_data, columns=new_columns)
                    sheet_df['Source_Sheet'] = sheet_name  # Add source sheet identifier
                    processed_sheets.append(sheet_df)
                    print(f"Processed {len(processed_data)} rows from {sheet_name}")
                
            except Exception as e:
                print(f"Error processing sheet {sheet_name}: {str(e)}")
        
        return processed_sheets
    
    def merge_sheets_and_save(self, processed_sheets):
        """Merge all processed sheets and save to new file"""
        if not processed_sheets:
            messagebox.showwarning("Warning", "No sheets were processed successfully!")
            return
        
        try:
            # Combine all processed sheets
            combined_df = pd.concat(processed_sheets, ignore_index=True)
            
            # Read original Summary sheet
            summary_df = pd.read_excel(self.file_path, sheet_name='Summary')
            
            # Create output file path
            base_name = os.path.splitext(self.file_path)[0]
            output_path = f"{base_name}_processed.xlsx"
            
            # Save to new Excel file
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                # Save Summary sheet (unchanged)
                summary_df.to_excel(writer, sheet_name='Summary', index=False)
                
                # Save combined sheet
                combined_df.to_excel(writer, sheet_name='Combined', index=False)
            
            print(f"File saved successfully: {output_path}")
            print(f"Combined sheet contains {len(combined_df)} total rows")
            
            # Show summary of ordered quantities found
            ordered_qty_count = combined_df['Ordered Qty'].apply(lambda x: pd.notna(x) and str(x) != "" and str(x) != "0").sum()
            total_items = len(combined_df)
            print(f"Found ordered quantities for {ordered_qty_count} out of {total_items} items")
            
            # Show some examples of matches/non-matches for debugging
            print("\nSample matching results:")
            sample_items = combined_df[['Item Number', 'Ordered Qty']].head(10)
            for idx, row in sample_items.iterrows():
                item_num = row['Item Number']
                ordered_qty = row['Ordered Qty']
                status = "MATCHED" if pd.notna(ordered_qty) and str(ordered_qty) != "" else "NO MATCH"
                print(f"  '{item_num}' -> {ordered_qty} ({status})")
            
            messagebox.showinfo("Success", f"Processing completed!\nOutput file: {os.path.basename(output_path)}\nOrdered quantities found for {ordered_qty_count} out of {total_items} items")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save file: {str(e)}")
    
    def run(self):
        """Main execution flow"""
        print("Excel Data Processing Tool with Order Quantities")
        print("=" * 50)
        
        # Step 1: Select main file
        if not self.select_file():
            return
        
        # Step 2: Select order file
        if not self.select_order_file():
            return
        
        # Step 3: Load Excel file
        if not self.load_excel_file():
            return
        
        # Step 4: Process order file
        print("\nStep 1: Processing order file...")
        if not self.process_order_file():
            print("Failed to process order file. Continuing without order quantities...")
        
        # Step 5: Process Summary sheet for lookup
        print("\nStep 2: Processing Summary sheet...")
        self.process_summary_sheet()
        
        # Step 6: Process other sheets
        print("\nStep 3: Processing other sheets...")
        processed_sheets = self.process_other_sheets()
        
        # Step 7: Merge and save
        print("\nStep 4: Merging sheets and saving...")
        self.merge_sheets_and_save(processed_sheets)
        
        print("\nProcessing completed!")

def main():
    """Main function to run the Excel processor"""
    processor = ExcelProcessor()
    processor.run()

if __name__ == "__main__":
    main()