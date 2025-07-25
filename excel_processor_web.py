import pandas as pd
import numpy as np
from typing import Dict, Any, List

class ExcelProcessorWeb:
    def __init__(self):
        self.summary_lookup = {}
        self.order_quantity_lookup = {}
    
    def process_files(self, main_file_path: str, order_file_path: str) -> Dict[str, Any]:
        """Process both files and return results"""
        try:
            # Process order file
            if not self._process_order_file(order_file_path):
                return {
                    'success': False,
                    'error': 'Failed to process order file - could not find Item and Order Quantity columns'
                }
            
            # Process main file
            workbook = pd.ExcelFile(main_file_path)
            
            # Process summary sheet
            self._process_summary_sheet(main_file_path)
            
            # Process other sheets
            processed_sheets = self._process_other_sheets(main_file_path, workbook)
            
            if not processed_sheets:
                return {
                    'success': False,
                    'error': 'No sheets were processed successfully - check if your main file has the required columns'
                }
            
            # Combine sheets
            combined_df = pd.concat(processed_sheets, ignore_index=True)
            summary_df = pd.read_excel(main_file_path, sheet_name='Summary')
            
            # Calculate statistics
            ordered_qty_count = combined_df['Ordered Qty'].apply(
                lambda x: pd.notna(x) and str(x) != "" and str(x) != "0"
            ).sum()
            total_items = len(combined_df)
            match_rate = (ordered_qty_count / total_items * 100) if total_items > 0 else 0
            
            return {
                'success': True,
                'combined_df': combined_df,
                'summary_df': summary_df,
                'total_items': total_items,
                'matched_items': int(ordered_qty_count),
                'match_rate': match_rate
            }
            
        except Exception as e:
            return {
                'success': False,
                'error': f'Processing error: {str(e)}'
            }
    
    def _process_order_file(self, order_file_path: str) -> bool:
        """Process the order file to create quantity lookup"""
        try:
            print(f"Processing order file: {order_file_path}")
            
            # Read the order file
            order_df = pd.read_excel(order_file_path, header=None)
            
            # Find columns
            item_col = None
            order_qty_col = None
            header_row = None
            
            print("Searching for Item and Order Quantity columns...")
            
            # Search for column headers with flexible matching
            for row_idx in range(min(15, len(order_df))):
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
                return False
            
            print(f"Using Item column at index {item_col}, Order Quantity at index {order_qty_col}")
            
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
                        
                    except (ValueError, TypeError):
                        print(f"Warning: Invalid quantity value '{qty_value}' for item '{item_str}'")
                        continue
            
            self.order_quantity_lookup = item_quantities
            print(f"Successfully processed {processed_rows} rows from order file")
            print(f"Created order quantity lookup with {len(self.order_quantity_lookup)} unique items")
            
            return True
            
        except Exception as e:
            print(f"Error processing order file: {str(e)}")
            return False
    
    def _process_summary_sheet(self, file_path: str):
        """Process Summary sheet and create lookup dictionary"""
        try:
            summary_df = pd.read_excel(file_path, sheet_name='Summary', header=None)
            
            # Find Issue Key and Summary columns
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
            
            # Find Summary column
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
            
            print(f"Created summary lookup dictionary with {len(self.summary_lookup)} entries")
                    
        except Exception as e:
            print(f"Error processing Summary sheet: {str(e)}")
    
    def _process_other_sheets(self, file_path: str, workbook) -> List[pd.DataFrame]:
        """Process all sheets except Summary sheet"""
        processed_sheets = []
        required_columns = ["Planner", "Published", "Item Number", "Item Description", "Oracle On Hand"]
        
        for sheet_name in workbook.sheet_names:
            if sheet_name == 'Summary':
                continue
                
            try:
                print(f"Processing sheet: {sheet_name}")
                
                # Read sheet without header to handle custom positioning
                df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
                
                # Get values from B1 and B2 (0-indexed: B1 = [0,1], B2 = [1,1])
                model_value = ""
                b2c_date_value = ""
                
                if len(df) > 0 and len(df.columns) > 1:
                    if pd.notna(df.iloc[0, 1]):
                        model_value = str(df.iloc[0, 1]).strip()
                
                if len(df) > 1 and len(df.columns) > 1:
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
                            print(f"Found {a1_str} in lookup, setting Model to: {model_value}")
                
                # Find table boundaries
                table_start_row = self._find_table_start(df, required_columns)
                
                if table_start_row is None:
                    print(f"Warning: Required table not found in sheet {sheet_name}")
                    continue
                
                # Process table data
                processed_data = self._extract_table_data(
                    df, table_start_row, required_columns, model_value, b2c_date_value
                )
                
                if processed_data:
                    new_columns = ["Model", "B2C Date"] + required_columns + ["Ordered Qty"]
                    sheet_df = pd.DataFrame(processed_data, columns=new_columns)
                    sheet_df['Source_Sheet'] = sheet_name
                    processed_sheets.append(sheet_df)
                    print(f"Processed {len(processed_data)} rows from {sheet_name}")
                    
            except Exception as e:
                print(f"Error processing sheet {sheet_name}: {str(e)}")
        
        return processed_sheets
    
    def _find_table_start(self, df, required_columns):
        """Find the start of the data table"""
        for row_idx in range(len(df)):
            row_data = df.iloc[row_idx]
            found_columns = 0
            for col in required_columns:
                if any(str(cell).strip() == col for cell in row_data if pd.notna(cell)):
                    found_columns += 1
            
            if found_columns >= 2:  # Found at least 2 required columns
                return row_idx
        
        return None
    
    def _extract_table_data(self, df, table_start_row, required_columns, model_value, b2c_date_value):
        """Extract data from the table"""
        header_row = df.iloc[table_start_row]
        
        # Find column positions for required columns
        column_positions = {}
        for col_idx, cell_value in enumerate(header_row):
            if pd.notna(cell_value):
                cell_str = str(cell_value).strip()
                if cell_str in required_columns:
                    column_positions[cell_str] = col_idx
        
        # Extract data
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
        
        return processed_data
