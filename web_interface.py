import streamlit as st
import pandas as pd
import os
import tempfile
from io import BytesIO
import sys

# Import your Excel processor
from excel_processor_web import ExcelProcessorWeb

def main():
    st.set_page_config(
        page_title="Excel Data Processor",
        page_icon="📊",
        layout="wide"
    )
    
    st.title("📊 Excel Data Processor with Order Quantities")
    st.markdown("---")
    
    # Initialize session state
    if 'processor' not in st.session_state:
        st.session_state.processor = ExcelProcessorWeb()
    
    # Instructions at the top
    with st.expander("📖 How to Use This Tool", expanded=False):
        st.markdown("""
        **Step-by-step instructions:**
        
        1. **Upload Main Excel File**: Your primary Excel file containing:
           - A 'Summary' sheet with 'Issue key' and 'Summary' columns
           - Data sheets with: Planner, Published, Item Number, Item Description, Oracle On Hand
        
        2. **Upload Order Excel File**: Excel file containing:
           - A column named 'Item' (or similar: 'Item Number', 'Part Number')
           - A column named 'Order Quantity' (or similar: 'Qty', 'Quantity')
        
        3. **Click Process**: The tool will process both files and combine the data
        
        4. **Download Results**: Get your processed Excel file with order quantities added
        
        **The tool will:**
        - ✅ Combine data from all sheets (except Summary)
        - ✅ Add Model and B2C Date information from each sheet
        - ✅ Lookup and add Order Quantities for each Item Number
        - ✅ Create a clean, consolidated output file
        """)
    
    # File upload section
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("1️⃣ Upload Main Excel File")
        main_file = st.file_uploader(
            "Choose your main Excel file",
            type=['xlsx', 'xls'],
            key="main_file",
            help="This should be your primary Excel file with Summary sheet and data sheets"
        )
        
        if main_file is not None:
            st.success(f"✅ Main file uploaded: {main_file.name}")
            st.info(f"File size: {main_file.size:,} bytes")
    
    with col2:
        st.subheader("2️⃣ Upload Order Excel File")
        order_file = st.file_uploader(
            "Choose your order Excel file",
            type=['xlsx', 'xls'],
            key="order_file",
            help="This should contain Item and Order Quantity columns"
        )
        
        if order_file is not None:
            st.success(f"✅ Order file uploaded: {order_file.name}")
            st.info(f"File size: {order_file.size:,} bytes")
    
    # Process files when both are uploaded
    if main_file is not None and order_file is not None:
        st.markdown("---")
        
        # Add a big, prominent process button
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            process_button = st.button(
                "🚀 Process Files", 
                type="primary", 
                use_container_width=True,
                help="Click to process both files and combine the data"
            )
        
        if process_button:
            # Create progress bar
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            try:
                # Step 1: Save files temporarily
                status_text.text("📁 Saving uploaded files...")
                progress_bar.progress(10)
                
                with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_main:
                    tmp_main.write(main_file.getvalue())
                    main_path = tmp_main.name
                
                with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_order:
                    tmp_order.write(order_file.getvalue())
                    order_path = tmp_order.name
                
                # Step 2: Process the files
                status_text.text("⚙️ Processing files...")
                progress_bar.progress(30)
                
                result = st.session_state.processor.process_files(main_path, order_path)
                progress_bar.progress(80)
                
                if result['success']:
                    status_text.text("✅ Processing completed successfully!")
                    progress_bar.progress(100)
                    
                    st.success("🎉 Files processed successfully!")
                    
                    # Display summary in attractive cards
                    st.subheader("📈 Processing Summary")
                    
                    col1, col2, col3, col4 = st.columns(4)
                    
                    with col1:
                        st.metric(
                            label="📦 Total Items", 
                            value=f"{result['total_items']:,}",
                            help="Total number of items processed from all sheets"
                        )
                    with col2:
                        st.metric(
                            label="✅ Items with Order Qty", 
                            value=f"{result['matched_items']:,}",
                            help="Items that had matching order quantities found"
                        )
                    with col3:
                        st.metric(
                            label="🎯 Match Rate", 
                            value=f"{result['match_rate']:.1f}%",
                            help="Percentage of items that got order quantities"
                        )
                    with col4:
                        missing_items = result['total_items'] - result['matched_items']
                        st.metric(
                            label="❌ Missing Qty", 
                            value=f"{missing_items:,}",
                            help="Items without order quantities"
                        )
                    
                    # Download section
                    st.markdown("---")
                    st.subheader("📥 Download Results")
                    
                    # Create Excel file in memory
                    output_buffer = BytesIO()
                    with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
                        result['summary_df'].to_excel(writer, sheet_name='Summary', index=False)
                        result['combined_df'].to_excel(writer, sheet_name='Combined', index=False)
                    
                    # Download button
                    col1, col2, col3 = st.columns([1, 2, 1])
                    with col2:
                        st.download_button(
                            label="📥 Download Processed Excel File",
                            data=output_buffer.getvalue(),
                            file_name=f"{main_file.name.split('.')[0]}_processed.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )
                    
                    # Data preview section
                    st.markdown("---")
                    st.subheader("👀 Data Preview")
                    
                    # Show tabs for different views
                    tab1, tab2, tab3 = st.tabs(["📊 Combined Data", "📋 Summary Sheet", "🔍 Sample Matches"])
                    
                    with tab1:
                        st.write("**First 20 rows of combined data:**")
                        st.dataframe(
                            result['combined_df'].head(20), 
                            use_container_width=True,
                            height=400
                        )
                        
                        if len(result['combined_df']) > 20:
                            st.info(f"Showing first 20 rows out of {len(result['combined_df']):,} total rows")
                    
                    with tab2:
                        st.write("**Summary sheet data:**")
                        st.dataframe(
                            result['summary_df'].head(10), 
                            use_container_width=True,
                            height=300
                        )
                    
                    with tab3:
                        st.write("**Sample of items with and without order quantities:**")
                        
                        # Show items with matches
                        items_with_qty = result['combined_df'][
                            result['combined_df']['Ordered Qty'].notna() & 
                            (result['combined_df']['Ordered Qty'] != "") & 
                            (result['combined_df']['Ordered Qty'] != 0)
                        ][['Item Number', 'Item Description', 'Ordered Qty']].head(5)
                        
                        if not items_with_qty.empty:
                            st.write("✅ **Items WITH order quantities:**")
                            st.dataframe(items_with_qty, use_container_width=True)
                        
                        # Show items without matches
                        items_without_qty = result['combined_df'][
                            result['combined_df']['Ordered Qty'].isna() | 
                            (result['combined_df']['Ordered Qty'] == "") | 
                            (result['combined_df']['Ordered Qty'] == 0)
                        ][['Item Number', 'Item Description', 'Ordered Qty']].head(5)
                        
                        if not items_without_qty.empty:
                            st.write("❌ **Items WITHOUT order quantities:**")
                            st.dataframe(items_without_qty, use_container_width=True)
                
                else:
                    status_text.text("❌ Processing failed")
                    progress_bar.progress(0)
                    st.error(f"❌ Processing Error: {result['error']}")
                    
                    # Show helpful error information
                    st.subheader("🔍 Troubleshooting Tips")
                    st.markdown("""
                    **Common issues and solutions:**
                    
                    1. **Order file columns not found:**
                       - Make sure your order file has columns named 'Item' and 'Order Quantity' (or similar)
                       - Check that column headers are in the first few rows
                    
                    2. **Main file missing required columns:**
                       - Ensure data sheets have: Planner, Published, Item Number, Item Description, Oracle On Hand
                       - Check that there's a 'Summary' sheet with 'Issue key' and 'Summary' columns
                    
                    3. **No data processed:**
                       - Verify that your files are valid Excel files (.xlsx or .xls)
                       - Check that sheets contain actual data, not just headers
                    """)
                
                # Clean up temp files
                try:
                    os.unlink(main_path)
                    os.unlink(order_path)
                except:
                    pass  # Ignore cleanup errors
                    
            except Exception as e:
                status_text.text("❌ An unexpected error occurred")
                progress_bar.progress(0)
                st.error(f"❌ An unexpected error occurred: {str(e)}")
                
                # Clean up temp files even on error
                try:
                    os.unlink(main_path)
                    os.unlink(order_path)
                except:
                    pass
    
    # Footer
    st.markdown("---")
    st.markdown(
        """
        <div style='text-align: center; color: #666;'>
        📊 Excel Data Processor | Built with Streamlit | For questions, contact your IT team
        </div>
        """, 
        unsafe_allow_html=True
    )

if __name__ == "__main__":
    main()
