import streamlit as st
import pandas as pd
import io
import zipfile
from datetime import datetime
import xlsxwriter # Ensure this is in requirements.txt

# --- Configuration ---
# Set the page configuration
st.set_page_config(
    page_title="Universal Publisher Report Generator",
    layout="centered",
    initial_sidebar_state="collapsed",
)

# --- Core Processing Function ---

def process_excel_and_create_reports(uploaded_file, original_filename):
    """
    Reads all sheets, filters by 'Publisher' column, groups data, and 
    creates a dictionary of publisher-specific, multi-sheet Excel files.
    """
    
    # 1. Read all sheets from the uploaded Excel file
    st.info(f"Reading all sheets from **{original_filename}**...")
    
    try:
        # Use sheet_name=None to read all sheets into an OrderedDict of DataFrames
        all_sheets_dict = pd.read_excel(uploaded_file, sheet_name=None)
    except Exception as e:
        st.error(f"Failed to read Excel sheets: {e}")
        return {}

    # 2. Identify and filter sheets that contain the 'Publisher' column
    valid_sheets = {}
    st.subheader("Sheet Review:")
    
    for sheet_name, df in all_sheets_dict.items():
        if 'Publisher' in df.columns:
            valid_sheets[sheet_name] = df
            st.markdown(f"- ✅ Sheet **'{sheet_name}'** contains 'Publisher' column.")
        else:
            st.markdown(f"- ❌ Sheet **'{sheet_name}'** ignored (No 'Publisher' column).")

    if not valid_sheets:
        st.warning("No sheets contained the required 'Publisher' column. Processing halted.")
        return {}

    # 3. Consolidate all 'Publisher' columns to get the unique list
    # We use the 'Publisher' column from all valid sheets to ensure we capture all of them
    publisher_series = [df['Publisher'] for df in valid_sheets.values()]
    all_publishers = pd.concat(publisher_series).unique()

    if len(all_publishers) == 0:
        st.warning("Unique publishers list is empty. Check data quality.")
        return {}

    st.success(f"\nFound **{len(all_publishers)}** unique Publishers across valid sheets.")
    st.caption(f"Valid sheets being processed: {', '.join(valid_sheets.keys())}")

    # 4. Loop through each unique publisher to create their specific multi-sheet Excel file
    publisher_files = {}
    report_date = datetime.now().strftime("%Y%m%d_%H%M%S") # Current date/time for naming
    
    for i, publisher in enumerate(all_publishers):
        
        # Clean the publisher name for use as a filename
        safe_publisher_name = "".join(c for c in str(publisher) if c.isalnum() or c in (' ', '_', '-')).rstrip()
        
        # Build the strictly defined output filename
        output_filename = f"{safe_publisher_name}_{original_filename.replace('.xlsx', '')}_{report_date}.xlsx"
        
        # Use an in-memory buffer (BytesIO) for the Excel file
        output = io.BytesIO()
        
        # Use ExcelWriter to write multiple sheets to one file
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            
            # Write the data for THIS specific publisher across ALL valid sheets
            for sheet_name, original_df in valid_sheets.items():
                
                # Filter data for the current publisher in the current sheet
                publisher_data = original_df[original_df['Publisher'] == publisher]
                
                if not publisher_data.empty:
                    # Write the filtered data to a sheet with its original name
                    publisher_data.to_excel(writer, sheet_name=sheet_name, index=False)
        
        # Save the bytes content to the dictionary
        publisher_files[output_filename] = output.getvalue()
        
        st.progress((i + 1) / len(all_publishers), text=f"Processing: {publisher}")

    return publisher_files

# --- Streamlit Application Layout ---

st.title("📄 Publisher Report Generator")
st.markdown("---")
st.info("""
**Instructions:** Upload your Excel file. The script will look for the **'Publisher'** column 
in every sheet. It will then create one combined Excel file per unique Publisher, 
containing all the original sheets that had the 'Publisher' column.
""")

# File Uploader Widget
uploaded_file = st.file_uploader(
    "1. Choose your Excel file:", 
    type=['xlsx']
)

if uploaded_file is not None:
    
    # Extract original filename for the naming convention
    original_filename = uploaded_file.name
    
    st.markdown("### 2. Processing File...")
    
    with st.spinner("Analyzing and preparing reports..."):
        # Run the core processing function
        reports_data = process_excel_and_create_reports(uploaded_file, original_filename)
        
    num_files = len(reports_data)
        
    if num_files > 0:
        st.success(f"✅ Successfully generated **{num_files}** individual Publisher reports!")

        # --- ZIP FILE CREATION FOR DOWNLOAD ---
        
        st.markdown("### 3. Download Reports")
        
        # Create a new in-memory zip file
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, 'a', zipfile.ZIP_DEFLATED, False) as zip_file:
            for filename, file_content in reports_data.items():
                # Add each individual excel report to the zip file
                zip_file.writestr(filename, file_content)
        
        # Reset the buffer position to the beginning before downloading
        zip_buffer.seek(0)

        st.download_button(
            label=f"⬇️ **Download All {num_files} Reports (.zip)**",
            data=zip_buffer,
            file_name="All_Publisher_Reports.zip",
            mime="application/zip",
            help="Click to download a ZIP file containing all the individual Excel reports, named by the specific convention.",
            key='download_zip'
        )
        
        st.markdown("---")
        st.caption("Each file strictly follows the naming convention: `[Publisher]_[Original File Name]_[Date of Generation]`.")

    else:
        st.error("No reports were generated. Please check the uploaded file and sheet contents.")
