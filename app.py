import streamlit as st
import pandas as pd
import io
import os
import xlsxwriter

# Set the page configuration
st.set_page_config(
    page_title="Publisher Report Generator",
    layout="centered",
    initial_sidebar_state="collapsed",
)

# --- Core Processing Function (Modified from Colab) ---
# This function now takes a dictionary of DataFrames and returns the final zip file content
def process_excel_and_create_reports(df_crm, df_adjust):
    """
    Groups data by Publisher from CRM and Adjust DataFrames and creates
    a dictionary where keys are filenames and values are the file contents (bytes).
    """
    
    # 1. Get all unique Publishers across both sheets
    all_publishers = pd.concat([df_crm['Publisher'], df_adjust['Publisher']]).unique()
    
    # Dictionary to hold the data for the final zip file
    publisher_files = {}

    for publisher in all_publishers:
        
        # 2. Filter data for the current publisher for both sheets
        crm_data = df_crm[df_crm['Publisher'] == publisher]
        adjust_data = df_adjust[df_adjust['Publisher'] == publisher]
        
        # 3. Clean the publisher name for use as a filename
        safe_publisher_name = "".join(c for c in str(publisher) if c.isalnum() or c in (' ', '_')).rstrip()
        output_filename = f"{safe_publisher_name}_Report.xlsx"
        
        # 4. Use an in-memory buffer (BytesIO) to create the Excel file
        output = io.BytesIO()
        
        # Save the data to the in-memory buffer with two sheets
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            if not crm_data.empty:
                crm_data.to_excel(writer, sheet_name='CRM', index=False)
            
            if not adjust_data.empty:
                adjust_data.to_excel(writer, sheet_name='Adjust', index=False)
        
        # Save the bytes content to the dictionary
        publisher_files[output_filename] = output.getvalue()

    return publisher_files

# --- Streamlit Application Layout ---

st.title("📄 Excel Publisher Report Generator")
st.markdown("---")
st.info("Upload your Excel file to generate separate reports for each Publisher.")

# File Uploader Widget
uploaded_file = st.file_uploader(
    "1. Choose your Excel file:", 
    type=['xlsx']
)

if uploaded_file is not None:
    
    st.markdown("### 2. Processing File...")
    
    try:
        # Read the two sheets directly from the uploaded file
        df_crm = pd.read_excel(uploaded_file, sheet_name='CRM')
        df_adjust = pd.read_excel(uploaded_file, sheet_name='Adjust')
        
        # Check if the mandatory columns exist
        if 'Publisher' not in df_crm.columns or 'Publisher' not in df_adjust.columns:
            st.error("🚨 Error: The uploaded sheets must contain a column named **'Publisher'**.")
        else:
            with st.spinner("Analyzing Publishers and generating reports..."):
                
                # Run the core processing function
                reports_data = process_excel_and_create_reports(df_crm, df_adjust)
                
                num_files = len(reports_data)
                
            if num_files > 0:
                st.success(f"✅ Successfully generated **{num_files}** individual Publisher reports!")

                # --- ZIP FILE CREATION FOR DOWNLOAD ---
                # A single ZIP is the easiest way to download multiple files
                
                import zipfile
                
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
                    help="Click to download a ZIP file containing all the individual Excel reports.",
                    key='download_zip'
                )
                
                st.markdown("---")
                st.markdown("The ZIP file contains a separate Excel file for every Publisher, each with the 'CRM' and 'Adjust' sheets.")

            else:
                st.warning("No unique Publishers found in the provided data. Please check the 'Publisher' column.")

    except Exception as e:
        st.error(f"An unexpected error occurred: {e}")
        st.error("Please ensure your Excel file has sheets named **'CRM'** and **'Adjust'**.")
