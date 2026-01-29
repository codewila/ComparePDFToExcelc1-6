import streamlit as st
import pandas as pd
import pdfplumber
import PyPDF2
from io import BytesIO
import re
from typing import Dict
from difflib import SequenceMatcher

# Set page configuration
st.set_page_config(
    page_title="Excel-PDF Color Matcher",
    page_icon="🎨",
    layout="wide"
)

# --- CSS for better table visibility on screen ---
st.markdown("""
<style>
    .stDataFrame { font-size: 14px; }
</style>
""", unsafe_allow_html=True)

def extract_text_from_pdf(pdf_file) -> Dict[int, str]:
    """Extract text from PDF pages"""
    text_dict = {}
    try:
        with pdfplumber.open(pdf_file) as pdf:
            for i, page in enumerate(pdf.pages):
                text = page.extract_text()
                if text:
                    text_dict[i + 1] = text
    except Exception as e:
        st.error(f"Error extracting text with pdfplumber: {e}")
        try:
            pdf_reader = PyPDF2.PdfReader(pdf_file)
            for i, page in enumerate(pdf_reader.pages):
                text = page.extract_text()
                if text:
                    text_dict[i + 1] = text
        except Exception as e2:
            st.error(f"Error extracting text: {e2}")
    return text_dict

def check_value_in_pdf(value, pdf_text_dict, threshold=0.8):
    """Check if value exists in PDF. Returns: (is_match, match_type)"""
    search_value = str(value).strip()
    
    if not search_value or search_value.lower() in ['nan', 'none', '', 'nat']:
        return False, 'Empty'

    search_lower = search_value.lower()

    for page_num, text in pdf_text_dict.items():
        clean_text = re.sub(r'\s+', ' ', text).lower()
        
        if search_lower in clean_text:
            return True, 'Exact'
            
        if len(search_lower) > 3:
            words = re.findall(r'\b\w+\b', clean_text)
            for word in words:
                if len(word) > 3:
                    similarity = SequenceMatcher(None, search_lower, word).ratio()
                    if similarity >= threshold:
                        return True, 'Fuzzy'
                        
    return False, 'No Match'

def generate_excel_with_colors(df, status_df):
    """Generates an Excel file where cells are colored based on status"""
    output = BytesIO()
    
    # Use xlsxwriter engine for styling
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        # Write the raw data first
        df.to_excel(writer, sheet_name='Analysis_Result', index=False)
        
        # Get workbook and worksheet objects
        workbook = writer.book
        worksheet = writer.sheets['Analysis_Result']
        
        # Define formats (Green for Match, Red for No Match)
        green_format = workbook.add_format({
            'bg_color': '#90EE90', # Light Green
            'font_color': '#000000',
            'border': 1
        })
        
        red_format = workbook.add_format({
            'bg_color': '#FFB6C1', # Light Red (Halka Red)
            'font_color': '#000000',
            'border': 1
        })
        
        default_format = workbook.add_format({'border': 1})

        # Iterate over the DataFrame to apply colors cell by cell
        # Note: xlsxwriter is 0-indexed. 
        # Row 0 is the header in the Excel file. Data starts at Row 1.
        for row_idx, row in df.iterrows():
            for col_idx, value in enumerate(row):
                # Get the status for this cell
                col_name = df.columns[col_idx]
                status = status_df.at[row_idx, col_name]
                
                # Apply format based on status
                # (row_idx + 1 because row 0 is header)
                if status in ['Exact', 'Fuzzy']:
                    worksheet.write(row_idx + 1, col_idx, value, green_format)
                elif status == 'No Match':
                    worksheet.write(row_idx + 1, col_idx, value, red_format)
                else:
                    # Write empty/other cells nicely with border but no color
                    if pd.isna(value):
                        worksheet.write(row_idx + 1, col_idx, "", default_format)
                    else:
                        worksheet.write(row_idx + 1, col_idx, value, default_format)
                        
    return output.getvalue()

def main():
    st.title("📊 Excel-PDF Color Matcher (Downloadable)")
    st.markdown("""
    Upload Excel & PDF. Match values. **Download colored Excel.**
    - 🟢 **Green**: Match Found
    - 🔴 **Red**: No Match Found
    """)

    with st.sidebar:
        st.header("⚙️ Settings")
        excel_file = st.file_uploader("Upload Excel", type=['xlsx', 'xls', 'csv'])
        pdf_file = st.file_uploader("Upload PDF", type=['pdf'])
        threshold = st.slider("Match Accuracy", 0.6, 1.0, 0.85)
        file_type = st.radio("File Type", ['Excel', 'CSV'])

    if excel_file and pdf_file:
        if st.button("🔍 Start Matching", type="primary", use_container_width=True):
            try:
                # Load Data
                if file_type == 'Excel':
                    df = pd.read_excel(excel_file)
                else:
                    df = pd.read_csv(excel_file)

                with st.spinner("Analyzing..."):
                    pdf_text = extract_text_from_pdf(pdf_file)
                    if not pdf_text:
                        st.error("No text found in PDF.")
                        st.stop()

                    # Create status DF
                    status_df = pd.DataFrame(index=df.index, columns=df.columns)
                    
                    # Process Cells
                    progress_bar = st.progress(0)
                    total_cells = df.size
                    done = 0
                    
                    for col in df.columns:
                        for idx in df.index:
                            val = df.at[idx, col]
                            _, status = check_value_in_pdf(val, pdf_text, threshold)
                            status_df.at[idx, col] = status
                            done += 1
                            if done % 50 == 0:
                                progress_bar.progress(min(done/total_cells, 1.0))
                    
                    progress_bar.progress(1.0)

                    # Show on Screen (Visual Preview)
                    def style_screen(val):
                        # This is just for the browser view, logic is disconnected from download
                        return '' 
                    
                    # Apply colors for browser view
                    def highlight_cells(data):
                        df_colors = pd.DataFrame('', index=data.index, columns=data.columns)
                        for col in data.columns:
                            for idx in data.index:
                                s = status_df.at[idx, col]
                                if s in ['Exact', 'Fuzzy']:
                                    df_colors.at[idx, col] = 'background-color: #90EE90'
                                elif s == 'No Match':
                                    df_colors.at[idx, col] = 'background-color: #FFB6C1'
                        return df_colors

                    st.subheader("👀 Preview Results")
                    st.dataframe(df.style.apply(highlight_cells, axis=None), use_container_width=True)

                    # Generate Colored Excel for Download
                    excel_data = generate_excel_with_colors(df, status_df)
                    
                    st.success("Analysis Complete!")
                    st.download_button(
                        label="📥 Download Colored Excel Report",
                        data=excel_data,
                        file_name="Colored_Match_Report.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )

            except Exception as e:
                st.error(f"Error: {e}")

if __name__ == "__main__":
    main()
