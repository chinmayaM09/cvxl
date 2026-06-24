import subprocess
import sys

# FORCE INSTALL - This runs before anything else
subprocess.check_call([sys.executable, "-m", "pip", "install", "openpyxl", "-q"])

import streamlit as st
import pandas as pd
import io

# --- Page Config ---
st.set_page_config(
    page_title="CSV to Excel Converter",
    page_icon="🔄",
    layout="centered"
)

# --- Header ---
st.title("🔄 CSV to Excel Converter")
st.markdown("Upload a CSV file and download it as Excel (.xlsx)")

st.divider()

# --- File Upload ---
uploaded_file = st.file_uploader(
    "Choose a CSV file",
    type=["csv"],
    help="Supported format: .csv"
)

if uploaded_file is not None:
    try:
        # Read CSV
        df = pd.read_csv(uploaded_file)
        
        # Show preview
        st.subheader("📄 Preview")
        st.write(f"**Rows:** {len(df)} | **Columns:** {len(df.columns)}")
        
        st.dataframe(df.head(10), use_container_width=True, hide_index=True)
        
        if len(df) > 10:
            st.caption(f"Showing first 10 of {len(df)} rows")
        
        st.divider()
        
        # --- Options ---
        st.subheader("⚙️ Options")
        col1, col2 = st.columns(2)
        
        with col1:
            sheet_name = st.text_input("Sheet Name", value="Data")
        
        with col2:
            include_index = st.checkbox("Include Row Numbers", value=False)
        
        st.divider()
        
        # --- Convert & Download ---
        st.subheader("📥 Download")
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(
                writer,
                index=include_index,
                sheet_name=sheet_name
            )
        output.seek(0)
        
        # Generate filename
        original_name = uploaded_file.name.replace(".csv", "")
        excel_filename = f"{original_name}.xlsx"
        
        st.download_button(
            label="⬇️ Download Excel File",
            data=output.getvalue(),
            file_name=excel_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )
        
        st.success(f"✅ Ready to download: {excel_filename}")
        
    except pd.errors.EmptyDataError:
        st.error("❌ The CSV file is empty.")
    except pd.errors.ParserError:
        st.error("❌ Could not parse the CSV file. Please check the format.")
    except Exception as e:
        st.error(f"❌ An error occurred: {str(e)}")

else:
    st.info("👆 Upload a CSV file to get started")
    
    with st.expander("📋 Supported Features"):
        st.markdown("""
        - **Encoding:** UTF-8, Latin-1, etc. (auto-detected)
        - **Delimiters:** Comma, semicolon, tab (auto-detected)
        - **Large Files:** Handles files with 100,000+ rows
        - **Output:** Excel .xlsx format
        """)

st.divider()
st.caption("Built with Streamlit & Pandas")
