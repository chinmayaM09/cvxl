import streamlit as st
import pandas as pd
import zipfile
import io
import datetime

# ============================================
# XLSX BUILDER (Zero external dependencies)
# ============================================

def format_cell(val):
    """Format cell value as string."""
    if pd.isna(val):
        return ""
    if isinstance(val, (datetime.datetime, datetime.date)):
        return val.isoformat()
    return str(val)


def get_column_letter(idx):
    """0,1,2... -> A,B,C..."""
    result = ""
    idx += 1
    while idx > 0:
        idx, remainder = divmod(idx - 1, 26)
        result = chr(65 + remainder) + result
    return result


def build_shared_strings(strings):
    xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    xml += f'<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="{len(strings)}" uniqueCount="{len(strings)}">'
    for s in strings:
        escaped = s.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
        xml += f'<si><t>{escaped}</t></si>'
    xml += '</sst>'
    return xml


def build_sheet(rows, string_map):
    xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    xml += '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
    xml += '<sheetData>'
    for row_idx, row in enumerate(rows, 1):
        xml += f'<row r="{row_idx}">'
        for col_idx, cell in enumerate(row):
            col_letter = get_column_letter(col_idx)
            cell_ref = f"{col_letter}{row_idx}"
            str_idx = string_map[cell]
            xml += f'<c r="{cell_ref}" t="s"><v>{str_idx}</v></c>'
        xml += '</row>'
    xml += '</sheetData></worksheet>'
    return xml


CONTENT_TYPES = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
</Types>'''

RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>'''

WORKBOOK = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>'''

WORKBOOK_RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
</Relationships>'''


def dataframe_to_xlsx(df, include_index=False):
    """Convert DataFrame to xlsx bytes using ZERO external packages."""
    if include_index:
        df = df.reset_index()

    # Build rows
    rows = [[str(c) for c in df.columns]]
    for _, row in df.iterrows():
        rows.append([format_cell(v) for v in row])

    # Build shared strings lookup
    string_map = {}
    shared_strings = []
    for row in rows:
        for cell in row:
            if cell not in string_map:
                string_map[cell] = len(shared_strings)
                shared_strings.append(cell)

    # Pack into xlsx (which is a zip file)
    output = io.BytesIO()
    with zipfile.ZipFile(output, 'w', zipfile.ZIP_DEFLATED) as zf:
        zf.writestr('[Content_Types].xml', CONTENT_TYPES)
        zf.writestr('_rels/.rels', RELS)
        zf.writestr('xl/workbook.xml', WORKBOOK)
        zf.writestr('xl/_rels/workbook.xml.rels', WORKBOOK_RELS)
        zf.writestr('xl/sharedStrings.xml', build_shared_strings(shared_strings))
        zf.writestr('xl/worksheets/sheet1.xml', build_sheet(rows, string_map))

    output.seek(0)
    return output.getvalue()


# ============================================
# STREAMLIT APP
# ============================================

st.set_page_config(page_title="CSV to Excel", page_icon="🔄", layout="centered")

st.title("🔄 CSV to Excel Converter")
st.markdown("Upload CSV → Download Excel. No external packages needed.")
st.divider()

uploaded_file = st.file_uploader("Choose a CSV file", type=["csv"])

if uploaded_file is not None:
    try:
        df = pd.read_csv(uploaded_file)

        st.subheader("📄 Preview")
        st.write(f"**Rows:** {len(df)} | **Columns:** {len(df.columns)}")
        st.dataframe(df.head(10), use_container_width=True, hide_index=True)
        if len(df) > 10:
            st.caption(f"Showing first 10 of {len(df)} rows")

        st.divider()

        include_index = st.checkbox("Include Row Numbers", value=False)
        st.divider()

        # Convert using our built-in xlsx builder
        xlsx_bytes = dataframe_to_xlsx(df, include_index=include_index)

        filename = uploaded_file.name.replace(".csv", ".xlsx")

        st.download_button(
            label="⬇️ Download Excel File",
            data=xlsx_bytes,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )
        st.success(f"✅ Ready: {filename}")

    except pd.errors.EmptyDataError:
        st.error("❌ The CSV file is empty.")
    except pd.errors.ParserError:
        st.error("❌ Could not parse the CSV. Check the format.")
    except Exception as e:
        st.error(f"❌ Error: {str(e)}")

else:
    st.info("👆 Upload a CSV file to get started")

st.divider()
st.caption("Built with Python built-in modules only — zero dependencies")
