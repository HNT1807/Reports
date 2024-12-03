import streamlit as st
import pandas as pd
import io
import re
from typing import List, Dict, Set
import zipfile


def get_column_letter(index):
    """Convert number to Excel column letter (0=A, 1=B, etc)"""
    letter = ''
    while index >= 0:
        letter = chr(65 + (index % 26)) + letter
        index = index // 26 - 1
    return letter


def normalize_title(s):
    """Remove parentheses and their content, then normalize text"""
    if pd.isna(s):
        return ""

    s = str(s)
    # Remove parentheses and their content
    while '(' in s and ')' in s:
        start = s.find('(')
        end = s.find(')') + 1
        if start < end:
            s = s[:start] + s[end:]  # This removes both parentheses and their content
        else:
            break

    return s.lower().strip().replace(" ", "")
def get_composer_words(s):
    """Get all words from composer name after cleaning"""
    if pd.isna(s):
        return [], ""

    s = str(s)
    # Remove parentheses and percentages
    s = re.sub(r'\([^)]*\)', '', s)
    s = re.sub(r'\d+%', '', s)
    # Remove special characters
    s = re.sub(r'[^a-zA-Z0-9\s]', ' ', s)
    # Remove extra spaces
    s = ' '.join(s.split())

    words = s.lower().split()
    return words, s.lower()


def process_files(ww_codes_file, tango_file, report_files):
    try:
        # Initialize lookup dictionary
        ww_lookup = {}

        # Determine which source we're using
        using_ww_codes = ww_codes_file is not None
        using_tango = tango_file is not None

        # Create the column mapping starting from AG (index 32)
        code_columns = []
        header_mapping = {}  # Dictionary for header mapping
        start_idx = 32  # Column AG
        ww_code_columns = []  # Track WW code columns for highlighting

        # Different column structure based on source
        for i in range(50):
            if using_ww_codes and not using_tango:
                # Original WW Codes structure
                base_idx = start_idx + (i * 5)
                columns = [
                    get_column_letter(base_idx),  # AG: WW Code
                    get_column_letter(base_idx + 1),  # AH: BMI Code
                    get_column_letter(base_idx + 2),  # AI: ASCAP Code
                    get_column_letter(base_idx + 3),  # AJ: SESAC Code
                    get_column_letter(base_idx + 4),  # AK: Publisher
                ]

                # Original WW Codes headers
                header_mapping[columns[0]] = f"Snowflakes - WW Code {i + 1}"
                header_mapping[columns[1]] = f"Snowflakes - BMI Code {i + 1}"
                header_mapping[columns[2]] = f"Snowflakes - ASCAP Code {i + 1}"
                header_mapping[columns[3]] = f"Snowflakes - SESAC Code {i + 1}"
                header_mapping[columns[4]] = f"Snowflakes - Publisher {i + 1}"

            else:
                # Tango Export structure
                base_idx = start_idx + (i * 5)
                columns = [
                    get_column_letter(base_idx),  # AG: WW Code
                    get_column_letter(base_idx + 1),  # AH: Publisher 1
                    get_column_letter(base_idx + 2),  # AI: Publisher 1 Code
                    get_column_letter(base_idx + 3),  # AJ: Publisher 2
                    get_column_letter(base_idx + 4),  # AK: Publisher 2 Code
                ]

                # Tango headers
                header_mapping[columns[0]] = f"Tango - WW Code {i + 1}"
                header_mapping[columns[1]] = f"Tango - Publisher 1 ({i + 1})"
                header_mapping[columns[2]] = f"Tango - Publisher 1 Code ({i + 1})"
                header_mapping[columns[3]] = f"Tango - Publisher 2 ({i + 1})"
                header_mapping[columns[4]] = f"Tango - Publisher 2 Code ({i + 1})"

            code_columns.extend(columns)
            ww_code_columns.append(columns[0])

        # Process WW Codes file if provided
        if using_ww_codes:
            ww_codes_df = pd.read_excel(ww_codes_file, engine='openpyxl', dtype=str).fillna('')

            for _, row in ww_codes_df.iterrows():
                title = normalize_title(row.iloc[0])
                if row.iloc[1] == '':
                    continue
                if title not in ww_lookup:
                    ww_lookup[title] = []
                ww_lookup[title].append({
                    'source': 'ww_codes',
                    'original_title': row.iloc[0],
                    'ww_code': row.iloc[1],
                    'composer': row.iloc[2],
                    'publisher': row.iloc[3],
                    'bmi_code': row.iloc[4],
                    'ascap_code': row.iloc[5],
                    'sesac_code': row.iloc[6]
                })

        # Process Tango export if provided
        if using_tango:
            tango_df = pd.read_excel(tango_file, engine='openpyxl', dtype=str).fillna('')

            # Filter out rows where column Q contains "Other (NP)"
            tango_df = tango_df[tango_df.iloc[:, 16] != "Other (NP)"]  # Column Q is index 16 (0-based)

            # Group by track title to handle multiple publishers
            grouped = tango_df.groupby(tango_df.iloc[:, 2])  # Group by track title (column C)

            for title, group in grouped:
                normalized_title = normalize_title(title)
                if normalized_title not in ww_lookup:
                    ww_lookup[normalized_title] = []

                # Get unique WW codes for this title
                ww_codes = group.iloc[:, 1].unique()  # WW Code (column B)

                for ww_code in ww_codes:
                    # Get all rows for this WW code
                    ww_rows = group[group.iloc[:, 1] == ww_code]

                    publishers = []
                    publisher_codes = []
                    composer = ww_rows.iloc[0, 4]  # Take composer from first row (column E)

                    for _, row in ww_rows.iterrows():
                        publishers.append(row.iloc[8])  # Publisher (column I)
                        publisher_codes.append(row.iloc[22])  # Publisher code (column W)

                    ww_lookup[normalized_title].append({
                        'source': 'tango',
                        'original_title': title,
                        'ww_code': ww_code,
                        'composer': composer,
                        'publishers': publishers,
                        'publisher_codes': publisher_codes
                    })

        all_matches = []
        processed_reports = []
        total_updates = 0

        for report_file in report_files:
            try:
                report_df = pd.read_excel(report_file, engine='openpyxl', dtype=str).fillna('')

                # Initialize columns if they don't exist
                for col in code_columns:
                    if col not in report_df.columns:
                        report_df[col] = ''

                style_df = pd.DataFrame('', index=report_df.index, columns=report_df.columns)

                for idx, row in report_df.iterrows():
                    title = normalize_title(str(row.iloc[4]))
                    composer = str(row.iloc[5])
                    # Use get_composer_words instead of get_first_two_words
                    report_words, _ = get_composer_words(composer)

                    if title in ww_lookup:
                        used_codes = set()
                        matches_found = 0

                        for entry in ww_lookup[title]:
                            composer_match = False

                            # Clean the entry composer string (whether from Tango or WW Codes)
                            entry_composer = str(entry['composer']).lower()
                            entry_composer = re.sub(r'\([^)]*\)', '', entry_composer)
                            entry_composer = re.sub(r'\d+%', '', entry_composer)
                            composer_parts = re.split(r'[|,;&/]', entry_composer)

                            # Clean each part
                            cleaned_parts = []
                            for part in composer_parts:
                                cleaned = re.sub(r'[^a-zA-Z0-9\s]', ' ', part)
                                cleaned = ' '.join(cleaned.split())
                                if cleaned:
                                    cleaned_parts.append(cleaned)

                            # Join all parts with space
                            full_composer = ' '.join(cleaned_parts)
                            full_composer_words = set(full_composer.split())

                            # Count how many words from report appear in entry composer
                            matching_words = sum(1 for word in report_words if word in full_composer_words)

                            # If we have at least 2 matching words, it's a match
                            if matching_words >= 2:
                                composer_match = True

                            if composer_match and entry['ww_code'] not in used_codes:
                                base_idx = matches_found * 5
                                if base_idx + 4 < len(code_columns):
                                    cols = code_columns[base_idx:base_idx + 5]

                                    # Add WW Code
                                    report_df.loc[idx, cols[0]] = str(entry['ww_code'])
                                    style_df.loc[idx, cols[0]] = 'background-color: lightgreen'

                                    if entry['source'] == 'tango':
                                        # Handle Tango export data structure
                                        if len(entry['publishers']) > 0:
                                            report_df.loc[idx, cols[1]] = entry['publishers'][0]
                                            report_df.loc[idx, cols[2]] = entry['publisher_codes'][0]
                                        if len(entry['publishers']) > 1:
                                            report_df.loc[idx, cols[3]] = entry['publishers'][1]
                                            report_df.loc[idx, cols[4]] = entry['publisher_codes'][1]
                                    else:
                                        # Handle WW Codes data structure
                                        report_df.loc[idx, cols[1]] = str(entry['bmi_code'])
                                        report_df.loc[idx, cols[2]] = str(entry['ascap_code'])
                                        report_df.loc[idx, cols[3]] = str(entry['sesac_code'])
                                        report_df.loc[idx, cols[4]] = str(entry['publisher'])

                                    all_matches.append({
                                        'File': report_file.name,
                                        'Row': idx + 1,
                                        'Report Title': row.iloc[4],
                                        'Report Composer': composer,
                                        'Match Title': entry['original_title'],
                                        'Match Composer': entry['composer'],
                                        'WW Code': str(entry['ww_code']),
                                        'Source': entry['source']
                                    })

                                    used_codes.add(entry['ww_code'])
                                    matches_found += 1
                                    total_updates += 1

                # After processing all rows, rename the columns and apply styling
                report_df = report_df.rename(columns=header_mapping)
                style_df = style_df.rename(columns=header_mapping)
                report_df = report_df.fillna('')
                styled_df = report_df.style.apply(lambda _: style_df, axis=None)

                processed_reports.append({
                    'name': report_file.name,
                    'df': styled_df
                })

            except Exception as e:
                st.error(f"Error processing {report_file.name}: {str(e)}")
                continue

        return processed_reports, all_matches, total_updates

    except Exception as e:
        st.error(f"Error: {str(e)}")
        return None, None, 0

def create_download_zip(reports):
    """Create a zip file containing all processed reports"""
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
        for report in reports:
            excel_buffer = io.BytesIO()
            with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                # Convert any remaining NaN values to empty strings before writing
                df_to_write = report['df'].data.fillna('')
                df_to_write.to_excel(writer, index=False)

                # Get the worksheet
                worksheet = writer.sheets['Sheet1']

                # Format code columns as text and apply styling
                for row_idx, row in enumerate(report['df'].index,
                                              start=2):  # start=2 because Excel is 1-based and we have headers
                    for col_idx, col in enumerate(report['df'].columns, start=1):
                        cell = worksheet.cell(row=row_idx, column=col_idx)

                        # Check if this is a code column (WW, BMI, ASCAP, or SESAC)
                        if any(code in col for code in ['WW Code', 'BMI Code', 'ASCAP Code', 'SESAC Code']):
                            # Format as text
                            cell.number_format = '@'
                            # If there's a value and it's not 'nan', ensure it's treated as text
                            if cell.value and str(cell.value).lower() != 'nan':
                                cell.value = str(cell.value)
                            else:
                                cell.value = ''  # Replace 'nan' with empty string

                        # Apply green highlighting to WW Code cells with values
                        if 'WW Code' in col and cell.value and str(cell.value).lower() != 'nan':
                            cell.fill = openpyxl.styles.PatternFill(
                                start_color='90EE90',
                                end_color='90EE90',
                                fill_type='solid'
                            )

            zip_file.writestr(f"processed_{report['name']}", excel_buffer.getvalue())
    return zip_buffer.getvalue()


def auto_download_component(data, filename, mime_type):
    """Create an HTML component that automatically triggers file download"""
    import base64
    b64 = base64.b64encode(data).decode()
    custom_html = f"""
        <html>
            <body>
                <script>
                    function downloadFile(data, filename, mime) {{
                        var bytes = atob(data);
                        var byteArrays = [];

                        for (var offset = 0; offset < bytes.length; offset += 512) {{
                            var slice = bytes.slice(offset, offset + 512);
                            var byteNumbers = new Array(slice.length);

                            for (var i = 0; i < slice.length; i++) {{
                                byteNumbers[i] = slice.charCodeAt(i);
                            }}

                            var byteArray = new Uint8Array(byteNumbers);
                            byteArrays.push(byteArray);
                        }}

                        var blob = new Blob(byteArrays, {{type: mime}});
                        var url = window.URL.createObjectURL(blob);
                        var a = document.createElement('a');

                        a.style.display = 'none';
                        a.href = url;
                        a.download = filename;

                        document.body.appendChild(a);
                        a.click();

                        window.URL.revokeObjectURL(url);
                    }}

                    downloadFile("{b64}", "{filename}", "{mime_type}");
                </script>
            </body>
        </html>
    """
    st.components.v1.html(custom_html, height=0)


def main():
    """Main application function"""
    st.set_page_config(page_title="PARAMOUNT CODE MATCHER", layout="wide")

    # Add custom CSS
    st.markdown("""
        <style>
        .stApp {
            max-width: 1400px;
            margin: 0 auto;
        }
        .title {
            text-align: center;
            font-size: 3.5em;
            font-weight: bold;
            margin-bottom: 1em;
            padding: 20px;
        }
        </style>
    """, unsafe_allow_html=True)

    st.markdown('<div class="title">PARAMOUNT CODE MATCHER</div>', unsafe_allow_html=True)

    # Create three columns for the interface
    col1, col2, col3 = st.columns(3)

    with col1:
        st.markdown("### WCPM Snowflakes Export")
        st.markdown("""
            Expected columns:
            - A: Track Title
            - B: WW Code
            - C: Composers
            - D: Publishers
            - E: BMI Code
            - F: ASCAP Code
            - G: SESAC Code
        """)
        ww_codes_file = st.file_uploader("Upload WW Codes Excel file", type=['xlsx', 'xls'])

    with col2:
        st.markdown("### Tango Export")
        st.markdown("""
            Expected columns:
            - B: WW Code
            - C: Track Title
            - E: Composers
            - I: Publisher
            - W: Publisher Code
        """)
        tango_file = st.file_uploader("Upload Tango Export Excel file", type=['xlsx', 'xls'])

    with col3:
        st.markdown("### Report Files")
        st.markdown("""
            Will add codes starting from AG
        """)
        report_files = st.file_uploader("Upload Report Excel files", type=['xlsx', 'xls'], accept_multiple_files=True)

    # Process files when uploads are ready
    if report_files and (ww_codes_file or tango_file):
        if st.button("Process Files"):
            with st.spinner("Processing files..."):
                processed_reports, matches, update_count = process_files(ww_codes_file, tango_file, report_files)

                if processed_reports:
                    st.success(
                        f"Successfully processed {len(processed_reports)} files! Updated {update_count} entries.")

                    if matches:
                        matches_df = pd.DataFrame(matches)

                        # Handle single file download
                        if len(processed_reports) == 1:
                            excel_buffer = io.BytesIO()
                            with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                                processed_reports[0]['df'].to_excel(writer, index=False)

                            auto_download_component(
                                excel_buffer.getvalue(),
                                f"processed_{processed_reports[0]['name']}",
                                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )

                            st.download_button(
                                "Click here if download doesn't start automatically",
                                data=excel_buffer.getvalue(),
                                file_name=f"processed_{processed_reports[0]['name']}",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
                        # Handle multiple files download
                        else:
                            zip_data = create_download_zip(processed_reports)

                            auto_download_component(
                                zip_data,
                                "processed_reports.zip",
                                "application/zip"
                            )

                            st.download_button(
                                "Click here if download doesn't start automatically",
                                data=zip_data,
                                file_name="processed_reports.zip",
                                mime="application/zip"
                            )

                        # Display matches summary
                        st.markdown("### Matches Found")
                        st.dataframe(matches_df)

                        # Add download button for matches summary
                        excel_buffer = io.BytesIO()
                        with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                            matches_df.to_excel(writer, index=False)

                        st.download_button(
                            "Download Matches Summary",
                            data=excel_buffer.getvalue(),
                            file_name="matches_summary.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                    else:
                        st.warning("No matches found between the files.")


if __name__ == "__main__":
    main()