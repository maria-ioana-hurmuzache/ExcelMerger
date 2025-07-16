import streamlit as st
import pandas as pd
import io
import zipfile

st.set_page_config(page_title="Excel Merger", layout="centered")
st.title("Excel Item Price Merger")


# Upload Main File (File A)
main_files = st.file_uploader("Upload Main File (Item List)", type=["xlsx"], accept_multiple_files=True)

# Upload Price File (File B)
price_file = st.file_uploader("Upload Price File (with Item Code + Price)", type=["xlsx"])

if main_files and price_file:
    st.success(f"{len(main_files)} file(s) loaded for merging.")

    # Optional: sheet selection
    price_sheets = pd.ExcelFile(price_file).sheet_names
    selected_sheets = st.multiselect("Select Sheets to Extract Prices From", price_sheets, default=price_sheets[:2])

    if st.button("Merge and Generate File"):
        try:
            # Load Main File
            # df_main = pd.read_excel(main_files)

            # Combine prices from selected sheets
            df_prices_combined = pd.DataFrame()
            for sheet in selected_sheets:
                df = pd.read_excel(price_file, sheet_name=sheet, skiprows=7)
                try:
                    df_clean = df[['Unnamed: 1', 'RON cu TVA']].copy()
                    df_clean.columns = ['Item Code', 'Pret']
                    df_clean = df_clean.dropna(subset=['Item Code', 'Pret'])
                    df_prices_combined = pd.concat([df_prices_combined, df_clean], ignore_index=True)
                except:
                    pass

            if df_prices_combined.empty:
                st.error("No valid data found in selected sheets.")
            else:
                zip_buffer = io.BytesIO()
                # Merge
                with zipfile.ZipFile(zip_buffer, mode='w', compression=zipfile.ZIP_DEFLATED) as zip_file:
                    for file in main_files:
                        df_main = pd.read_excel(file)
                        df_result = pd.merge(df_main, df_prices_combined, on='Item Code', how='left')

                        # Fill existing price fields
                        if 'ITEM Price' in df_result.columns:
                            df_result['ITEM Price'] = df_result['Pret']

                        if 'Quantity in Bucket' in df_result.columns:
                            df_result['Total Price'] = df_result['ITEM Price'] * df_result['Quantity in Bucket']

                        # Drop helper column
                        df_result.drop(columns=['Pret'], inplace=True)

                        # Prepare file to download
                        output = io.BytesIO()
                        df_result.to_excel(output, index=False, engine='openpyxl')
                        output.seek(0)
                        new_filename = file.name.replace('.xlsx', '_merged.xlsx')
                        zip_file.writestr(new_filename, output.read())
                zip_buffer.seek(0)
                st.download_button(label="Download All files as a ZIP", data=zip_buffer, file_name="merged_files.zip", mime="application/zip")
        except Exception as e:
            st.error(f"Error: {e}")
