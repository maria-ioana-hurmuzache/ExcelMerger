import streamlit as st
import pandas as pd
import io
import zipfile

st.set_page_config(page_title="Excel Merger", layout="centered")
st.title("Excel Item Price Merger")


main_files = st.file_uploader("Încarcă aici fișierele care trebuie completate", type=["xlsx"], accept_multiple_files=True)

# Upload Price File (File B)
price_file = st.file_uploader("Încarcă aici fișierul cu prețuri", type=["xlsx"])

if main_files and price_file:
    st.success(f"{len(main_files)} fișiere încărcate pentru completare.")
    price_file.seek(0)
    # Optional: sheet selection
    price_sheets = pd.ExcelFile(price_file, engine='openpyxl').sheet_names
    selected_sheets = price_sheets[:2]

    if st.button("Generează fișierele completate"):
        try:
            df_prices_combined = pd.DataFrame()
            for sheet in selected_sheets:
                df = pd.read_excel(price_file, sheet_name=sheet, skiprows=7, engine='openpyxl')
                try:
                    exchange_rate_df = pd.read_excel(price_file, sheet_name=sheet, header=None, nrows=3, usecols="C", engine='openpyxl')
                    exchange_rate = pd.to_numeric(exchange_rate_df.iloc[1, 0], errors='coerce')

                    df_clean = df[['Unnamed: 1', 'EUR fara TVA']].copy()
                    df_clean.columns = ['Item Code', 'EUR Pret']
                    df_clean = df_clean.dropna(subset=['Item Code', 'EUR Pret'])

                    df_clean['Pret'] = df_clean['EUR Pret'] * exchange_rate
                    df_clean['Pret'] = df_clean['Pret'].round(2)

                    df_clean = df_clean[['Item Code', 'Pret']]
                    df_prices_combined = pd.concat([df_prices_combined, df_clean], ignore_index=True)
                except Exception as e:
                    st.warning(f"Eroare la procesarea foii '{sheet}': {e}")
                    

            if df_prices_combined.empty:
                st.error("Nu s-au găsit date în fișierul cu prețuri.")
            else:
                zip_buffer = io.BytesIO()
                # Merge
                with zipfile.ZipFile(zip_buffer, mode='w', compression=zipfile.ZIP_DEFLATED) as zip_file:
                    for file in main_files:
                        file.seek(0)
                        df_main = pd.read_excel(file, engine='openpyxl')
                        
                        # --- MODIFICARE PENTRU COLOANE DINAMICE ---
                        cols = df_main.columns.tolist()
                        # Găsim numele real al coloanei 'Item Code' (case insensitive)
                        actual_item_col = next((c for c in cols if c.lower().strip() == 'item code'), 'Item Code')
                        # Găsim numele real al coloanei 'Unit Price' (case insensitive)
                        actual_unit_price_col = next((c for c in cols if c.lower().strip() == 'unit price'), 'Unit Price')
                        # Găsim numele real al coloanei care începe cu 'Quantity' (case insensitive)
                        actual_quantity_col = next((c for c in cols if c.lower().strip().startswith('quantity')), 'Quantity')
                        actual_total_price_col = next((c for c in cols if c.lower().strip() == 'total price'), 'Total Price')
                        # ------------------------------------------                        
                        
                        df_result = pd.merge(df_main, df_prices_combined, on=actual_item_col, how='left')

                        # Fill existing price fields
                        if actual_unit_price_col in df_result.columns:
                            df_result[actual_unit_price_col] = df_result['Pret']

                        if actual_quantity_col in df_result.columns:
                            df_result[actual_total_price_col] = df_result[actual_unit_price_col] * df_result[actual_quantity_col]
                        # Drop helper column
                        df_result.drop(columns=['Pret'], inplace=True)

                        # Prepare file to download
                        output = io.BytesIO()
                        df_result.to_excel(output, index=False, engine='openpyxl')
                        output.seek(0)
                        new_filename = file.name.replace('.xlsx', '_merged.xlsx')
                        zip_file.writestr(new_filename, output.read())
                zip_buffer.seek(0)
                st.download_button(label="Descarcă arhiva ZIP cu fișierele completate", data=zip_buffer, file_name="rezultate.zip", mime="application/zip")
        except Exception as e:
            st.error(f"Eroare: {e}")
