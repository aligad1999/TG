import pandas as pd
import streamlit as st
from io import BytesIO

# Required sheet names
REQUIRED_SHEETS = ['دليل الاصناف EN', 'زمالك', 'معادي', 'جاردن', 'force instock']

# Function to clean stock sheet
def clean_stock_sheet(xls, sheet_name):
    df = pd.read_excel(xls, sheet_name=sheet_name, skiprows=2)
    df = df.rename(columns={'Micro Category:': 'Item Code', 'Unnamed: 13': 'Balance Qty Pieces'})
    return df[['Item Code', 'Balance Qty Pieces']].dropna()

# Streamlit app
def main():
    st.image("logo.png", width=200)  # Replace "logo.png" with the path to your logo file
    st.title("The Grocer Stock Data Processor")

    # File uploader
    uploaded_file = st.file_uploader("Upload your Excel file", type="xlsx")

    if uploaded_file is not None:
        # Read the uploaded Excel file
        try:
            xls = pd.ExcelFile(uploaded_file)
        except Exception as e:
            st.error(f"Error reading the file: {e}")
            return

        # Validate required sheets
        missing_sheets = [sheet for sheet in REQUIRED_SHEETS if sheet not in xls.sheet_names]
        if missing_sheets:
            st.error(f"The uploaded file is missing the following required sheets: {', '.join(missing_sheets)}. Please add them and reupload the file.")
            return

        # Process stock data
        try:
            zamalek_data = clean_stock_sheet(xls, 'زمالك')
            zamalek_data['Store'] = 'زمالك'
            maadi_data = clean_stock_sheet(xls, 'معادي')
            maadi_data['Store'] = 'معادي'
            garden_data = clean_stock_sheet(xls, 'جاردن')
            garden_data['Store'] = 'جاردن'

            # Combine stock data
            stock_data = pd.concat([zamalek_data, maadi_data, garden_data], ignore_index=True)

            # Load the main sheet
            main_sheet = pd.read_excel(xls, sheet_name='دليل الاصناف EN', skiprows=3)
            main_sheet = main_sheet.rename(columns={
                'Micro Category :': 'Item Code',
                'Unnamed: 1': 'BarCode',
                'Unnamed: 3': 'Item Name',
                'Unnamed: 4': 'Retail Price',
                'Unnamed: 5': 'STOCK',
                'Unnamed: 6': 'Discounted Price'
            })

            # Process BarCode column by removing 'plus'
            main_sheet['BarCode'] = main_sheet['BarCode'].apply(lambda x: str(x).replace('plus', '').strip() if pd.notnull(x) else x)

            # Standardize Item Code
            stock_data['Item Code'] = stock_data['Item Code'].astype(str).str.strip()
            main_sheet['Item Code'] = main_sheet['Item Code'].astype(str).str.strip()

            # Map stock (1 if stock >= 1, else 0)
            main_sheet['STOCK'] = (main_sheet['STOCK'] >= 1).astype(int)

            # Replace empty Discounted Price with Retail Price
            main_sheet['Discounted Price'] = main_sheet['Discounted Price'].fillna(main_sheet['Retail Price'])

            # Load the force instock sheet
            force_instock = pd.read_excel(xls, sheet_name='force instock')
            force_instock['Item No'] = force_instock['Item No'].astype(str).str.strip()

            # Merge stock data with main sheet
            final_data = stock_data.merge(main_sheet, on='Item Code', how='left')

            # Map store names
            store_mapping = {
                'معادي': 'Maadi',
                'MDI': 'Maadi',
                'زمالك': 'Zamalek',
                'ZMK': 'Zamalek',
                'جاردن': 'Garden 8',
                'GRD': 'Garden 8'
            }
            final_data['Store'] = final_data['Store'].replace(store_mapping)

            # Select and rename columns to match exact mapping
            final_data = final_data[[
                'Store', 
                'BarCode', 
                'Item Name', 
                'Retail Price', 
                'Discounted Price', 
                'STOCK'
            ]]

            # Prepare data for download
            output = BytesIO()
            final_data.to_excel(output, index=False, engine='openpyxl')
            output.seek(0)

            # Provide download link
            st.download_button(
                label="Download Processed File",
                data=output,
                file_name="final_stock_data.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"Error processing the file: {e}")

if __name__ == "__main__":
    main()
