import pandas as pd
import streamlit as st
from io import BytesIO

# Function to clean stock sheet
def clean_stock_sheet(xls, sheet_name):
    df = pd.read_excel(xls, sheet_name=sheet_name, skiprows=2)
    df = df.rename(columns={'Micro Category:': 'Item Code', 'Unnamed: 13': 'Balance Qty Pieces'})
    return df[['Item Code', 'Balance Qty Pieces']].dropna()

# Streamlit app
def main():
    st.title("The Grocer Stock Data Processor")

    # File uploader
    uploaded_file = st.file_uploader("Upload your Excel file", type="xlsx")

    if uploaded_file is not None:
        # Read the uploaded Excel file
        xls = pd.ExcelFile(uploaded_file)

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

            # Map stock to binary (1 if stock >= 1, else 0)
            stock_data['STOCK'] = (stock_data['Balance Qty Pieces'] >= 1).astype(int)

            # Load and clean the main sheet
            main_sheet = pd.read_excel(xls, sheet_name='دليل الاصناف EN', skiprows=3)
            main_sheet = main_sheet.rename(columns={
                'Micro Category :': 'Item Code',
                'Unnamed: 2': 'BarCode',
                'Unnamed: 5': 'Item Name',
                'Unnamed: 9': 'Retail Price'
            })
            main_sheet = main_sheet[['Item Code', 'BarCode', 'Item Name', 'Retail Price']]

            # Load the force instock sheet
            force_instock = pd.read_excel(xls, sheet_name='force instock')
            force_instock['Item No'] = force_instock['Item No'].astype(str).str.strip()

            # Create records for forced instock items
            forced_stock_data = []
            for _, row in force_instock.iterrows():
                matching_item = main_sheet[main_sheet['Item Code'].astype(str).str.strip() == row['Item No']]
                if not matching_item.empty:
                    forced_stock_data.append({
                        'Item Code': matching_item.iloc[0]['Item Code'],
                        'Balance Qty Pieces': 1,
                        'Store': row['Store'],
                        'STOCK': 1
                    })

            # Convert forced stock data to DataFrame
            forced_stock_df = pd.DataFrame(forced_stock_data)

            # Combine regular stock data with forced stock data
            if forced_stock_data:
                stock_data = pd.concat([stock_data, forced_stock_df], ignore_index=True)

            # Drop duplicates keeping the forced stock entries
            stock_data = stock_data.drop_duplicates(subset=['Store', 'Item Code'], keep='last')

            # Standardize Item Code to ensure consistent formatting
            stock_data['Item Code'] = stock_data['Item Code'].astype(str).str.strip()
            main_sheet['Item Code'] = main_sheet['Item Code'].astype(str).str.strip()

            # Merge main sheet with stock data
            final_data = stock_data.merge(main_sheet, on='Item Code', how='left')

            # Rearrange columns
            final_data = final_data[['Store', 'Item Code', 'BarCode', 'Item Name', 'Retail Price', 'STOCK']]

            # Map the Store names to standardized values
            store_mapping = {
                'معادي': 'Maadi',
                'MDI': 'Maadi',
                'زمالك': 'Zamalek',
                'ZMK': 'Zamalek',
                'جاردن': 'Garden 8',
                'GRD': 'Garden 8'
            }

            final_data['Store'] = final_data['Store'].replace(store_mapping)

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
