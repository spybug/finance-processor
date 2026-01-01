import pandas as pd
import io
import streamlit as st


def process_csv(file_path='ExportData.csv'):
    """Process the downloaded CSV file."""
    st.title("Finance Processor")
    uploaded_file = st.file_uploader("Choose a CSV file", type="csv")

    if uploaded_file is not None:
        progress_text = "Processing file..."
        my_bar = st.progress(0, text=progress_text)

        try:
            df = pd.read_csv(uploaded_file)
        except Exception as e:
            st.error(f"Error reading file: {e}")
            return

        # Filter out unwanted rows
        df = df[~df['Status'].str.contains('scheduled', na=False)]
        df = df[~df['Original Description'].str.contains('Credit Card Payment', na=False)]
        df = df[~df['Original Description'].str.contains('External Withdrawal.*(?:ALLY|GRDOBK|CAPITAL ONE)', na=False)]
        df = df[~df['Original Description'].str.contains('Dividend.*Interest', na=False)]
        df = df[df['Category-Subcategory'] != 'Credit Card Payments']
        df = df[df['Account Name'] != 'BECU - Loan - Auto Loan']

        my_bar.progress(30, text="Cleaning data...")

        # Select and reorder columns
        df = df[['Date', 'Original Description', 'Amount', 'Category-Subcategory', 'Account Name']]
        df = df.rename(columns={'Category-Subcategory': 'Category'})

        # Clean and convert 'Amount' column to numeric
        df['Amount'] = df['Amount'].replace({'$': '', ',': ''}, regex=True).astype(float)

        # Convert 'Date' column to datetime
        df['Date'] = pd.to_datetime(df['Date'])

        # Sort entries by date ascending
        df = df.sort_values(by='Date', ascending=True)

        my_bar.progress(60, text="Separating income and expenses...")

        # Separate income and expenses
        income_df = df[df['Amount'] > 0].copy()
        expenses_df = df[df['Amount'] < 0].copy()

        # Invert expense amounts
        expenses_df['Amount'] = expenses_df['Amount'] * -1

        # Format date to remove time
        income_df['Date'] = income_df['Date'].dt.date
        expenses_df['Date'] = expenses_df['Date'].dt.date

        # Save to Excel
        my_bar.progress(80, text="Saving to Excel...")
        month_year = df['Date'].iloc[0].strftime('%B-%Y')
        output_filename = f'{month_year}.xlsx'
        
        def style_df(df):
            return df.style.set_properties(**{
                'background-color': '#333333',
                'color': 'white',
                'border': '1px solid white'
            }).set_table_styles([{
                'selector': 'th',
                'props': [('background-color', '#333333'), ('color', 'white'), ('border', '1px solid white')]
            }])

        output_buffer = io.BytesIO()
        with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
            style_df(income_df).to_excel(writer, sheet_name='Income', index=False)
            style_df(expenses_df).to_excel(writer, sheet_name='Expenses', index=False)

            # Get workbook and sheets
            income_sheet = writer.sheets['Income']
            expenses_sheet = writer.sheets['Expenses']

            for sheet in [income_sheet, expenses_sheet]:
                for column in sheet.columns:
                    max_length = 0
                    column_letter = column[0].column_letter
                    
                    # Find the max length in a column
                    for cell in column:
                        try:
                            if len(str(cell.value)) > max_length:
                                max_length = len(str(cell.value))
                        except:
                            pass
                    
                    # Adjust the column width with a little extra padding
                    adjusted_width = max_length + 2
                    sheet.column_dimensions[column_letter].width = adjusted_width
        
        my_bar.progress(100, text="Complete!")
        st.success(f'Successfully processed data for {month_year}')
        
        st.download_button(
            label="Download Excel file",
            data=output_buffer.getvalue(),
            file_name=output_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

if __name__ == "__main__":
    process_csv()