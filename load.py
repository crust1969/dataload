import streamlit as st
import pandas as pd
import io

# Set page config
st.set_page_config(page_title='Excel File Processor', layout='wide')

# Sidebar for file upload
st.sidebar.header('Upload Excel File')
uploaded_file = st.sidebar.file_uploader("Choose a file", type=['xlsx'])

# Sidebar for row extension
st.sidebar.header('Extend Rows')
extend_rows = st.sidebar.number_input('Enter number of rows to extend', min_value=0, value=0)

# Sidebar for output file name
st.sidebar.header('Output File Name')
output_file = st.sidebar.text_input('Enter the output file name', value='output.xlsx')

# Initialize an empty DataFrame for new rows
new_rows_df = pd.DataFrame()

if uploaded_file is not None:
    # Load the Excel file into a dataframe
    df = pd.read_excel(uploaded_file, header=0)

    # Display the original dataframe
    st.write('## Original Excel File Content', df)

    # If the user wants to extend rows, display input fields for each column
    if extend_rows > 0:
        new_row_data = {}
        for col in df.columns:
            new_row_data[col] = st.sidebar.text_input(f'New value for {col}', key=f'{col}_{len(df) + len(new_rows_df)}')

        # Button to add new row to the dataframe
        add_row_button = st.sidebar.button('Add New Row')

        if add_row_button:
            # Convert the new row data to a DataFrame
            new_row_df = pd.DataFrame([new_row_data])

            # Ensure the data types match the original DataFrame
            for col in df.columns:
                if pd.api.types.is_numeric_dtype(df[col]):
                    new_row_df[col] = pd.to_numeric(new_row_df[col], errors='coerce')

            # Append the new row to the existing DataFrame
            df = pd.concat([df, new_row_df], ignore_index=True)

            # Display the updated dataframe
            st.write('## Updated Excel File Content', df)

    # Button to save the updated dataframe to an Excel file
    save_button = st.sidebar.button('Save Excel File')

    if save_button:
        # Save the modified dataframe to an in-memory buffer
        buffer = io.BytesIO()
        df.to_excel(buffer, index=False)
        buffer.seek(0)

        # Provide a download link
        st.download_button(
            label='Download Updated Excel File',
            data=buffer,
            file_name=output_file,
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
else:
    st.write('## Waiting for Excel file upload...')
