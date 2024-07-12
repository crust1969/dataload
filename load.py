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

if uploaded_file is not None:
    # Load the Excel file into a dataframe
    df = pd.read_excel(uploaded_file, header=0)

    # Display the original dataframe
    st.write('## Original Excel File Content', df)

    # If the user wants to extend rows, display input fields for each column
    if extend_rows > 0:
        # Create a dictionary to hold the new row data
        new_row_data = {col: None for col in df.columns}
        for col in df.columns:
            # Determine the data type of the column
            col_type = df[col].dtype
            # Create input fields based on the data type
            if pd.api.types.is_numeric_dtype(col_type):
                new_row_data[col] = st.sidebar.number_input(f'New value for {col}', key=f'{col}_{len(df)}')
            elif pd.api.types.is_datetime64_any_dtype(col_type):
                new_row_data[col] = st.sidebar.date_input(f'New value for {col}', key=f'{col}_{len(df)}')
            else:
                new_row_data[col] = st.sidebar.text_input(f'New value for {col}', key=f'{col}_{len(df)}')

        # Button to add new row to the dataframe
        add_row_button = st.sidebar.button('Add New Row')

        if add_row_button:
            # Convert the new row data to a DataFrame with the same data types as the original
            new_row_df = pd.DataFrame([new_row_data])
            for col in df.columns:
                new_row_df[col] = new_row_df[col].astype(df[col].dtype)

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

        # Display the updated dataframe before download
        st.write('## Updated Excel File Content Before Download', df)

        # Provide a download link
        st.download_button(
            label='Download Updated Excel File',
            data=buffer,
            file_name=output_file,
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
else:
    st.write('## Waiting for Excel file upload...')
