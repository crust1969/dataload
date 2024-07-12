import streamlit as st
import pandas as pd
import io

# Set page config
st.set_page_config(page_title='Excel File Processor', layout='wide')

# Function to convert input to the correct data type
def convert_input(column_name, input_value, data_type):
    if pd.api.types.is_integer_dtype(data_type):
        return pd.to_numeric(input_value, errors='coerce', downcast='integer')
    elif pd.api.types.is_float_dtype(data_type):
        return pd.to_numeric(input_value, errors='coerce')
    elif pd.api.types.is_datetime64_any_dtype(data_type):
        return pd.to_datetime(input_value, errors='coerce')
    else:
        return input_value

# Sidebar for file upload
st.sidebar.header('Upload Excel File')
uploaded_file = st.sidebar.file_uploader("Choose a file", type=['xlsx'])

# Sidebar for row extension
st.sidebar.header('Extend Rows')
extend_rows = st.sidebar.number_input('Enter number of rows to extend', min_value=0, value=0)

# Sidebar for output file name
st.sidebar.header('Output File Name')
output_file = st.sidebar.text_input('Enter the output file name', value='output.xlsx')

# Initialize an empty list to store new row data
new_rows_data = []

if uploaded_file is not None:
    # Load the Excel file into a dataframe
    df = pd.read_excel(uploaded_file, header=0)

    # Display the original dataframe
    st.write('## Original Excel File Content', df)

    # If the user wants to extend rows, display input fields for each column
    if extend_rows > 0:
        for _ in range(extend_rows):
            new_row = {}
            for col in df.columns:
                col_type = df[col].dtype
                default_value = df[col].iloc[-1] if not df[col].empty else None
                new_value = st.sidebar.text_input(f'New value for {col}', value=default_value, key=f'{col}_{len(df) + len(new_rows_data)}')
                new_row[col] = convert_input(col, new_value, col_type)
            new_rows_data.append(new_row)

    # Button to save the updated dataframe to an Excel file
    save_button = st.sidebar.button('Save Excel File')

    if save_button:
        # Convert the list of new rows into a DataFrame
        new_rows_df = pd.DataFrame(new_rows_data, columns=df.columns)

        # Append the new rows to the existing DataFrame
        updated_df = pd.concat([df, new_rows_df], ignore_index=True)

        # Display the updated dataframe
        st.write('## Updated Excel File Content', updated_df)

        # Save the updated dataframe to an in-memory buffer
        buffer = io.BytesIO()
        updated_df.to_excel(buffer, index=False)
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
