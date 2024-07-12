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

# Initialize an empty list to store new row data
new_rows_data = []

if uploaded_file is not None:
    # Load the Excel file into a dataframe
    df = pd.read_excel(uploaded_file, header=0)

    # Display the original dataframe
    st.write('## Original Excel File Content', df)

    # If the user wants to extend rows, display input fields for each column
    if extend_rows > 0:
        with st.sidebar.form(key='row_extension_form'):
            for i in range(extend_rows):
                st.subheader(f'Row {len(df) + i + 1}')
                new_row = []
                for col in df.columns:
                    new_value = st.text_input(f'{col} (Row {len(df) + i + 1})', key=f'{col}_{i}')
                    new_row.append(new_value)
                new_rows_data.append(new_row)
            submit_button = st.form_submit_button(label='Add Rows to DataFrame')

    if submit_button:
        # Convert the list of new rows into a DataFrame and append to the existing DataFrame
        new_rows_df = pd.DataFrame(new_rows_data, columns=df.columns)
        df = pd.concat([df, new_rows_df], ignore_index=True)

        # Display the updated dataframe
        st.write('## Updated Excel File Content', df)

        # Save the modified dataframe to an in-memory buffer
        buffer = io.BytesIO()
        df.to_excel(buffer, index=False)
        buffer.seek(0)

        # Provide a download link
        st.download_button(
            label='Download Updated Excel File',
            data=buffer,
            file_name='updated_file.xlsx',
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
else:
    st.write('## Waiting for Excel file upload...')
