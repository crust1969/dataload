import streamlit as st
import pandas as pd

# Set page config
st.set_page_config(page_title='Excel File Processor', layout='wide')

# Sidebar for file upload
st.sidebar.header('Upload Excel File')
uploaded_file = st.sidebar.file_uploader("Choose a file", type=['xlsx'])

# Sidebar for column extension
st.sidebar.header('Extend Columns')
extend_columns = st.sidebar.number_input('Enter number of columns to extend', min_value=0, value=0)

if uploaded_file is not None:
    # Load the Excel file into a dataframe
    df = pd.read_excel(uploaded_file, header=0)

    # Extend the dataframe with empty columns if needed
    for i in range(extend_columns):
        df[f'Extended Column {i+1}'] = ''

    # Display the dataframe
    st.write('## Excel File Content', df)

    # File browser to save the file back
    st.header('Save File')
    output_file = st.text_input('Enter the output file name', value='output.xlsx')
    save_button = st.button('Save Excel File')

    if save_button:
        # Save the modified dataframe back to an Excel file
        df.to_excel(output_file, index=False)
        st.success(f'File saved as {output_file}')
else:
    st.write('## Waiting for Excel file upload...')
