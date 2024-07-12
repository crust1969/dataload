import streamlit as st
import pandas as pd
import os

# Set page config
st.set_page_config(page_title='Excel File Processor', layout='wide')

# Sidebar for file upload
st.sidebar.header('Upload Excel File')
uploaded_file = st.sidebar.file_uploader("Choose a file", type=['xlsx'])

# Sidebar for row extension
st.sidebar.header('Extend Rows')
extend_rows = st.sidebar.number_input('Enter number of rows to extend', min_value=0, value=0)

# Sidebar for saving the file
st.sidebar.header('Save File')
output_file = st.sidebar.text_input('Enter the output file name', value='output.xlsx')
save_directory = st.sidebar.text_input('Enter the directory to save the file', value=os.getcwd())
save_button = st.sidebar.button('Save Excel File')

if uploaded_file is not None:
    # Load the Excel file into a dataframe
    df = pd.read_excel(uploaded_file, header=0)

    # Extend the dataframe with empty rows if needed
    if extend_rows > 0:
        extra_rows = pd.DataFrame([[''] * len(df.columns)] * extend_rows, columns=df.columns)
        df = pd.concat([df, extra_rows], ignore_index=True)

    # Display the dataframe
    st.write('## Excel File Content', df)

    if save_button:
        # Save the modified dataframe back to an Excel file
        save_path = os.path.join(save_directory, output_file)
        df.to_excel(save_path, index=False)
        st.success(f'File saved as {save_path}')
else:
    st.write('## Waiting for Excel file upload...')
