!pip install --upgrade pip
import streamlit as st
import pandas as pd
import os
from tkinter import filedialog
from tkinter import Tk

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

    # If the user wants to extend rows, display input fields for each column
    if extend_rows > 0:
        for i in range(extend_rows):
            st.sidebar.subheader(f'Row {len(df) + i + 1}')
            new_row = []
            for col in df.columns:
                new_value = st.sidebar.text_input(f'{col} (Row {len(df) + i + 1})', key=f'{col}_{i}')
                new_row.append(new_value)
            new_rows_data.append(new_row)

    # Display the dataframe
    st.write('## Excel File Content', df)

    # Button to add new rows to the dataframe
    add_rows_button = st.sidebar.button('Add Rows to DataFrame')

    if add_rows_button:
        # Convert the list of new rows into a DataFrame and append to the existing DataFrame
        new_rows_df = pd.DataFrame(new_rows_data, columns=df.columns)
        df = pd.concat([df, new_rows_df], ignore_index=True)
        st.write('## Updated Excel File Content', df)

    # Sidebar for saving the file
    st.sidebar.header('Save File')
    output_file = st.sidebar.text_input('Enter the output file name', value='output.xlsx')
    save_button = st.sidebar.button('Save Excel File')

    if save_button:
        # Use tkinter filedialog to choose directory
        root = Tk()
        root.withdraw()  # Hide the main window
        save_directory = filedialog.askdirectory()
        root.destroy()

        # Save the modified dataframe back to an Excel file
        save_path = os.path.join(save_directory, output_file)
        df.to_excel(save_path, index=False)
        st.success(f'File saved as {save_path}')
else:
    st.write('## Waiting for Excel file upload...')
