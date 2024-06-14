import streamlit as st
import pandas as pd
from docx import Document
from io import BytesIO

def replace_placeholders(doc, data):
    """
    Replace placeholders in the Word document with data from the DataFrame.
    
    Parameters:
    doc (Document): The Word document object.
    data (dict): A dictionary with placeholders as keys and replacement data as values.
    """
    for p in doc.paragraphs:
        for key, value in data.items():
            if key in p.text:
                inline = p.runs
                for i in range(len(inline)):
                    if key in inline[i].text:
                        text = inline[i].text.replace(key, str(value))
                        inline[i].text = text
    return doc

def preprocess_excel(df):
    """
    Preprocess the DataFrame by unmerging cells and forward-filling values.
    
    Parameters:
    df (DataFrame): The original DataFrame read from the Excel file.
    
    Returns:
    DataFrame: The cleaned DataFrame.
    """
    # Forward fill values for merged cells
    df.fillna(method='ffill', axis=0, inplace=True)
    return df

def inspect_and_map_data(df):
    """
    Inspect the DataFrame to determine the correct data mappings for placeholders.
    
    Parameters:
    df (DataFrame): The preprocessed DataFrame.
    
    Returns:
    dict: A dictionary with placeholders as keys and corresponding DataFrame values as values.
    """
    # Safe access to DataFrame values with fallback
    def safe_get_value(df, row, col):
        try:
            return df.iloc[row, col]
        except IndexError:
            return "N/A"

    # Define the mapping from Excel to Word placeholders based on inspection
    data_mapping = {
        "{{Nome do município}}": safe_get_value(df, 0, 1),  # Adjust based on actual data
        "{{População residente}}": safe_get_value(df, 6, 1),  # Adjust based on actual data
        "{{Área da unidade territorial}}": safe_get_value(df, 7, 1),  # Adjust based on actual data
        "{{Densidade demográfica}}": safe_get_value(df, 8, 1),  # Adjust based on actual data
        "{{Área total}}": safe_get_value(df, 12, 7),  # Adjust based on actual data
        "{{Plantio em nível}}": safe_get_value(df, 13, 8),  # Adjust based on actual data
        "{{Rotação de culturas}}": safe_get_value(df, 14, 9),  # Adjust based on actual data
        "{{Pousio ou descanso}}": safe_get_value(df, 15, 10),  # Adjust based on actual data
        "{{Proteção de encostas}}": safe_get_value(df, 16, 11),  # Adjust based on actual data
        "{{Recuperação de mata ciliar}}": safe_get_value(df, 17, 12),  # Adjust based on actual data
        "{{Reflorestamento de nascentes}}": safe_get_value(df, 18, 13),  # Adjust based on actual data
        "{{Estabilização de voçorocas}}": safe_get_value(df, 19, 14),  # Adjust based on actual data
        "{{Manejo florestal}}": safe_get_value(df, 20, 15),  # Adjust based on actual data
        "{{Outras}}": safe_get_value(df, 21, 16),  # Adjust based on actual data
        "{{PIB}}": safe_get_value(df, 17, 1),  # Adjust based on actual data
        "{{Percentual da agricultura}}": safe_get_value(df, 18, 1),  # Adjust based on actual data
        "{{Valor Adicionado Bruto Agropecuária}}": safe_get_value(df, 25, 1),  # Adjust based on actual data
        "{{Valor Adicionado Bruto Indústria}}": safe_get_value(df, 26, 1),  # Adjust based on actual data
        "{{Valor Adicionado Bruto Serviços}}": safe_get_value(df, 27, 1),  # Adjust based on actual data
        "{{Valor Adicionado Bruto Administração Pública}}": safe_get_value(df, 28, 1)  # Adjust based on actual data
    }
    return data_mapping

# Streamlit app layout
st.title('Excel to Word Document Generator with Template')

# File uploaders for Excel file and Word template
uploaded_excel = st.file_uploader("Choose an Excel file", type="xlsx")
uploaded_word = st.file_uploader("Choose a Word template", type="docx")

if uploaded_excel is not None and uploaded_word is not None:
    # Load the Excel file with proper headers and skip metadata rows
    df = pd.read_excel(uploaded_excel, skiprows=7, header=1, engine='openpyxl')
    
    # Preprocess the DataFrame to handle merged cells
    df = preprocess_excel(df)

    # Display the DataFrame columns and the first few rows for debugging
    st.write("Columns in the Excel file:")
    st.write(df.columns)
    st.write("First few rows in the Excel file:")
    st.write(df.head(20))

    # Inspect the DataFrame and map the data
    data_mapping = inspect_and_map_data(df)

    # Read the Word document
    doc = Document(uploaded_word)

    # Replace placeholders in the Word document
    doc = replace_placeholders(doc, data_mapping)

    # Save the updated document to a BytesIO object
    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)

    st.success("Word document updated successfully!")

    # Download button to download the updated Word document
    st.download_button(
        label="Download Updated Word Document",
        data=buffer,
        file_name="updated_document.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
