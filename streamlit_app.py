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

# Streamlit app layout
st.title('Excel to Word Document Generator with Template')

# File uploaders for Excel file and Word template
uploaded_excel = st.file_uploader("Choose an Excel file", type="xlsx")
uploaded_word = st.file_uploader("Choose a Word template", type="docx")

if uploaded_excel is not None and uploaded_word is not None:
    # Load the Excel file with proper headers and skip metadata rows
    df = pd.read_excel(uploaded_excel, skiprows=7, header=1, engine='openpyxl')

    # Display the DataFrame columns and the first few rows for debugging
    st.write("Columns in the Excel file:")
    st.write(df.columns)
    st.write("First few rows in the Excel file:")
    st.write(df.head(20))

    # Read the Word document
    doc = Document(uploaded_word)

    # Define the mapping from Excel to Word placeholders based on inspection
    data_mapping = {
        "{{Nome do município}}": df.iloc[0, df.columns.get_loc("Unnamed: 1")],
        "{{População residente}}": df.iloc[6, df.columns.get_loc("Unnamed: 1")],  # Adjusted based on actual data
        "{{Área da unidade territorial}}": df.iloc[7, df.columns.get_loc("Unnamed: 1")],  # Adjusted based on actual data
        "{{Densidade demográfica}}": df.iloc[8, df.columns.get_loc("Unnamed: 1")],  # Adjusted based on actual data
        "{{Área total}}": df.iloc[12, df.columns.get_loc("de 50 a 500 ha")],
        "{{Plantio em nível}}": df.iloc[13, df.columns.get_loc(422)],
        "{{Rotação de culturas}}": df.iloc[14, df.columns.get_loc(12)],
        "{{Pousio ou descanso}}": df.iloc[15, df.columns.get_loc(51)],
        "{{Proteção de encostas}}": df.iloc[16, df.columns.get_loc(58)],
        "{{Recuperação de mata ciliar}}": df.iloc[17, df.columns.get_loc(4)],
        "{{Reflorestamento de nascentes}}": df.iloc[18, df.columns.get_loc(1)],
        "{{Estabilização de voçorocas}}": df.iloc[19, df.columns.get_loc(0)],
        "{{Manejo florestal}}": df.iloc[20, df.columns.get_loc("1.1")],
        "{{Outras}}": df.iloc[21, df.columns.get_loc(8)],
        "{{PIB}}": df.iloc[17, df.columns.get_loc("Unnamed: 1")],  # Adjusted based on actual data
        "{{Percentual da agricultura}}": df.iloc[18, df.columns.get_loc("Unnamed: 1")],  # Adjusted based on actual data
        "{{Valor Adicionado Bruto Agropecuária}}": df.iloc[25, df.columns.get_loc("Unnamed: 1")],  # Adjusted based on actual data
        "{{Valor Adicionado Bruto Indústria}}": df.iloc[26, df.columns.get_loc("Unnamed: 1")],  # Adjusted based on actual data
        "{{Valor Adicionado Bruto Serviços}}": df.iloc[27, df.columns.get_loc("Unnamed: 1")],  # Adjusted based on actual data
        "{{Valor Adicionado Bruto Administração Pública}}": df.iloc[28, df.columns.get_loc("Unnamed: 1")]  # Adjusted based on actual data
    }

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
