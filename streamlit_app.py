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
    # Read the Excel file into a DataFrame
    df = pd.read_excel(uploaded_excel, skiprows=4, engine='openpyxl')

    # Read the Word document
    doc = Document(uploaded_word)

    # Define the mapping from Excel to Word placeholders
    data_mapping = {
        "{{Nome do município}}": df.loc[0, "Unnamed: 3"],
        "{{Área total}}": df.loc[0, "Unnamed: 7"],
        "{{Plantio em nível}}": df.loc[0, "Unnamed: 9"],
        "{{Rotação de culturas}}": df.loc[0, "Unnamed: 10"],
        "{{Pousio ou descanso}}": df.loc[0, "Unnamed: 11"],
        "{{Proteção de encostas}}": df.loc[0, "Unnamed: 12"],
        "{{Recuperação de mata ciliar}}": df.loc[0, "Unnamed: 13"],
        "{{Reflorestamento de nascentes}}": df.loc[0, "Unnamed: 14"],
        "{{Estabilização de voçorocas}}": df.loc[0, "Unnamed: 15"],
        "{{Manejo florestal}}": df.loc[0, "Unnamed: 16"],
        "{{Outras}}": df.loc[0, "Unnamed: 17"],
        "{{População residente}}": df.loc[10, "Tucano"],
        "{{Área da unidade territorial}}": df.loc[11, "Tucano"],
        "{{Densidade demográfica}}": df.loc[12, "Tucano"],
        "{{PIB}}": df.loc[17, "Tucano"],
        "{{Percentual da agricultura}}": df.loc[18, "Tucano"],
        "{{Valor Adicionado Bruto Agropecuária}}": df.loc[25, "Tucano"],
        "{{Valor Adicionado Bruto Indústria}}": df.loc[26, "Tucano"],
        "{{Valor Adicionado Bruto Serviços}}": df.loc[27, "Tucano"],
        "{{Valor Adicionado Bruto Administração Pública}}": df.loc[28, "Tucano"]
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

