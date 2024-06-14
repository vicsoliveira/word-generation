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

    # Display the DataFrame columns for debugging
    st.write("Columns in the Excel file:")
    st.write(df.columns)

    # Read the Word document
    doc = Document(uploaded_word)

    # Define the mapping from Excel to Word placeholders
    data_mapping = {
        "{{Nome do município}}": df.at[0, "Unnamed: 3"],
        "{{Área total}}": df.at[0, "de 50 a 500 ha"], # Adjusted based on inspection
        "{{Plantio em nível}}": df.at[0, 422], # Column index or appropriate name
        "{{Rotação de culturas}}": df.at[0, 12],
        "{{Pousio ou descanso}}": df.at[0, 51],
        "{{Proteção de encostas}}": df.at[0, 58],
        "{{Recuperação de mata ciliar}}": df.at[0, 4],
        "{{Reflorestamento de nascentes}}": df.at[0, 1],
        "{{Estabilização de voçorocas}}": df.at[0, 0],
        "{{Manejo florestal}}": df.at[0, '1.1'], # Example, adjust based on actual columns
        "{{Outras}}": df.at[0, 8],
        "{{População residente}}": df.at[10, "Tucano"],
        "{{Área da unidade territorial}}": df.at[11, "Tucano"],
        "{{Densidade demográfica}}": df.at[12, "Tucano"],
        "{{PIB}}": df.at[17, "Tucano"],
        "{{Percentual da agricultura}}": df.at[18, "Tucano"],
        "{{Valor Adicionado Bruto Agropecuária}}": df.at[25, "Tucano"],
        "{{Valor Adicionado Bruto Indústria}}": df.at[26, "Tucano"],
        "{{Valor Adicionado Bruto Serviços}}": df.at[27, "Tucano"],
        "{{Valor Adicionado Bruto Administração Pública}}": df.at[28, "Tucano"]
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

