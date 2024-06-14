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

# File uploaders for CSV file and Word template
uploaded_csv = st.file_uploader("Choose a CSV file", type="csv")
uploaded_word = st.file_uploader("Choose a Word template", type="docx")

if uploaded_csv is not None and uploaded_word is not None:
    # Load the CSV file and inspect the first few rows
    df = pd.read_csv(uploaded_csv)
    st.write("Raw CSV data:")
    st.write(df.head(20))
    
    # Clean and align the data
    df_cleaned = df.iloc[6:].reset_index(drop=True)
    df_cleaned.columns = df.iloc[5]
    df_cleaned = df_cleaned.dropna(how='all').reset_index(drop=True)
    
    st.write("Cleaned DataFrame:")
    st.write(df_cleaned.head(20))

    # Read the Word document
    doc = Document(uploaded_word)

    # Define the mapping from CSV to Word placeholders
    data_mapping = {
        "{{Nome do município}}": df_cleaned.at[0, "Unnamed: 1"],
        "{{População residente}}": df_cleaned.at[0, "População residente (Pessoas)"],
        "{{Área da unidade territorial}}": df_cleaned.at[1, "Área da unidade territorial (Quilômetros quadrados)"],
        "{{Densidade demográfica}}": df_cleaned.at[2, "Densidade demográfica (Habitante por quilômetro quadrado)"],
        "{{Área total}}": df_cleaned.at[7, "Área total do estabelecimento agropecuário"],
        "{{Plantio em nível}}": df_cleaned.at[8, "Plantio em nível"],
        "{{Rotação de culturas}}": df_cleaned.at[9, "Rotação de culturas"],
        "{{Pousio ou descanso}}": df_cleaned.at[10, "Pousio ou descanso de solos"],
        "{{Proteção de encostas}}": df_cleaned.at[11, "Proteção e/ou conservação de encostas"],
        "{{Recuperação de mata ciliar}}": df_cleaned.at[12, "Recuperação de mata ciliar"],
        "{{Reflorestamento de nascentes}}": df_cleaned.at[13, "Reflorestamento para proteção de nascentes"],
        "{{Estabilização de voçorocas}}": df_cleaned.at[14, "Estabilização de voçorocas"],
        "{{Manejo florestal}}": df_cleaned.at[15, "Manejo florestal"],
        "{{Outras}}": df_cleaned.at[16, "Outras"],
        "{{PIB}}": df_cleaned.at[17, "Produto Interno Bruto - PIB (Mil R$)"],
        "{{Percentual da agricultura}}": df_cleaned.at[18, "Percentual da agricultura no PIB"],
        "{{Valor Adicionado Bruto Agropecuária}}": df_cleaned.at[25, "Valor Adicionado Bruto Agropecuária"],
        "{{Valor Adicionado Bruto Indústria}}": df_cleaned.at[26, "Valor Adicionado Bruto Indústria"],
        "{{Valor Adicionado Bruto Serviços}}": df_cleaned.at[27, "Valor Adicionado Bruto Serviços"],
        "{{Valor Adicionado Bruto Administração Pública}}": df_cleaned.at[28, "Valor Adicionado Bruto Administração Pública"]
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


