import streamlit as st
import pandas as pd
from docx import Document
from io import BytesIO

def extract_table_from_docx(docx_file):
    doc = Document(docx_file)
    for table in doc.tables:
        data = []
        for row in table.rows:
            data.append([cell.text.strip() for cell in row.cells])
        df = pd.DataFrame(data)
        return df  # Возвращаем первую найденную таблицу
    return None

def main():
    st.title("DOCX to CSV Converter")
    uploaded_file = st.file_uploader("Upload a DOCX file", type=["docx"])
    
    if uploaded_file:
        df = extract_table_from_docx(uploaded_file)
        
        if df is not None:
            st.write("Extracted Table:")
            st.dataframe(df)
            
            column_names = []
            for i in range(df.shape[1]):
                column_names.append(st.text_input(f"Column {i+1} name", value=f"Column_{i+1}"))
            
            df.columns = column_names
            
            st.write("Updated Table:")
            st.dataframe(df)
            
            buffer = BytesIO()
            df.to_csv(buffer, index=False, encoding='utf-8')
            buffer.seek(0)
            
            st.download_button(
                label="Download CSV",
                data=buffer,
                file_name="converted_table.csv",
                mime="text/csv"
            )
        else:
            st.error("No tables found in the document.")

if __name__ == "__main__":
    main()
