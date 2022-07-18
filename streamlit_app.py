import pandas as pd
from docxtpl import DocxTemplate
from slugify import slugify
import streamlit as st
import pandas as pd
import os
import time
from zipfile import ZipFile

st.set_page_config(layout="centered", page_icon="📄", page_title="Documents generator")

left, right = st.columns([4, 1])
left.title("📄 Documents generator")


left.markdown("""
    This app will generate documents files (`*.docx`) based on a template (`*.docx`) and a list of variables in Excel format (`*.xlsx`).

    After processing the files, the app will generate a zip file with all the documents.

    If the source of variables contains a column with the name `filename`, the app will use it to generate the filename of the document. Otherwise, the app will use the index of the row to generate the filename.
    """
)


right.image("https://raw.githubusercontent.com/data-cfwb/.github/main/logo_data_office.png", width=150)

st.image("./schema.drawio.png", width=400, caption="How does it work?")

form = st.form("template_form")
vars_file = form.file_uploader("Upload your variables file", type="xlsx")
tpl_file = form.file_uploader("Upload your template file", type="docx")

submit = form.form_submit_button("Generate Docx")

if submit:
    st.write("Generating...")

    vars_filename = "source_vars.xlsx"
    tpl_filename = "source_tpl.docx"

    with open(os.path.join("source",vars_filename),"wb") as f: 
      f.write(vars_file.getbuffer())         

    with open(os.path.join("source",tpl_filename),"wb") as f: 
      f.write(tpl_file.getbuffer())         

    df = pd.read_excel("./source/" + vars_filename)

    # define an object for results
    results = []

    zip_file = './results/results.zip'
    zipObj = ZipFile(zip_file, 'w')

    use_filename_column = False
    #check if filename column exist and is unique
    if 'filename' in df.columns and df['filename'].is_unique: 
        use_filename_column = True

    for idx, context in enumerate(df.to_dict(orient='records')):
        doc = DocxTemplate("source/" + tpl_filename)
        doc.render(context)
        resulted_file = "./results/result_" + str(idx) + ".docx"
        if use_filename_column:
            filename = slugify(context['filename'], separator='_', lowercase=False)
            resulted_file = "./results/" + filename + ".docx"
        doc.save(resulted_file)
        
        # append file to zip
        zipObj.write(resulted_file)
        results.append(resulted_file)    
        
    zipObj.close()
    # delete temporary files
    for result in results:
        os.remove(result)

    st.balloons()
    
    my_bar = st.progress(0)

    for result in range(len(results)):
        time.sleep(0.1)
        percent = (result + 1) / len(results) * 100
        my_bar.progress(int(percent))

    st.success("🎉 Your files are ready")
    
    with open(zip_file, "rb") as fp:
        st.download_button(
            label="Download ZIP",
            data=fp,
            file_name="results.zip",
            mime="application/zip"
        )
# align right
st.markdown("""
This app is a work in progress and has been made by the [Data Office](https://github.com/data-cfwb) of CFWB.
""")