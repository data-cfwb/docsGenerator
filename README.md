## DocsGenerator

This app is a work in progress and has been made by the [Data Office](https://github.com/data-cfwb) of CFWB.

# Run it

First install the dependencies:

```bash
python3 -m pip install -r requirements.txt
```

Then run the app:

```bash
streamlit run streamlit_app.py
```

# How does it work?

This app will generate documents files (`*.docx`) based on a template (`*.docx`) and a list of variables in Excel format (`*.xlsx`).

After processing the files, the app will generate a zip file with all the documents.

If the source of variables contains a column with the name `filename`, the app will use it to generate the filename of the document. Otherwise, the app will use the index of the row to generate the filename.

![](https://raw.githubusercontent.com/data-cfwb/docsGenerator/main/schema.drawio.png)

# Demo

- Link to Imgur : <https://i.imgur.com/RddjA1G.gif>

<img src="https://i.imgur.com/RddjA1G.gif" />
