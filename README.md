
# Word Document Formatter

A simple Streamlit application that formats `.docx` documents consistently.

## Features

- Upload Word `.docx` files
- Standardize body font and font size
- Configure paragraph alignment
- Configure line spacing and paragraph spacing
- Configure document margins
- Format Heading 1, Heading 2 and Heading 3
- Add centered page numbers
- Preserve document content, images and tables as much as `python-docx` allows
- Download the formatted Word document

## Run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Deploy on Streamlit Community Cloud

1. Create a GitHub repository.
2. Upload `app.py` and `requirements.txt`.
3. Go to Streamlit Community Cloud.
4. Choose **Create app**.
5. Select your GitHub repository.
6. Set the entry point to `app.py`.
7. Deploy.
8. Share the generated `.streamlit.app` URL.

## Important

Do not upload confidential documents to a public application unless you are comfortable with the hosting/privacy setup.
