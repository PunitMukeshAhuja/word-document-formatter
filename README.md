# Word Document Formatter v4

A Streamlit application for cleaning and standardizing Word documents.

## New in v4

- Editable preview using Streamlit Data Editor
- Manually change any paragraph to:
  - Title
  - Heading 1
  - Heading 2
  - Heading 3
  - Body
- Automatic detection remains visible alongside manual classification
- Manual overrides are applied to the final downloaded Word document
- Reset button restores automatic heading detection
- Updated heading/body counts reflect manual changes

## Run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Update an existing Streamlit deployment

Replace the existing `app.py` and `requirements.txt` in GitHub and commit the changes.
Streamlit Community Cloud should redeploy automatically.
