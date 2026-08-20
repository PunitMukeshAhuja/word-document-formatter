# Word Document Formatter v6

## New in v6

- Grid now shows proposed output formatting:
  - output font size
  - bold
  - alignment
  - line spacing
  - paragraph spacing
- Heading/body size changes are visible in the review grid.
- Manual paragraph-type correction remains editable.
- Added print-style visual document preview before download.
- Preview reflects current:
  - font
  - body size
  - title size
  - heading sizes
  - alignment
  - spacing
  - margins
  - manual structure corrections
- Generate Word document only after visual review.

## Preview limitation

The browser preview is an approximate structural/visual preview. Microsoft Word may paginate
slightly differently because browser rendering and Word's layout engine are different.

## Deploy

Replace `app.py` and `requirements.txt` in your existing GitHub repository and commit.
Streamlit Community Cloud should automatically redeploy.
