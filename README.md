# Excel Deck

A simple MVP for turning Excel data into polished PowerPoint presentations.

Features:
- Upload `.xlsx` or `.csv` files
- Preview sheets and columns
- Select which columns to include in the presentation
- Apply calculation rules such as:
  - direct value
  - subtraction
  - multiplication
  - growth percentage
  - percentage of total
- Generate a PowerPoint deck with a summary slide and chart-ready data
- Download the `.pptx` file

## Tech stack
- Backend: FastAPI + Python
- Excel processing: openpyxl + pandas
- PowerPoint export: python-pptx
- Frontend: plain HTML + JavaScript

## Project structure

```text
backend/
  app/
    main.py
    services/
      excel_service.py
      ppt_service.py
frontend/
  index.html
  styles.css
  app.js
requirements.txt
README.md
.gitignore
```

## Local setup

1. Create a virtual environment:

```bash
python -m venv .venv
source .venv/bin/activate
```

2. Install dependencies:

```bash
pip install -r requirements.txt
```

3. Start the application:

```bash
uvicorn backend.app.main:app --reload
```

4. Open the browser at:

```text
http://localhost:8000
```

## Example usage

1. Upload an Excel file.
2. Choose the sheet and select the columns you want in the deck.
3. Pick a formula for each selected column.
4. Click `Generate PPT`.
5. Download the created presentation.

## MVP formulas supported

- direct
- subtraction
- multiplication
- growth_percent
- percentage_of_total

## Notes

This is intentionally a strong MVP for fast validation. It can be extended with:
- better templates and theme layouts
- AI-based recommendations
- user authentication
- chart styling
- multi-sheet support
- reusable report templates
