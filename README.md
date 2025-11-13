# Two-Step Order Cleaner

Streamlit app for transforming horizontal promo order files into an ERP-ready format
by **clearing orders for selected clients** while keeping the original structure.

## Key Rules

- Incoming order files are **non-changeable**; the app only creates a modified copy.
- Supplier column: `ძირითადი მომწოდებელი` (first row).
- Protected supplier: rows where supplier == `გაგრა პლუსი` are never changed.
- Targeted clear:
  - For each column where the **shop_code** or **shop_nickname** appears in
    `config/client_removal_template.xlsx` (sheet `clients_to_clear`):
    - Clear values in that column for rows where supplier != `გაგრა პლუსი`.
- Drop all columns whose first-row header starts with `დასავლეთი`.

## Structure

- `two_step_order_cleaner.py` — main Streamlit app.
- `config/client_removal_template.xlsx` — default client removal template.
- `config/README.md` — description of config files.
- `requirements.txt` — Python dependencies.
- `.gitignore` — standard ignores for Python/Streamlit.

## How to Run Locally

```bash
pip install -r requirements.txt
streamlit run two_step_order_cleaner.py
```

Then open the URL printed in the terminal (usually http://localhost:8501).

## How to Use

1. Prepare your order file from the two-step system as Excel.
2. Update `config/client_removal_template.xlsx` with the list of clients
   you want to clear (by `shop_code` and/or `shop_nickname_optional`).
3. Run the app and upload the order file.
4. Download the cleaned Excel file from the app.
