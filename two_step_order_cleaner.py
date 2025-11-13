# -*- coding: utf-8 -*-
"""
Two-Step Order Cleaner — Targeted Clear by shop_code and nickname

Policy: Incoming order files are NON-CHANGEABLE; we adapt on our side.
The original workbook is never overwritten. We only produce a modified copy
for download.

Logic:
1) Order file structure (TDSheet):
   - Row 1 (index 0): nicknames, plus master fields
       e.g. "ძირითადი მომწოდებელი", "ნომენკლატურა", "ძირითადი შტრიხ-კოდი",
            "აღმოსავლეთი", "დასავლეთი", "სულ", shop nicknames ("ვაკე", ...)
   - Row 2 (index 1): addresses with #ID# tokens
       e.g. "#003# ქ.თბილისი, წყნეთის ქ. #2"
   - Row 3 (index 2): labels
       e.g. "შესაკვეთი რაოდენობა"
   - Data rows start at row index 3

2) Client removal template (config/client_removal_template.xlsx):
   Sheet 'clients_to_clear':
       - "shop_code" (required, e.g. 003, 465, 037; numeric or text)
       - "shop_nickname_optional" (optional, e.g. "ვაკე", "ვანთა", ...)
       - "notes_optional" (ignored by logic, for your comments)

3) Business rules:
   - Protected supplier: any row where supplier == "გაგრა პლუსი" is not altered.
   - Targeted clear:
       For each column where
           shop_code ∈ template.shop_code
       OR  nickname ∈ template.shop_nickname_optional,
       clear the cells (set to empty "") in all rows where supplier != "გაგრა პლუსი".
   - Drop “West” aggregates:
       Any column whose first-row header starts with "დასავლეთი"
       (nickname row, strip + startswith) is removed entirely.
   - Everything else (sheet name, other columns, header rows) is preserved.

Usage (local):
    pip install -r requirements.txt
    streamlit run two_step_order_cleaner.py
"""

import io
import re
from pathlib import Path
from typing import Dict, Set, Tuple

import pandas as pd
import streamlit as st


# ---------------------------------------------------------------------------
# Core transformation
# ---------------------------------------------------------------------------

def load_template_from_file(template_path: Path) -> Tuple[Set[str], Set[str]]:
    """Load client removal template from a local Excel file."""
    if not template_path.exists():
        raise FileNotFoundError(f"Template not found at {template_path}")

    tpl = pd.read_excel(
        template_path,
        sheet_name="clients_to_clear",
        dtype=str
    ).fillna("")

    if "shop_code" not in tpl.columns:
        raise ValueError("Template must contain a column named 'shop_code'.")

    if "shop_nickname_optional" not in tpl.columns:
        tpl["shop_nickname_optional"] = ""

    tpl["shop_code"] = (
        tpl["shop_code"]
        .astype(str)
        .str.replace(r"\\.0$", "", regex=True)
        .str.strip()
    )
    tpl["shop_nickname_optional"] = tpl["shop_nickname_optional"].astype(str).str.strip()

    shop_codes = set(tpl["shop_code"][tpl["shop_code"] != ""].tolist())
    nicknames = set(
        tpl["shop_nickname_optional"][tpl["shop_nickname_optional"] != ""].tolist()
    )

    return shop_codes, nicknames


def load_template_from_bytes(file_bytes: bytes) -> Tuple[Set[str], Set[str]]:
    """Load client removal template from an uploaded Excel file (Streamlit)."""
    tpl = pd.read_excel(
        io.BytesIO(file_bytes),
        sheet_name="clients_to_clear",
        dtype=str
    ).fillna("")

    if "shop_code" not in tpl.columns:
        raise ValueError("Template must contain a column named 'shop_code'.")

    if "shop_nickname_optional" not in tpl.columns:
        tpl["shop_nickname_optional"] = ""

    tpl["shop_code"] = (
        tpl["shop_code"]
        .astype(str)
        .str.replace(r"\\.0$", "", regex=True)
        .str.strip()
    )
    tpl["shop_nickname_optional"] = tpl["shop_nickname_optional"].astype(str).str.strip()

    shop_codes = set(tpl["shop_code"][tpl["shop_code"] != ""].tolist())
    nicknames = set(
        tpl["shop_nickname_optional"][tpl["shop_nickname_optional"] != ""].tolist()
    )

    return shop_codes, nicknames


def find_col(hdr_series: pd.Series, name_exact: str):
    """Find first column index in hdr_series where stripped value == name_exact."""
    for i, v in enumerate(hdr_series):
        if str(v).strip() == name_exact:
            return i
    return None


def transform_order(
    order_bytes: bytes,
    shop_codes_to_clear: Set[str],
    nicknames_to_clear: Set[str],
    protected_supplier: str = "გაგრა პლუსი",
    west_prefix: str = "დასავლეთი",
) -> Tuple[bytes, Dict]:
    """
    Apply business rules to the order workbook and return:
        (modified_workbook_bytes, summary_dict)
    """

    xls = pd.ExcelFile(io.BytesIO(order_bytes))
    sheet_name = xls.sheet_names[0]
    raw = xls.parse(sheet_name, header=None, dtype=object)

    # Header rows
    hdr_nickname = raw.iloc[0].astype(str).fillna("")  # row 0
    hdr_address = raw.iloc[1].astype(str).fillna("")   # row 1

    # Supplier column
    col_supplier = find_col(hdr_nickname, "ძირითადი მომწოდებელი")
    if col_supplier is None:
        raise RuntimeError(
            "Could not find supplier column 'ძირითადი მომწოდებელი' in the first row."
        )

    # Map shop columns by shop_code from address row (#123#)
    shop_cols_map: Dict[int, str] = {}
    for i, meta in enumerate(hdr_address):
        m = re.search(r"#(\\d+)#", str(meta))
        if m:
            shop_cols_map[i] = m.group(1)

    columns_to_clear = set()

    # Match by shop_code
    for col_idx, code in shop_cols_map.items():
        if code in shop_codes_to_clear:
            columns_to_clear.add(col_idx)

    # Match by nickname
    for col_idx, nickname in enumerate(hdr_nickname):
        if nickname.strip() in nicknames_to_clear:
            columns_to_clear.add(col_idx)

    # Detect "West" columns
    west_cols = [
        i
        for i, v in enumerate(hdr_nickname)
        if str(v).strip().startswith(west_prefix)
    ]

    out = raw.copy()

    # Data rows start at index 3
    data_start = 3

    suppliers = out.iloc[data_start:, col_supplier].astype(str).fillna("")
    mask_not_protected = suppliers.str.strip() != protected_supplier
    rows_to_edit = suppliers.index[mask_not_protected]

    cleared_cells = 0
    for c in columns_to_clear:
        col_slice = out.loc[rows_to_edit, c]
        cleared_cells += int(col_slice.notna().sum())
        out.loc[rows_to_edit, c] = ""

    if west_cols:
        out = out.drop(columns=west_cols)

    output_buffer = io.BytesIO()
    with pd.ExcelWriter(output_buffer, engine="xlsxwriter") as writer:
        out.to_excel(
            writer,
            index=False,
            header=False,
            sheet_name=sheet_name,
        )
    output_buffer.seek(0)

    summary = {
        "sheet_name": sheet_name,
        "columns_to_clear_count": len(columns_to_clear),
        "west_columns_dropped": len(west_cols),
        "rows_eligible_by_supplier_rule": int(mask_not_protected.sum()),
        "cleared_cells_estimate": cleared_cells,
        "protected_supplier": protected_supplier,
    }

    return output_buffer.getvalue(), summary


# ---------------------------------------------------------------------------
# Streamlit UI
# ---------------------------------------------------------------------------

def main():
    st.set_page_config(
        page_title="Two-Step Order Cleaner",
        layout="wide",
    )

    st.title("Two-Step Order Cleaner")
    st.caption(
        "Incoming order files are non-changeable; the app only creates a "
        "modified copy for ERP upload, based on your client removal list."
    )

    st.markdown("### 1. Upload order file")

    order_file = st.file_uploader(
        "Order file (Excel)",
        type=["xlsx", "xls"],
        key="order_file",
    )

    st.markdown("---")
    st.markdown("### 2. Template source")

    config_template_path = Path("config/client_removal_template.xlsx")
    use_config_template = st.checkbox(
        f"Use default template from {config_template_path}",
        value=True,
    )

    uploaded_template_file = None
    if not use_config_template:
        uploaded_template_file = st.file_uploader(
            "OR upload a custom client removal template (Excel with 'clients_to_clear')",
            type=["xlsx", "xls"],
            key="template_file",
        )

    st.markdown("---")
    st.markdown("### 3. Parameters")

    protected_supplier = st.text_input(
        "Protected supplier (rows with this value are never touched):",
        value="გაგრა პლუსი",
    )
    west_prefix = st.text_input(
        "Prefix for 'West' aggregate columns to drop:",
        value="დასავლეთი",
    )

    st.markdown("---")
    st.markdown("### 4. Run transformation")

    run_btn = st.button("Run cleaning and generate download", type="primary")

    if run_btn:
        if order_file is None:
            st.error("Please upload the order file.")
            return

        try:
            # Load template (config vs upload)
            if use_config_template:
                shop_codes, nicknames = load_template_from_file(config_template_path)
            else:
                if uploaded_template_file is None:
                    st.error("Please upload a template or enable 'Use default template'.")
                    return
                shop_codes, nicknames = load_template_from_bytes(
                    uploaded_template_file.getvalue()
                )

            if not shop_codes and not nicknames:
                st.warning(
                    "Template does not contain any shop_code or nickname values. "
                    "Nothing to clear."
                )
                return

            cleaned_bytes, summary = transform_order(
                order_bytes=order_file.getvalue(),
                shop_codes_to_clear=shop_codes,
                nicknames_to_clear=nicknames,
                protected_supplier=protected_supplier,
                west_prefix=west_prefix,
            )

        except Exception as e:
            st.error(f"Error during transformation: {e}")
            st.stop()

        st.success("Cleaning completed. Review the summary and download the new file.")

        st.subheader("Summary")
        st.json(summary)

        st.download_button(
            label="Download cleaned order file",
            data=cleaned_bytes,
            file_name="Ori_Nabiji_შეკვეთა(ასატვირთი).xlsx",
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
        )

        st.subheader("Preview (top 10 rows after cleaning)")
        try:
            preview_df = pd.read_excel(
                io.BytesIO(cleaned_bytes),
                sheet_name=summary["sheet_name"],
                header=None,
            ).head(10)
            st.dataframe(preview_df)
        except Exception:
            st.info("Preview not available, but the file is ready for download.")


if __name__ == "__main__":
    main()
