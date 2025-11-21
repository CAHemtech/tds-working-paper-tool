import io
import pandas as pd
import streamlit as st

st.set_page_config(page_title="TDS Working Paper Tool", layout="wide")

st.title("TDS Working Paper Tool")

st.write("""
Upload your Tally export Excel file (one ledger per sheet) and generate a single TDS working paper.

You only need to:
- Enter how many **fixed columns** are common across all sheets (X)
- Enter **TDS ledger names / keywords** used in your file

The app will:
- Detect the ledger table starting from the row where the **first column is 'Date'**
- Keep the first **X columns** and treat the last of those as **`Ledger_Amount`**
- Remove rows where **Particulars** contains **'Grand Total'**
- Find all columns whose header contains your **TDS keywords** and sum them into **`TDS_Amount`**
- Add **`LedgerName`** from the sheet name
- Merge all sheets into one clean **TDS Working Paper**.
""")

# 1) File upload
uploaded_file = st.file_uploader("Upload Tally Excel file (.xlsx)", type=["xlsx"])

# 2) Number of fixed columns
fixed_cols = st.number_input(
    "Number of fixed columns (X)",
    min_value=1,
    max_value=100,
    value=21,
    step=1,
    help="These are the first X columns that are common in every sheet. For your current file, 21 works."
)

# 3) TDS keywords / ledger names
tds_keywords_input = st.text_input(
    "TDS ledger names / keywords (comma-separated)",
    value="tds payable, tds payable-llp, interest on tds, tds",
    help="Any column whose header contains these strings (case-insensitive) will be treated as TDS columns and summed row-wise."
)

run_button = st.button("Generate TDS Working Paper")


def process_file(file, fixed_cols: int, tds_keywords_raw: str) -> pd.DataFrame | None:
    xls = pd.ExcelFile(file)
    all_ledgers: list[pd.DataFrame] = []
    template_cols: list[str] | None = None

    # Clean keywords
    tds_keywords = [k.strip().lower() for k in tds_keywords_raw.split(",") if k.strip()]
    if not tds_keywords:
        tds_keywords = ["tds"]

    for sheet_name in xls.sheet_names:
        raw = pd.read_excel(xls, sheet_name=sheet_name, header=None)

        if raw.dropna(how="all").empty:
            continue

        # 1️⃣ Detect header row: first column value == 'Date'
        header_row_idx = None
        first_col = raw.iloc[:, 0]
        for i, v in first_col.items():
            if str(v).strip().lower() == "date":
                header_row_idx = i
                break

        if header_row_idx is None:
            # Sheet doesn't contain the expected table structure
            continue

        # 2️⃣ Build proper DataFrame
        header = raw.iloc[header_row_idx].tolist()
        data = raw.iloc[header_row_idx + 1 :].copy()
        data.columns = header
        data = data.dropna(how="all")
        if data.empty:
            continue

        # 3️⃣ Fixed first X columns
        base_cols = data.columns[:fixed_cols]
        base_data = data.loc[:, base_cols].copy()

        # Rename last fixed column as Ledger_Amount
        base_data.rename(columns={base_cols[-1]: "Ledger_Amount"}, inplace=True)

        # 4️⃣ Remove 'Grand Total' rows (if Particulars column exists)
        particulars_col = None
        for col in base_data.columns:
            if "particular" in str(col).lower():
                particulars_col = col
                break

        if particulars_col:
            base_data = base_data[
                ~base_data[particulars_col]
                .astype(str)
                .str.lower()
                .str.contains("grand total")
            ]

        # 5️⃣ Lock / align to a fixed template for first X columns
        if template_cols is None:
            template_cols = base_data.columns.tolist()
        else:
            # Add missing template columns as NaN, drop extras, reorder
            for col in template_cols:
                if col not in base_data.columns:
                    base_data[col] = pd.NA
            base_data = base_data[template_cols]

        # 6️⃣ Find ALL TDS-related columns and sum them
        tds_cols: list[str] = []
        for col in data.columns:
            col_lower = str(col).lower()
            if any(kw in col_lower for kw in tds_keywords):
                tds_cols.append(col)

        if tds_cols:
            tds_sum = data[tds_cols].apply(pd.to_numeric, errors="coerce").sum(axis=1)
            base_data["TDS_Amount"] = tds_sum
        else:
            base_data["TDS_Amount"] = pd.NA

        # 7️⃣ Ledger name from sheet
        base_data["LedgerName"] = sheet_name

        all_ledgers.append(base_data)

    if not all_ledgers:
        return None

    final_df = pd.concat(all_ledgers, ignore_index=True)
    final_cols = template_cols + ["TDS_Amount", "LedgerName"]
    final_df = final_df[final_cols]
    return final_df


if run_button:
    if uploaded_file is None:
        st.error("Please upload an Excel file first.")
    else:
        with st.spinner("Processing file..."):
            result_df = process_file(uploaded_file, fixed_cols, tds_keywords_input)

        if result_df is None:
            st.error("No valid data found. Check the file structure / header row.")
        else:
            st.success(f"TDS Working Paper created successfully. Total rows: {len(result_df)}")

            st.subheader("Preview")
            st.dataframe(result_df.head(100))

            # Prepare Excel file for download
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                result_df.to_excel(writer, index=False, sheet_name="TDS_Working_Paper")
            buffer.seek(0)

            st.download_button(
                label="⬇️ Download TDS Working Paper (Excel)",
                data=buffer,
                file_name="TDS_Working_Paper.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

st.markdown("---")
st.caption("TDS Working Paper Tool")

