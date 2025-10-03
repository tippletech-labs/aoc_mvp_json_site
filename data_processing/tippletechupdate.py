import json
import os
import re
import glob

def validate_input_path(candidate_path: str) -> str:
    """
    Validate that the input file:
      1) exists,
      2) is an .xlsx file,
      3) matches wine-properties-YYYY-MM-DD.xlsx naming.
    Return the normalized path or raise ValueError with a clear reason.
    """
    if os.path.isdir(candidate_path):
        candidates = glob.glob(os.path.join(candidate_path, "wine-properties-*.xlsx"))
        if not candidates:
            raise ValueError(
                "No file found matching pattern 'wine-properties-YYYY-MM-DD.xlsx' in the provided directory."
            )
        if len(candidates) > 1:
            raise ValueError(
                "Multiple files match 'wine-properties-YYYY-MM-DD.xlsx'. "
                "Please keep only one or specify the exact file path."
            )
        candidate_path = candidates[0]

    if not os.path.exists(candidate_path):
        raise ValueError(f"Input file not found: {candidate_path}")

    filename = os.path.basename(candidate_path)

    # Enforce .xlsx
    if not filename.lower().endswith(".xlsx"):
        raise ValueError("Invalid file type: expected an .xlsx Excel file.")

    # Enforce exact pattern wine-properties-YYYY-MM-DD.xlsx
    pattern = r"^wine-properties-\d{4}-\d{2}-\d{2}\.xlsx$"
    if not re.match(pattern, filename):
        raise ValueError(
            "Invalid file name: expected 'wine-properties-YYYY-MM-DD.xlsx' "
            "(e.g., wine-properties-2025-10-03.xlsx)."
        )
    return candidate_path


def parse_data(input_path: str, output_file: str) -> None:
    import pandas as pd
    # Give a clearer error if openpyxl is missing
    try:
        import openpyxl  # noqa: F401
    except ImportError as e:
        raise RuntimeError(
            "openpyxl is required to read .xlsx files. Inside your venv run: pip install openpyxl"
        ) from e

    input_file = validate_input_path(input_path)

    # Required columns from Excel — NOTE: Country ID will be GENERATED from Country
    required_cols = [
        "Name",
        "Producer",
        "Country",  # used to generate Country ID and for Search String
        "Product Type",
        "Region",
        "Custom Field:Franchise Tag",
        "Marketplace: Available?",
    ]

    # Optional columns used only for the Search String
    search_candidates = ["Appellation", "Varietal", "Importer"]

    usecols = sorted(set(required_cols + search_candidates))

    # Read only what we might need; keep everything as strings for consistency
    df = pd.read_excel(
        input_file,
        dtype=str,
        usecols=lambda c: c in usecols,
        engine="openpyxl",
    ).fillna("Unknown")

    # Validate required columns exist
    missing = [c for c in required_cols if c not in df.columns]
    if missing:
        raise ValueError("Missing required column(s): " + ", ".join(missing))

    # Filter to rows where Marketplace: Available? == "yes" (case-insensitive)
    avail_col = "Marketplace: Available?"
    mask_yes = df[avail_col].astype(str).str.strip().str.lower().eq("yes")
    df = df.loc[mask_yes].copy()

    # If empty after filtering, still write a valid empty payload
    if df.empty:
        os.makedirs(os.path.dirname(output_file), exist_ok=True)
        with open(output_file, "w", encoding="utf-8") as f:
            json.dump({"products": []}, f, indent=4, ensure_ascii=False)
        return

    # ---- Generate Country ID from Country (order of first appearance) ----
    # pd.factorize returns (codes, uniques); codes start at 0, so add 1
    codes, _uniques = pd.factorize(df["Country"], sort=False)
    country_ids = (codes + 1)

    # ---- Helper: always return a Series (never a bare string) ----
    def series_or_default(col_name: str, default_value: str = "Unknown") -> pd.Series:
        """Return df[col_name] as a string Series if present; otherwise a same-length Series of default_value."""
        if col_name in df.columns:
            return df[col_name].astype(str)
        # Same-length Series filled with default_value
        return pd.Series(default_value, index=df.index)

    # ---- Build Search String (vectorized) ----
    # "{Name} {Country} {Region} {Appellation} {Varietal} {Importer}".lower()
    search_string = (
        series_or_default("Name") + " " +
        series_or_default("Country") + " " +
        series_or_default("Region") + " " +
        series_or_default("Appellation") + " " +
        series_or_default("Varietal") + " " +
        series_or_default("Importer")
    ).str.lower().str.replace(r"\s+", " ", regex=True).str.strip()

    # ---- Final output (only requested fields + generated Country ID + Search String) ----
    out_df = pd.DataFrame({
        "Name": df["Name"].astype(str),
        "Producer": df["Producer"].astype(str),
        "Country ID": country_ids,  # GENERATED here
        "Product Type": df["Product Type"].astype(str),
        "Region": df["Region"].astype(str),
        "Search String": search_string,
        "Custom Field:Franchise Tag": df["Custom Field:Franchise Tag"].astype(str),
        "Marketplace: Available?": df["Marketplace: Available?"].astype(str),
    })

    records = out_df.to_dict(orient="records")
    payload = {"products": records}

    os.makedirs(os.path.dirname(output_file), exist_ok=True)
    with open(output_file, "w", encoding="utf-8") as f:
        json.dump(payload, f, indent=4, ensure_ascii=False)


if __name__ == "__main__":
    # By default, look in ./data for exactly one matching Excel
    input_path = os.path.join("data")
    output_file = os.path.join("docs", "data.json")
    parse_data(input_path, output_file)
