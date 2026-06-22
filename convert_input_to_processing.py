from pathlib import Path

import pandas as pd


BASE_DIR = Path(__file__).resolve().parent
INPUT_DIR = BASE_DIR / "input"
PROCESSING_DIR = BASE_DIR / "processing"

SUPPORTED_EXTENSIONS = {".xlsx", ".xls", ".xlsm", ".xlsb", ".csv"}


def _canonical_name(name: object) -> str:
    value = str(name).strip().lower()
    for token in (" ", "_", "-", "/", "(", ")", '"', "'", ":", ".", ",", "#"):
        value = value.replace(token, "")
    return value


def _normalize_and_deduplicate_headers(headers: list[object]) -> list[str]:
    cleaned: list[str] = []
    seen: dict[str, int] = {}

    for idx, header in enumerate(headers, start=1):
        value = str(header).strip() if header is not None else ""
        if not value or value.lower() == "nan":
            value = f"column_{idx}"

        base = value
        counter = seen.get(base, 0)
        if counter:
            value = f"{base}_{counter + 1}"
        seen[base] = counter + 1
        cleaned.append(value)

    return cleaned


def _headers_look_invalid(df: pd.DataFrame) -> bool:
    columns = [str(c).strip() for c in df.columns]
    if not columns:
        return False

    unnamed_count = sum(1 for c in columns if not c or c.lower().startswith("unnamed:"))
    numeric_count = sum(1 for c in columns if c.isdigit())

    # If most headers are unnamed or numeric indexes, likely true headers are in row 1.
    return unnamed_count >= max(1, len(columns) // 2) or numeric_count == len(columns)


def _detect_header_row_index(df: pd.DataFrame, max_scan_rows: int = 25) -> int | None:
    if df.empty:
        return None

    scan_limit = min(len(df), max_scan_rows)
    best_idx: int | None = None
    best_score = 0.0

    for i in range(scan_limit):
        row = df.iloc[i]
        values = [
            str(v).strip()
            for v in row
            if pd.notna(v) and str(v).strip() and str(v).strip().lower() != "nan"
        ]
        if not values:
            continue

        non_empty_ratio = len(values) / max(1, len(df.columns))
        text_like_count = sum(1 for v in values if not v.replace(".", "", 1).isdigit())
        text_ratio = text_like_count / len(values)
        score = (non_empty_ratio * 0.7) + (text_ratio * 0.3)

        if score > best_score:
            best_score = score
            best_idx = i

    # Conservative threshold to avoid false positives on sparse metadata rows.
    if best_idx is None or best_score < 0.45:
        return None
    return best_idx


def _promote_first_row_to_header_if_needed(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty or not _headers_look_invalid(df):
        df.columns = _normalize_and_deduplicate_headers(list(df.columns))
        return df

    header_row_idx = _detect_header_row_index(df)
    if header_row_idx is None:
        df.columns = _normalize_and_deduplicate_headers(list(df.columns))
        return df

    candidate_headers = _normalize_and_deduplicate_headers(df.iloc[header_row_idx].tolist())
    cleaned_df = df.iloc[header_row_idx + 1 :].reset_index(drop=True).copy()
    cleaned_df.columns = candidate_headers
    return cleaned_df


def _find_column(df: pd.DataFrame, aliases: list[str]) -> str | None:
    # Preserve alias priority order: first alias has highest priority.
    canonical_to_columns: dict[str, list[str]] = {}
    for column in df.columns:
        key = _canonical_name(column)
        canonical_to_columns.setdefault(key, []).append(str(column))

    for alias in aliases:
        matched = canonical_to_columns.get(_canonical_name(alias))
        if matched:
            return matched[0]
    return None


def _map_origin_postal_code(origin_location: object) -> object:
    if pd.isna(origin_location):
        return pd.NA

    location = str(origin_location).strip()
    location_key = _canonical_name(location)
    if not location_key:
        return pd.NA

    postal_by_location_key = {
        _canonical_name("Basauri"): "Basauri",
        _canonical_name("BAC U TACHOVA"): "34802",
        _canonical_name("BOR U TACHOVA"): "34802",
        _canonical_name("Burgos"): "09007",
        _canonical_name("Igorre"): "48140",
        _canonical_name("Modugno"): "70026",
        _canonical_name("Paczkowo"): "62021",
        _canonical_name("Poznan"): "Poznan",
        _canonical_name("Puente San Miguel"): "39380",
        _canonical_name("Rome"): "001",
        _canonical_name("Stargard"): "Stargard/Goleniow",
        _canonical_name("Starsgar"): "Starsgar/Goleniow",
        _canonical_name("Tatabanya"): "2851",
        _canonical_name("Villalonquejar"): "09006",
        _canonical_name("Zeebrugge"): "8380",
        _canonical_name("Rivabellosa"): "01213",
    }

    # Requested special handling: Bitonto (not Modugno).
    if "bitonto" in location_key and "modugno" not in location_key:
        return "Bitonto"

    return postal_by_location_key.get(location_key, pd.NA)


def _apply_origin_postal_disambiguation(df: pd.DataFrame) -> pd.DataFrame:
    required_columns = {
        "ORIGIN LOCATION",
        "Origin Postal Code",
        "Destination City",
        "Equipment type",
    }
    if not required_columns.issubset(df.columns):
        return df

    def _apply_pair_rule(
        left_origin: str,
        right_origin: str,
        special_value: str,
    ) -> None:
        left_key = _canonical_name(left_origin)
        right_key = _canonical_name(right_origin)

        for _, lane_group in df.groupby(["Destination City", "Equipment type"], dropna=False):
            if lane_group.empty:
                continue

            group_origin_keys = lane_group["ORIGIN LOCATION"].apply(_canonical_name)
            has_left = (group_origin_keys == left_key).any()
            has_right = (group_origin_keys == right_key).any()
            if not (has_left and has_right):
                continue

            left_rows_mask = (
                (df["Destination City"] == lane_group["Destination City"].iloc[0])
                & (df["Equipment type"] == lane_group["Equipment type"].iloc[0])
                & (df["ORIGIN LOCATION"].apply(_canonical_name) == left_key)
            )
            df.loc[left_rows_mask, "Origin Postal Code"] = special_value

    _apply_pair_rule(
        left_origin="Basauri",
        right_origin="Igorre",
        special_value="Basauri (not Igorre)",
    )
    _apply_pair_rule(
        left_origin="Bitonto",
        right_origin="Modugno",
        special_value="Bitonto (not Modugno)",
    )

    return df


def _last_column_with_values(df: pd.DataFrame) -> str | None:
    for column in reversed(df.columns):
        series = df[column]
        has_value = series.apply(
            lambda value: pd.notna(value) and str(value).strip() and str(value).strip().lower() != "nan"
        ).any()
        if has_value:
            return str(column)
    return None


def _select_and_rename_columns(df: pd.DataFrame) -> pd.DataFrame:
    selected = pd.DataFrame(index=df.index)

    column_specs = [
        ("Carrier", ["Supplier Name", "Supplier"]),
        ("Carrier Code", ["Carrier Code"]),
        ("Origin Country", ["ORIGIN Country", "Origin Country"]),
        ("ORIGIN LOCATION", ["ORIGIN LOCATION", "Origin Location"]),
        ("Destination City", ["DESTINATION", "Destination", "Destination City"]),
        ("Destination Country", ["DESTINATON COUNTRY", "DESTINATION COUNTRY", "Destination Country"]),
        ("Equipment type", ["Equipment Type Offered", "Equipment Type"]),
        ("Destination postal code", ["DESTINATION POSTAL CODE", "Destination Postal Code"]),
        ("Bid Currency", ["Bid Currency", "Currency"]),
    ]

    for target_name, aliases in column_specs:
        source_column = _find_column(df, aliases)
        if source_column is None:
            selected[target_name] = pd.NA
        else:
            selected[target_name] = df[source_column]

    # User requested new empty helper columns for downstream lane mapping.
    selected.insert(3, "Origin Postal Code", pd.NA)
    selected.insert(6, "Destination city to differentiate lanes", pd.NA)
    selected["Origin Postal Code"] = selected["ORIGIN LOCATION"].apply(_map_origin_postal_code)
    selected = _apply_origin_postal_disambiguation(selected)

    dynamic_last_column = _last_column_with_values(df)
    if dynamic_last_column is not None and dynamic_last_column not in selected.columns:
        selected[dynamic_last_column] = df[dynamic_last_column]
    else:
        selected["Last Value Column"] = pd.NA

    return selected


def list_input_files() -> list[Path]:
    if not INPUT_DIR.exists():
        return []
    return sorted(
        [p for p in INPUT_DIR.iterdir() if p.is_file() and p.suffix.lower() in SUPPORTED_EXTENSIONS],
        key=lambda p: p.name.lower(),
    )


def choose_file(files: list[Path]) -> Path:
    print("Files available in input folder:\n")
    for idx, file_path in enumerate(files, start=1):
        print(f"{idx}. {file_path.name}")

    while True:
        user_input = input("\nChoose file number to convert: ").strip()
        if not user_input.isdigit():
            print("Please enter a valid number.")
            continue
        selected_index = int(user_input)
        if 1 <= selected_index <= len(files):
            return files[selected_index - 1]
        print("Number is out of range. Try again.")


def load_dataframe(file_path: Path) -> pd.DataFrame:
    if file_path.suffix.lower() == ".csv":
        return pd.read_csv(file_path)

    workbook = pd.ExcelFile(file_path)
    sheet_names = workbook.sheet_names

    if len(sheet_names) == 1:
        return pd.read_excel(file_path, sheet_name=sheet_names[0])

    preferred_sheet = next((s for s in sheet_names if s.strip().lower() == "rate table"), None)
    if preferred_sheet:
        print(f"Multiple tabs found. Using preferred tab '{preferred_sheet}'.")
        return pd.read_excel(file_path, sheet_name=preferred_sheet)

    df_sheet = next((s for s in sheet_names if s.strip().lower() == "df"), None)
    if df_sheet:
        print(f"Multiple tabs found. Using fallback tab '{df_sheet}'.")
        return pd.read_excel(file_path, sheet_name=df_sheet)

    print(
        "Multiple tabs found, but no tab named 'DF'. "
        "Using the first tab: "
        f"'{sheet_names[0]}'."
    )
    return pd.read_excel(file_path, sheet_name=sheet_names[0])


def save_to_processing(df: pd.DataFrame, source_file: Path) -> Path:
    PROCESSING_DIR.mkdir(parents=True, exist_ok=True)
    output_file = PROCESSING_DIR / f"{source_file.stem}_df.xlsx"
    df.to_excel(output_file, index=False)
    return output_file


def main() -> int:
    files = list_input_files()
    if not files:
        print("No supported files found in input folder.")
        return 1

    selected_file = choose_file(files)
    print(f"\nReading '{selected_file.name}'...")
    df = load_dataframe(selected_file)
    df = _promote_first_row_to_header_if_needed(df)
    df = _select_and_rename_columns(df)
    output_path = save_to_processing(df, selected_file)
    print(f"Done. Saved DataFrame to '{output_path}'.")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except KeyboardInterrupt:
        print("\nOperation cancelled.")
        raise SystemExit(130)
    except Exception as exc:
        print(f"Error: {exc}")
        raise SystemExit(1)
