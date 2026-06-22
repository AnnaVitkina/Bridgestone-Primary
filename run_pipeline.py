from __future__ import annotations

import argparse
from pathlib import Path
import sys

from convert_input_to_processing import (
    _promote_first_row_to_header_if_needed,
    _select_and_rename_columns,
    load_dataframe as load_input_dataframe,
    save_to_processing,
)
from format_extracted_to_output import save_formatted_output

script_dir_candidates = []
if "__file__" in globals():
    script_dir_candidates.append(Path(__file__).resolve().parent)
script_dir_candidates.extend([Path("/content/Bridgestone-Primary"), Path.cwd()])
for candidate in script_dir_candidates:
    if candidate.exists() and str(candidate) not in sys.path:
        sys.path.insert(0, str(candidate))

BASE_DIR = Path(__file__).resolve().parent
DEFAULT_CODE_ROOT = Path("/content/Bridgestone-Primary")
DEFAULT_DATA_ROOT = Path(
    "/content/drive/Shareddrives/FA Ops Europe: Rate Maintenance Team/"
    "Documents/AI Adoption RMT/RMT_Bridgestone/Primary"
)
DEFAULT_INPUT_DIR = DEFAULT_DATA_ROOT / "input"
DEFAULT_PROCESSING_DIR = DEFAULT_DATA_ROOT / "processing"
DEFAULT_OUTPUT_DIR = DEFAULT_DATA_ROOT / "output"
SUPPORTED_EXTENSIONS = {".xlsx", ".xls", ".xlsm", ".xlsb", ".csv"}


def _list_input_files(input_dir: Path) -> list[Path]:
    if not input_dir.exists():
        return []
    return sorted(
        [
            p
            for p in input_dir.iterdir()
            if p.is_file()
            and p.suffix.lower() in SUPPORTED_EXTENSIONS
            and not p.name.startswith("~$")
        ],
        key=lambda p: p.name.lower(),
    )


def _choose_input_files(files: list[Path]) -> list[Path]:
    print("Files available in input folder:\n")
    for idx, file_path in enumerate(files, start=1):
        print(f"{idx}. {file_path.name}")

    print("\nChoose file number(s) to process.")
    print("Examples: 1      or      1,3      or      all")

    while True:
        user_input = input("Your selection: ").strip().lower()
        if user_input == "all":
            return files

        parts = [part.strip() for part in user_input.split(",") if part.strip()]
        if not parts:
            print("Please enter at least one file number or 'all'.")
            continue
        if not all(part.isdigit() for part in parts):
            print("Use numbers separated by commas, or 'all'.")
            continue

        indices = sorted(set(int(part) for part in parts))
        if not all(1 <= idx <= len(files) for idx in indices):
            print("One or more numbers are out of range.")
            continue
        return [files[idx - 1] for idx in indices]


def _process_single_file(source_file: Path) -> tuple[Path, Path]:
    extracted_df = load_input_dataframe(source_file)
    extracted_df = _promote_first_row_to_header_if_needed(extracted_df)
    extracted_df = _select_and_rename_columns(extracted_df)

    extracted_output = save_to_processing(extracted_df, source_file)
    layout_output = save_formatted_output(extracted_df, extracted_output)
    return extracted_output, layout_output


def run_pipeline(
    input_dir: Path,
    processing_dir: Path,
    output_dir: Path,
) -> None:
    files = _list_input_files(input_dir)
    if not files:
        raise FileNotFoundError(f"No supported files found in '{input_dir}'.")

    processing_dir.mkdir(parents=True, exist_ok=True)
    output_dir.mkdir(parents=True, exist_ok=True)

    selected_files = _choose_input_files(files)
    for source_file in selected_files:
        print(f"\nProcessing: {source_file.name}")
        extracted_path, output_path = _process_single_file(source_file)
        print(f"Extracted saved to: {extracted_path}")
        print(f"Formatted output saved to: {output_path}")

    print("\nPipeline completed.")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Run Primary end-to-end pipeline (input -> processing -> output)."
    )
    parser.add_argument(
        "--base-dir",
        type=Path,
        default=DEFAULT_CODE_ROOT,
        help="Base directory containing input/processing/output folders.",
    )
    parser.add_argument("--input-dir", type=Path, default=None, help="Override input directory.")
    parser.add_argument(
        "--processing-dir",
        type=Path,
        default=None,
        help="Override processing directory.",
    )
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=None,
        help="Override output directory.",
    )
    return parser.parse_args()


if __name__ == "__main__":
    args = parse_args()
    default_input = DEFAULT_INPUT_DIR if DEFAULT_INPUT_DIR.exists() else BASE_DIR / "input"
    default_processing = (
        DEFAULT_PROCESSING_DIR if DEFAULT_PROCESSING_DIR.exists() else BASE_DIR / "processing"
    )
    default_output = DEFAULT_OUTPUT_DIR if DEFAULT_OUTPUT_DIR.exists() else BASE_DIR / "output"

    run_pipeline(
        input_dir=args.input_dir or default_input,
        processing_dir=args.processing_dir or default_processing,
        output_dir=args.output_dir or default_output,
    )
