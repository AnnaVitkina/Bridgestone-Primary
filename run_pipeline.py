from __future__ import annotations

import argparse
from pathlib import Path
import sys

script_dir_candidates = []
if "__file__" in globals():
    script_file = Path(__file__).resolve()
else:
    fallback_script = Path("/content/Bridgestone-Primary/run_pipeline.py")
    script_file = fallback_script if fallback_script.exists() else Path.cwd() / "run_pipeline.py"

if "__file__" in globals():
    script_dir_candidates.append(Path(__file__).resolve().parent)
script_dir_candidates.extend([script_file.parent, Path("/content/Bridgestone-Primary"), Path.cwd()])
for candidate in script_dir_candidates:
    if candidate.exists() and str(candidate) not in sys.path:
        sys.path.insert(0, str(candidate))

import convert_input_to_processing as convert_module
import format_extracted_to_output as format_module
from convert_input_to_processing import (
    _promote_first_row_to_header_if_needed,
    _select_and_rename_columns,
    load_dataframe as load_input_dataframe,
    save_to_processing,
)
from format_extracted_to_output import save_formatted_output

BASE_DIR = script_file.parent
DEFAULT_CODE_ROOT = Path("/content/Bridgestone-Primary")
DEFAULT_DATA_ROOT = Path(
    "/content/drive/Shareddrives/FA Ops Europe: Rate Maintenance Team /"
    "Documents/AI Adoption RMT/RMT_Bridgestone/Primary"
)
DEFAULT_INPUT_DIR = DEFAULT_DATA_ROOT / "input"
DEFAULT_PROCESSING_DIR = DEFAULT_DATA_ROOT / "processing"
DEFAULT_OUTPUT_DIR = DEFAULT_DATA_ROOT / "output"
LOCAL_INPUT_DIR = BASE_DIR / "input"
LOCAL_PROCESSING_DIR = BASE_DIR / "processing"
LOCAL_OUTPUT_DIR = BASE_DIR / "output"
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
    # Force downstream modules to save into the pipeline target folders.
    convert_module.PROCESSING_DIR = processing_dir
    format_module.OUTPUT_DIR = output_dir

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
    parser.add_argument(
        "--local",
        action="store_true",
        help="Force local folders (next to this script) instead of Google Drive paths.",
    )
    args, _unknown = parser.parse_known_args()
    return args


def _drive_paths_available() -> bool:
    return DEFAULT_DATA_ROOT.exists() and DEFAULT_INPUT_DIR.exists()


def _local_paths(base_dir: Path | None = None) -> tuple[Path, Path, Path]:
    root = base_dir or BASE_DIR
    return root / "input", root / "processing", root / "output"


def _resolve_directories(args: argparse.Namespace) -> tuple[Path, Path, Path, str]:
    # Explicit overrides always win.
    if args.input_dir or args.processing_dir or args.output_dir:
        local_base = args.base_dir if args.base_dir != DEFAULT_CODE_ROOT else BASE_DIR
        local_input, local_processing, local_output = _local_paths(local_base)
        return (
            args.input_dir or local_input,
            args.processing_dir or local_processing,
            args.output_dir or local_output,
            "custom",
        )

    if args.local or not _drive_paths_available():
        local_base = args.base_dir if args.base_dir != DEFAULT_CODE_ROOT else BASE_DIR
        input_dir, processing_dir, output_dir = _local_paths(local_base)
        if args.local:
            mode = "local (forced)"
        else:
            mode = "local (Google Drive path not found)"
        return input_dir, processing_dir, output_dir, mode

    return DEFAULT_INPUT_DIR, DEFAULT_PROCESSING_DIR, DEFAULT_OUTPUT_DIR, "google-drive"


if __name__ == "__main__":
    args = parse_args()
    input_dir, processing_dir, output_dir, mode = _resolve_directories(args)
    print(f"Running mode: {mode}")
    print(f"Using input dir: {input_dir}")
    print(f"Using processing dir: {processing_dir}")
    print(f"Using output dir: {output_dir}")

    run_pipeline(
        input_dir=input_dir,
        processing_dir=processing_dir,
        output_dir=output_dir,
    )
