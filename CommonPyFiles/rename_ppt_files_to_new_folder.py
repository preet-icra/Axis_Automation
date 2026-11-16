import argparse
import logging
import re
import shutil
from pathlib import Path


POWERPOINT_EXTENSIONS = {".ppt", ".pptx"}
DEFAULT_OUTPUT_FOLDER = "renamed_ppt_files"


def build_new_filename(file_path: Path) -> str:
    """Remove all case-insensitive occurrences of '_updated' from the filename."""
    new_stem = re.sub(r"_updated", "", file_path.stem, flags=re.IGNORECASE)
    return f"{new_stem}{file_path.suffix}"


def save_renamed_powerpoint_files(folder_path: str, output_name: str = DEFAULT_OUTPUT_FOLDER) -> dict:
    """Copy renamed PowerPoint files into a new subfolder, preserving originals."""
    source_folder = Path(folder_path).expanduser().resolve()

    if not source_folder.exists():
        raise FileNotFoundError(f"Folder does not exist: {source_folder}")
    if not source_folder.is_dir():
        raise NotADirectoryError(f"Path is not a folder: {source_folder}")

    output_folder = source_folder / output_name
    output_folder.mkdir(parents=True, exist_ok=True)

    summary = {
        "found": 0,
        "copied": 0,
        "skipped": 0,
        "failed": 0,
        "output_folder": str(output_folder),
    }

    for file_path in source_folder.iterdir():
        if not file_path.is_file():
            continue
        if file_path.suffix.lower() not in POWERPOINT_EXTENSIONS:
            continue
        if "_updated" not in file_path.stem.lower():
            continue

        summary["found"] += 1
        destination = output_folder / build_new_filename(file_path)

        if destination.exists():
            logging.warning("Skipped, destination already exists: %s", destination.name)
            summary["skipped"] += 1
            continue

        try:
            shutil.copy2(file_path, destination)
            logging.info("Saved: %s -> %s", file_path.name, destination)
            summary["copied"] += 1
        except (PermissionError, OSError) as error:
            logging.error("Could not copy %s: %s", file_path.name, error)
            summary["failed"] += 1

    return summary


def main() -> None:
    parser = argparse.ArgumentParser(
        description=(
            "Create a new folder and save renamed copies of .ppt and .pptx files "
            "after removing '_updated' from their filenames."
        )
    )
    parser.add_argument("folder", nargs="?", help="Folder containing PowerPoint files")
    parser.add_argument(
        "--output-folder",
        default=DEFAULT_OUTPUT_FOLDER,
        help=f"Name of the new output folder. Default: {DEFAULT_OUTPUT_FOLDER}",
    )
    args = parser.parse_args()

    folder_path = args.folder or input("Enter the folder path: ").strip().strip('"')
    logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")

    try:
        result = save_renamed_powerpoint_files(folder_path, args.output_folder)
        print("\nSummary")
        print(f"Matching files found : {result['found']}")
        print(f"Renamed copies saved : {result['copied']}")
        print(f"Files skipped        : {result['skipped']}")
        print(f"Failures             : {result['failed']}")
        print(f"Output folder        : {result['output_folder']}")
    except (FileNotFoundError, NotADirectoryError) as error:
        logging.error(error)


if __name__ == "__main__":
    main()
