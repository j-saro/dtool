import os
import zipfile
from lxml import etree
from typing import Tuple


def parse_docx(file_path: str, extraction_folder: str) -> None:
    os.makedirs(extraction_folder, exist_ok=True)

    if not os.path.isfile(file_path):
        raise FileNotFoundError(f"Input file not found: {file_path}")

    if not zipfile.is_zipfile(file_path):
        raise zipfile.BadZipFile(f"Invalid DOCX format: {file_path}")

    try:
        with zipfile.ZipFile(file_path, "r") as zip_ref:
            zip_ref.extractall(extraction_folder)
    except (zipfile.BadZipFile, zipfile.LargeZipFile) as e:
        raise RuntimeError(f"Corrupted/invalid DOCX file: {file_path}") from e
    except PermissionError as e:
        raise RuntimeError(f"Permission denied extracting {file_path}") from e


def parse_file(file_path: str) -> Tuple[str, etree._Element, bytes]:
    parser = etree.XMLParser(remove_blank_text=True, ns_clean=True)
    tree = etree.parse(file_path, parser)
    root = tree.getroot()
    data = etree.tostring(root, encoding="utf-8")

    return file_path, tree, root, data


def create_zip_from_folder(folder_path: str, output_path: str) -> None:
    with zipfile.ZipFile(
        output_path, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=5
    ) as zf:
        for root, _, files in os.walk(folder_path):
            for file in files:
                file_path = os.path.join(root, file)
                arcname = os.path.relpath(file_path, folder_path)

                # Skip compression for binaries
                compress_type = zipfile.ZIP_DEFLATED
                if file_path.lower().endswith((".png", ".jpg", ".jpeg")):
                    compress_type = zipfile.ZIP_STORED

                zf.write(file_path, arcname=arcname, compress_type=compress_type)
