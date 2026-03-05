import os
import logging
from typing import Union, List

from ..utils.helpers import timer
from ..core.merge_core import (
    preprocess_docx,
    get_output_filename,
    merge_documents,
)


@timer
def merge_docx(
    input_files: Union[List[str], str],
    output_path: str = os.getcwd(),
    remove_encryption=False,
) -> bool:

    processed_file_paths = []

    if isinstance(input_files, str):
        if os.path.isdir(input_files):
            processed_file_paths = [
                os.path.join(input_files, f)
                for f in os.listdir(input_files)
                if f.lower().endswith(".docx") and not f.startswith("~")
            ]
            processed_file_paths.sort()
        else:
            raise NotADirectoryError(f"path '{input_files}' is not a valid directory.")
    elif isinstance(input_files, list):
        processed_file_paths = input_files
    else:
        raise TypeError("Input must be folder path or list of files.")

    if not processed_file_paths:
        raise ValueError("No input files provided")

    try:
        logging.info("Starting Merge...")

        doc_streams = []

        for file_path in processed_file_paths:
            try:
                with open(file_path, "rb") as f:
                    file_bytes = f.read()

                stream = preprocess_docx(
                    file_bytes,
                    remove_encryption,
                    # TODO
                    # "deepml:banner",
                    # {"deepml": "http://www.deepl.com/document-translation/deepml"},
                )

                doc_streams.append(stream)

            except Exception as e:
                logging.error(f"Failed to load/process {file_path}: {e}")
                raise

        final_filename = get_output_filename(processed_file_paths[0])
        full_output_path = os.path.join(output_path, final_filename)

        os.makedirs(output_path, exist_ok=True)

        merge_documents(doc_streams, full_output_path)

        for s in doc_streams:
            s.close()

        return True

    except Exception as e:
        logging.error(f"Critical merging failure: {str(e)}", exc_info=True)
        return False
