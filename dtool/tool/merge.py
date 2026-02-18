import os
import logging
from typing import Union, List

from ..core.workspace import temporary_workspace
from ..core.processor import parse_docx, create_zip_from_folder
from ..utils.helpers import (
    generate_filename,
    timer,
    copy_files,
)
from ..core.merge_core import (
    remove_elements_with_tag,
    remove_encryption_from_settings,
    get_output_path,
    merge_documents,
)


@timer
def merge_docx(
    input_files: Union[List[str], str], output_path: str = os.getcwd(), remove_encryption=False
) -> bool:
    processed_file_list = []

    if isinstance(input_files, str):
        if os.path.isdir(input_files):
            processed_file_list = [
                os.path.join(input_files, f) 
                for f in os.listdir(input_files) 
                if f.lower().endswith(".docx") and not f.startswith("~")
            ]
            processed_file_list.sort()
        else:
            raise NotADirectoryError(f"The path '{input_files}' is not a valid directory.")
            
    elif isinstance(input_files, list):
        processed_file_list = input_files
    else:
        raise TypeError("Input must be either a folder path (str) or a list of files.")

    if not processed_file_list:
        raise ValueError("No input files provided for merging")

    if not all(os.path.isfile(f) for f in processed_file_list):
        missing = [f for f in processed_file_list if not os.path.exists(f)]
        raise FileNotFoundError(f"Missing input files: {missing}")

    try:
        logging.info("Merging Files...")
        with temporary_workspace() as workspace:
            try:
                files_to_merge = sorted([file for file in processed_file_list])
                basename = os.path.splitext(os.path.basename(processed_file_list[0]))[0]

                if remove_encryption:
                    for index, file in enumerate(files_to_merge):
                        try:
                            parse_docx(file, workspace.temp_dir_path)
                            # remove_elements_with_tag(
                            #    workspace,
                            #    "deepml:banner",
                            #    {
                            #        "deepml": "http://www.deepl.com/document-translation/deepml"
                            #    },
                            # )
                            remove_encryption_from_settings(workspace)

                            base_parts = os.path.basename(file).split("_")
                            if len(base_parts) < 2 or not base_parts[0].isdigit():
                                raise ValueError(f"Invalid filename format: {file}")

                            basename = base_parts[1]
                            file_name = generate_filename(basename, index)

                            create_zip_from_folder(
                                workspace.temp_dir_path,
                                os.path.join(workspace.temp_copy_path, file_name),
                            )
                        except Exception as e:
                            logging.error(
                                f"Failed to process file {file}: {str(e)}", exc_info=e
                            )
                            raise
                else:
                    copy_files(processed_file_list, workspace.temp_copy_path)

                files_in_dir = os.listdir(workspace.temp_copy_path)
                modified_files = sorted(
                    [
                        os.path.abspath(os.path.join(workspace.temp_copy_path, file))
                        for file in files_in_dir
                    ]
                )

                merge_documents(
                    modified_files, get_output_path(output_path, processed_file_list[0])
                )

            except Exception as e:
                logging.error(f"Merging aborted: {str(e)}", exc_info=e)
                workspace.cleanup()
                return False

        return True

    except Exception as e:
        logging.error(f"Critical merging failure: {str(e)}", exc_info=e)
        return False
