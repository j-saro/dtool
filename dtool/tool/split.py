import os
import shutil
import time
from concurrent.futures import ThreadPoolExecutor
from lxml import etree
from copy import deepcopy
import logging

from ..core.config import Config
from ..core.processor import (
    parse_docx,
    parse_file,
    create_zip_from_folder,
)
from ..core.split_core import (
    preprocess_content,
    remove_by_index,
    remove_empty_wp_before,
    remove_empty_wp_after,
    remove_unreferenced_images,
)
from ..core.workspace import temporary_workspace, Workspace
from ..utils import constants
from ..utils.helpers import (
    generate_filename,
    move_files,
    timer,
)

logging.basicConfig(
    level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s"
)


@timer
def split_docx(config: Config) -> int:
    errors = []
    return_value = 0
    try:
        logging.info(f"Splitting {os.path.basename(config.file_path)}")
        basename = os.path.splitext(os.path.basename(config.file_path))[0]

        with temporary_workspace() as workspace:
            shared_extract_path = os.path.join(
                workspace.container_path, "shared_extract"
            )
            parse_docx(config.file_path, shared_extract_path)

            _, xml_root = parse_file(
                os.path.join(shared_extract_path, "word/document.xml")
            )
            _, rel_root = parse_file(
                os.path.join(shared_extract_path, "word/_rels/document.xml.rels")
            )

            logging.info(
                f"Splitting by {config.unit} (target: {config.count}) using {config.boundary} boundary"
            )
            modified_xml_root, index_list = preprocess_content(
                xml_root, config.count, config.unit, config.boundary
            )

            num_copies = len(index_list)
            if num_copies <= 1:
                logging.info(
                    "Document is smaller than target or no valid split points found. No splitting needed.",
                )
                if num_copies == 0:
                    return 0
                shutil.copy(config.file_path, config.output_path)
                return 1

            return_value = num_copies
            logging.info(f"Determined to create {num_copies} files.")

            modified_xml_data = etree.tostring(modified_xml_root, encoding="utf-8")

            if len(index_list) != num_copies:
                index_list.append((index_list[-1][1], constants.MAX_INDEX))
                logging.info("Adjusted index list")

            with ThreadPoolExecutor(
                max_workers=min(constants.MAX_WORKERS, num_copies)
            ) as executor:
                futures = []
                for i in range(num_copies):
                    futures.append(
                        executor.submit(
                            process_split,
                            i,
                            basename,
                            shared_extract_path,
                            modified_xml_data,
                            deepcopy(rel_root),
                            workspace,
                            index_list[i],
                            num_copies,
                        )
                    )

                for future in futures:
                    try:
                        future.result()
                    except Exception as e:
                        errors.append(str(e))

            if errors:
                logging.error(
                    f"Split completed with {len(errors)} errors:\n" + "\n".join(errors)
                )
                return 0

            logging.info(f"Output path: {config.output_path}")
            move_files(workspace.temp_copy_path, config.output_path)

        return return_value

    except Exception as e:
        logging.error(f"Split failed: {str(e)}")
        return 0


def process_split(
    index: int,
    basename: str,
    shared_extract_path: str,
    xml_data: bytes,
    rel_root: etree._Element,
    workspace: Workspace,
    index_range: tuple,
    total_splits: int,
) -> None:
    logging.info(f"Processing file {index+1}...")
    try:
        if not os.path.exists(shared_extract_path):
            raise FileNotFoundError(
                f"Shared extraction path missing: {shared_extract_path}"
            )

        split_work_dir = os.path.join(
            workspace.temp_dir_path, f"split_{index}_{time.time_ns()}"
        )
        try:
            shutil.copytree(shared_extract_path, split_work_dir)
        except shutil.Error as e:
            raise RuntimeError(f"File collision during copy: {str(e)}") from e

        xml_parser = etree.XMLParser(remove_blank_text=True)
        xml_root = etree.fromstring(xml_data, xml_parser)
        xml_tree = etree.ElementTree(xml_root)

        remove_by_index(
            xml_root,
            start=index_range[0],
            end=index_range[1],
        )

        if index != 0:
            remove_empty_wp_before(xml_root)
        if index != total_splits - 1:
            remove_empty_wp_after(xml_root)

        xml_root = remove_unreferenced_images(split_work_dir, xml_root, rel_root)

        # Save modified XML
        xml_path = os.path.join(split_work_dir, "word/document.xml")
        with open(xml_path, "wb") as f:
            xml_tree.write(
                f, encoding="utf-8", xml_declaration=True, pretty_print=False
            )

        # Save relationships
        rel_path = os.path.join(split_work_dir, "word/_rels/document.xml.rels")
        with open(rel_path, "wb") as f:
            f.write(etree.tostring(rel_root, encoding="utf-8", pretty_print=False))

        # Create final DOCX
        output_path = os.path.join(
            workspace.temp_copy_path, generate_filename(basename, index)
        )
        create_zip_from_folder(split_work_dir, output_path)

    except Exception as e:
        logging.error(f"Split {index+1} failed: {str(e)}", exc_info=e)
        if os.path.exists(split_work_dir):
            shutil.rmtree(split_work_dir, ignore_errors=True)
        raise
