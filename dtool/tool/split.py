import os
import zipfile
from concurrent.futures import ThreadPoolExecutor
from lxml import etree
import logging

from ..core.config import Config
from ..core.split_core import (
    preprocess_content,
    remove_by_index,
    remove_empty_wp_before,
    remove_empty_wp_after,
    get_used_images,
)
from ..core.workspace import InMemoryDocx
from ..utils import constants
from ..utils.helpers import (
    generate_filename,
    timer,
)

logging.basicConfig(
    level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s"
)


@timer
def split_docx(config: Config) -> int:
    errors = []
    try:
        logging.info(f"Splitting {os.path.basename(config.file_path)}")
        basename = os.path.splitext(os.path.basename(config.file_path))[0]

        doc_context = InMemoryDocx(config.file_path)

        xml_bytes = doc_context.get_xml("word/document.xml")
        rel_bytes = doc_context.get_xml("word/_rels/document.xml.rels")
        xml_root = etree.fromstring(xml_bytes, etree.XMLParser(remove_blank_text=True))

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
            return 0

        base_xml_data = etree.tostring(modified_xml_root, encoding="utf-8")
        os.makedirs(config.output_path, exist_ok=True)

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
                        doc_context.assets,
                        base_xml_data,
                        rel_bytes,
                        index_list[i],
                        num_copies,
                        config.output_path,
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
        logging.info(f"Output path: {config.output_path}")
        return num_copies

    except Exception as e:
        logging.error(f"Split failed: {str(e)}")
        return 0


def process_split(
    index: int,
    basename: str,
    assets: dict,
    xml_data: bytes,
    rel_data: bytes,
    index_range: tuple,
    total_splits: int,
    output_folder: str,
) -> None:

    xml_root = etree.fromstring(xml_data)
    rel_root = etree.fromstring(rel_data)

    remove_by_index(xml_root, start=index_range[0], end=index_range[1])

    if index != 0:
        remove_empty_wp_before(xml_root)
    if index != total_splits - 1:
        remove_empty_wp_after(xml_root)

    needed_images = get_used_images(xml_root, rel_root, constants.NS)

    for rel in rel_root.findall(".//ns:Relationship", constants.NS):
        if "image" in rel.get("Type", ""):
            target = rel.get("Target")
            img_name = os.path.basename(target)
            if img_name not in needed_images:
                rel.getparent().remove(rel)

    final_filename = generate_filename(basename, index)
    final_path = os.path.join(output_folder, final_filename)

    with zipfile.ZipFile(final_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for filename, content in assets.items():
            if filename == "word/document.xml":
                continue
            if filename == "word/_rels/document.xml.rels":
                continue
            if filename.startswith("word/media/"):
                img_name = os.path.basename(filename)
                if img_name in needed_images:
                    zf.writestr(filename, content)
                continue

            zf.writestr(filename, content)

        zf.writestr(
            "word/document.xml",
            etree.tostring(xml_root, encoding="utf-8", xml_declaration=True),
        )
        zf.writestr(
            "word/_rels/document.xml.rels",
            etree.tostring(rel_root, encoding="utf-8", xml_declaration=True),
        )

    logging.info(f"File {index+1} written.")
