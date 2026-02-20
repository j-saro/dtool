import os
import io
import zipfile
from lxml import etree
from docx import Document
from docxcompose.composer import Composer
import logging
from typing import List, IO

from ..utils import constants


def preprocess_docx(
    docx_bytes: bytes,
    remove_enc: bool = False,
    remove_tag: str = None,
    tag_ns: dict = None,
) -> io.BytesIO:
    input_stream = io.BytesIO(docx_bytes)
    output_stream = io.BytesIO()

    current_ns = constants.NS.copy()
    if tag_ns:
        current_ns.update(tag_ns)

    with zipfile.ZipFile(input_stream, "r") as z_in, zipfile.ZipFile(
        output_stream, "w", zipfile.ZIP_DEFLATED
    ) as z_out:

        for filename in z_in.namelist():
            content = z_in.read(filename)
            if remove_enc and filename == "word/settings.xml":
                try:
                    root = etree.fromstring(content)
                    protections = root.findall(
                        "w:documentProtection", namespaces=constants.NS
                    )
                    if protections:
                        for p in protections:
                            root.remove(p)
                        content = etree.tostring(
                            root, encoding="utf-8", xml_declaration=True
                        )
                        logging.info("Encryption removed.")
                except:
                    pass

            if remove_tag and filename == "word/document.xml":
                try:
                    root = etree.fromstring(content)
                    if ":" in remove_tag:
                        prefix, local = remove_tag.split(":")
                        xpath = f"//w:p[@{prefix}:{local}]"
                    else:
                        xpath = f"//w:p[@{remove_tag}]"

                    elements = root.xpath(xpath, namespaces=current_ns)
                    if elements:
                        for el in elements:
                            el.getparent().remove(el)
                        content = etree.tostring(
                            root, encoding="utf-8", xml_declaration=True
                        )
                        logging.info(f"Removed {len(elements)} tags.")
                except:
                    pass

            z_out.writestr(filename, content)

    output_stream.seek(0)
    return output_stream


def get_output_filename(input_file_path: str) -> str:
    base_name = os.path.basename(input_file_path)
    parts = base_name.split("_")
    if len(parts) > 1 and parts[0].isdigit():
        return "_".join(parts[1:])
    return base_name


def merge_documents(streams: List[IO[bytes]], output_path: str) -> None:
    if not streams:
        return

    logging.info(f"Initializing master document 1...")
    master_doc = Document(streams[0])
    composer = Composer(master_doc)

    for i, stream in enumerate(streams[1:]):
        logging.info(f"Appending document {i+2}...")
        doc_to_append = Document(stream)
        composer.append(doc_to_append)

    logging.info(f"Saving merged document to {output_path}")
    composer.save(output_path)
