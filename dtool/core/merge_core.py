import os
from lxml import etree
from docx import Document
from docxcompose.composer import Composer
import logging

from .processor import parse_file
from .workspace import Workspace
from ..utils.constants import NS


def remove_elements_with_tag(
    workspace: Workspace, tag: str, namespace: dict = None
) -> None:
    if namespace:
        NS.update(namespace)

    xml_path, xml_tree, xml_root, _ = parse_file(
        os.path.join(workspace.temp_dir_path, "word", "document.xml")
    )

    if ":" in tag:
        prefix, local_name = tag.split(":")
        xpath = f"//w:p[@{prefix}:{local_name}]"
    else:
        xpath = f"//w:p[@{tag}]"

    try:
        elements_to_remove = xml_root.xpath(xpath, namespaces=NS)
        for element in elements_to_remove:
            parent = element.getparent()
            if parent is not None:
                parent.remove(element)

        xml_tree.write(
            xml_path, encoding="utf-8", pretty_print=True, xml_declaration=True
        )
    except etree.XPathEvalError as e:
        raise ValueError(f"Failed XPath evaluation: {e}")


def remove_encryption_from_settings(workspace: Workspace) -> None:
    settings_path = os.path.join(workspace.temp_dir_path, "word", "settings.xml")

    if not os.path.exists(settings_path):
        logging.info("No settings.xml found - skipping encryption removal")
        return

    try:
        _, rel_tree, rel_root, _ = parse_file(settings_path)
        protections = rel_root.findall("w:documentProtection", namespaces=NS)

        if not protections:
            logging.info("No encryption tags found - nothing to remove")
            return

        for protection in protections:
            rel_root.remove(protection)

        rel_tree.write(
            settings_path, xml_declaration=True, encoding="utf-8", pretty_print=True
        )
        logging.info("Successfully removed document encryption")

    except etree.XMLSyntaxError as e:
        logging.error(
            f"Invalid XML structure in settings.xml: {str(e)}", exc_info=False
        )
    except Exception as e:
        logging.error(
            f"Encountered issue during encryption removal: {str(e)}", exc_info=False
        )


def get_output_path(output_path: str, input_file: list) -> str:
    base_path = os.path.basename(input_file)
    basename_parts = (
        base_path.split("_")[1:] if base_path.split("_")[0].isdigit() else []
    )
    basename = "_".join(basename_parts) or os.path.basename(input_file)

    if not os.path.exists(output_path):
        os.makedirs(output_path)

    output_docx = os.path.join(output_path, basename)
    if not output_docx.endswith(".docx"):
        output_docx += ".docx"
    return output_docx


def merge_documents(source_files: list, output_file_path: str) -> None:
    # Create initial document
    master_doc = Document(source_files[0])
    composer = Composer(master_doc)

    # Add page breaks and subsequent documents
    for index, file in enumerate(source_files[1:]):
        logging.info(f"Processing file {index+1}...")
        break_doc = Document()
        composer.append(break_doc)

        # Add document content
        composer.append(Document(file))

    # Save merged document
    composer.save(output_file_path)
