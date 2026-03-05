import os
from lxml import etree
from copy import copy
from typing import Tuple

from .config import Boundary, Unit
from ..utils.constants import NS


def _find_last_valid_split_offset(text: str, strategy: Boundary) -> int:
    if not text:
        return -1

    last_pos = -1

    if strategy == Boundary.NEAREST_SENTENCE:
        terminators = [". ", "! ", "? ", ".\n", "!\n", "?\n"]
        for term in terminators:
            pos = text.rfind(term)
            if pos > last_pos:
                last_pos = pos

        if text.endswith((".", "!", "?")) and len(text) - 1 > last_pos:
            last_pos = len(text) - 1

    elif strategy == Boundary.NEAREST_WORD:
        pos = text.rfind(" ")
        if pos > last_pos:
            last_pos = pos

    if last_pos != -1:
        return last_pos + 1

    return -1


def _split_paragraph_node(
    p_element: etree._Element,
    text_node_to_split: etree._Element,
    char_offset_in_node: int,
):
    new_p = copy(p_element)

    original_text_nodes = list(p_element.iterfind(".//w:t", namespaces=NS))
    split_node_index = original_text_nodes.index(text_node_to_split)

    original_split_node = original_text_nodes[split_node_index]
    original_split_node.text = original_split_node.text[:char_offset_in_node]

    nodes_to_remove = original_text_nodes[split_node_index + 1 :]
    for node in nodes_to_remove:
        parent_run = node.getparent()
        parent_run.remove(node)
        if not parent_run.getchildren() and parent_run.text is None:
            parent_run.getparent().remove(parent_run)

    new_text_nodes = list(new_p.iterfind(".//w:t", namespaces=NS))

    new_split_node = new_text_nodes[split_node_index]
    new_split_node.text = new_split_node.text[char_offset_in_node:]

    nodes_to_remove = new_text_nodes[:split_node_index]
    for node in nodes_to_remove:
        parent_run = node.getparent()
        parent_run.remove(node)
        if not parent_run.getchildren() and parent_run.text is None:
            parent_run.getparent().remove(parent_run)

    if not new_split_node.text:
        parent_run = new_split_node.getparent()
        parent_run.remove(new_split_node)

    p_element.addnext(new_p)


def preprocess_content(
    root: etree._Element, count: int, unit: Unit, boundary: Boundary
) -> Tuple[etree._Element, list]:

    all_text_nodes = root.xpath(".//w:t", namespaces=NS)
    if not all_text_nodes:
        return root, []

    planned_splits = []

    current_doc_val = 0
    node_idx = 0
    total_nodes = len(all_text_nodes)

    while node_idx < total_nodes:
        node = all_text_nodes[node_idx]
        node_text = node.text or ""

        if unit == Unit.CHARS:
            node_val = len(node_text)
        else:
            node_val = len(node_text.split())
        potential_split_offset = -1
        if boundary != Boundary.STRICT:
            potential_split_offset = _find_last_valid_split_offset(node_text, boundary)

        temp_val = current_doc_val + node_val

        if temp_val >= count:
            final_split_node = node
            final_offset = -1

            if boundary != Boundary.STRICT and potential_split_offset != -1:
                final_offset = potential_split_offset

            else:
                if unit == Unit.CHARS:
                    overhang = temp_val - count
                    final_offset = len(node_text) - overhang
                else:
                    final_offset = len(node_text)

            if final_offset > 0 and final_offset < len(node_text):
                planned_splits.append((final_split_node, final_offset))

                remainder_text = node_text[final_offset:]
                if unit == Unit.CHARS:
                    current_doc_val = len(remainder_text)
                else:
                    current_doc_val = len(remainder_text.split())
            else:
                current_doc_val = 0

        else:
            current_doc_val = temp_val

        node_idx += 1

    for node, offset in reversed(planned_splits):
        p_element = node.getparent()
        W_P = etree.QName(NS['w'], 'p').text
        while p_element.tag != W_P:
            p_element = p_element.getparent()

        if p_element is not None:
            _split_paragraph_node(p_element, node, offset)

    final_text_nodes = root.xpath(".//w:t", namespaces=NS)
    split_ranges = []
    current_start_idx = 0
    current_val = 0

    for idx, node in enumerate(final_text_nodes):
        txt = node.text or ""
        val = len(txt) if unit == Unit.CHARS else len(txt.split())
        current_val += val

        if current_val >= count:
            split_ranges.append((current_start_idx, idx + 1))
            current_start_idx = idx + 1
            current_val = 0

    if current_start_idx < len(final_text_nodes):
        split_ranges.append((current_start_idx, len(final_text_nodes)))

    return root, split_ranges


def remove_by_index(
    root: etree._Element,
    start: int,
    end: int,
) -> None:
    paragraphs_to_remove = root.xpath(".//w:t", namespaces=NS)
    to_remove = paragraphs_to_remove[:start] + paragraphs_to_remove[end:]
    for paragraph in to_remove:
        paragraph.getparent().remove(paragraph)


def remove_empty_wp_before(root: etree._Element) -> None:
    first_text_found = False
    for paragraph in root.findall(".//w:p", namespaces=NS):
        text_elements = paragraph.findall(".//w:t", namespaces=NS)
        if text_elements and not first_text_found:
            first_text_found = True
        elif not text_elements and not first_text_found:
            paragraph.getparent().remove(paragraph)


def remove_empty_wp_after(root: etree._Element) -> None:
    last_text_index = 0
    for index, paragraph in enumerate(root.findall(".//w:p", namespaces=NS)):
        text_elements = paragraph.findall(".//w:t", namespaces=NS)
        if text_elements:
            last_text_index = index

    for index, paragraph in enumerate(root.findall(".//w:p", namespaces=NS)):
        if index > last_text_index:
            paragraph.getparent().remove(paragraph)


def get_used_images(xml_root, rel_root, namespace) -> set:
    used_image_names = set()

    rel_id_to_target = {}
    for rel in rel_root.findall(".//ns:Relationship", namespace):
        if "image" in rel.get("Type", ""):
            rel_id_to_target[rel.get("Id")] = rel.get("Target")

    for blip in xml_root.findall(".//w:drawing//a:blip", namespace):
        embed_id = blip.get(f"{{{namespace['r']}}}embed")
        if embed_id in rel_id_to_target:
            target = rel_id_to_target[embed_id]
            used_image_names.add(os.path.basename(target))

    return used_image_names
