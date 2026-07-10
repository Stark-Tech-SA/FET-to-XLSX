"""XML helpers that tolerate namespaces and minor FET tag variants."""
from __future__ import annotations

import re
import xml.etree.ElementTree as ET


def local_name(tag: str) -> str:
    """Return an XML tag without namespace information."""
    return tag.rsplit("}", 1)[-1] if "}" in tag else tag


def normalized_name(tag: str) -> str:
    """Normalize a tag for case/underscore-insensitive comparisons."""
    return re.sub(r"[^a-z0-9]", "", local_name(tag).lower())


def child(element: ET.Element, *names: str) -> ET.Element | None:
    """Return the first direct child matching one of the provided tag names."""
    wanted = {normalized_name(name) for name in names}
    for item in list(element):
        if normalized_name(item.tag) in wanted:
            return item
    return None


def children(element: ET.Element, *names: str) -> list[ET.Element]:
    """Return direct children matching any provided tag name."""
    wanted = {normalized_name(name) for name in names}
    return [item for item in list(element) if normalized_name(item.tag) in wanted]


def descendant(element: ET.Element, *names: str) -> ET.Element | None:
    """Return the first descendant matching any provided tag name."""
    wanted = {normalized_name(name) for name in names}
    for item in element.iter():
        if item is not element and normalized_name(item.tag) in wanted:
            return item
    return None


def descendants(element: ET.Element, *names: str) -> list[ET.Element]:
    """Return all descendants matching any provided tag name."""
    wanted = {normalized_name(name) for name in names}
    return [item for item in element.iter() if item is not element and normalized_name(item.tag) in wanted]


def text(element: ET.Element | None, default: str = "") -> str:
    """Read stripped text from an XML element."""
    return (element.text or default).strip() if element is not None else default
