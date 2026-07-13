"""Extract images from .docx and replace them with text labels."""

from __future__ import annotations

from pathlib import Path

from docx import Document
from docx.oxml.ns import qn
from lxml import etree

from app.services.image_utils import bytes_to_png, image_label

_A_BLIP = qn("a:blip")
_R_EMBED = qn("r:embed")
_R_LINK = qn("r:link")
_W_R = qn("w:r")
_W_T = qn("w:t")
_V_IMAGEDATA = "{urn:schemas-microsoft-com:vml}imagedata"
_R_ID_VML = qn("r:id")
_XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"


def _local(tag: str) -> str:
    return tag.rsplit("}", 1)[-1] if "}" in tag else tag


def _find_blip_rids(drawing_el) -> list[str]:
    rids: list[str] = []
    for blip in drawing_el.iter(_A_BLIP):
        rid = blip.get(_R_EMBED) or blip.get(_R_LINK)
        if rid:
            rids.append(rid)
    for imagedata in drawing_el.iter(_V_IMAGEDATA):
        rid = imagedata.get(_R_ID_VML) or imagedata.get(
            "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
        )
        if rid:
            rids.append(rid)
    return rids


def _make_text_run(text: str):
    r = etree.Element(_W_R)
    t = etree.SubElement(r, _W_T)
    t.text = text
    t.set(_XML_SPACE, "preserve")
    return r


def _nearest_paragraph(el):
    node = el
    while node is not None:
        if _local(node.tag) == "p":
            return node
        node = node.getparent()
    return None


def _replace_drawing_with_label(drawing_el, label: str) -> None:
    """Replace w:drawing / w:pict with text label img_XXX."""
    parent = drawing_el.getparent()
    if parent is None:
        return

    if _local(parent.tag) == "r":
        idx = list(parent).index(drawing_el)
        parent.remove(drawing_el)
        t = etree.Element(_W_T)
        t.text = label
        t.set(_XML_SPACE, "preserve")
        parent.insert(idx, t)
        return

    para = _nearest_paragraph(drawing_el)
    parent.remove(drawing_el)
    if para is not None:
        para.append(_make_text_run(label))


def _iter_story_parts(doc: Document):
    """Yield (oxml_root, related_parts_dict) for body + headers/footers."""
    yield doc.element.body, doc.part.related_parts

    for section in doc.sections:
        for part in (
            section.header,
            section.footer,
            section.first_page_header,
            section.first_page_footer,
            section.even_page_header,
            section.even_page_footer,
        ):
            try:
                yield part._element, part.part.related_parts
            except Exception:
                continue


def _resolve_image_blob(rid: str, related_parts: dict, doc: Document):
    if rid in related_parts:
        return related_parts[rid].blob
    # Fallback: search all story parts
    for _, parts in _iter_story_parts(doc):
        if rid in parts:
            return parts[rid].blob
    return None


def process_docx(source: Path, images_dir: Path, output_docx: Path) -> list[tuple[str, Path]]:
    """
    Extract images, save PNGs, replace with labels, save modified docx.

    Returns list of (label, png_path) in document order.
    """
    doc = Document(str(source))
    entries: list[tuple[str, Path]] = []
    counter = 0

    # Collect (element, related_parts) in document order
    candidates: list[tuple] = []
    seen_ids: set[int] = set()
    for root, related_parts in _iter_story_parts(doc):
        for el in root.iter():
            if _local(el.tag) not in ("drawing", "pict"):
                continue
            eid = id(el)
            if eid in seen_ids:
                continue
            seen_ids.add(eid)
            candidates.append((el, related_parts))

    for drawing_el, related_parts in candidates:
        if drawing_el.getparent() is None:
            continue

        rids = _find_blip_rids(drawing_el)
        if not rids:
            continue

        try:
            blob = _resolve_image_blob(rids[0], related_parts, doc)
            if blob is None:
                continue

            counter += 1
            label = image_label(counter)
            png_path = images_dir / f"{label}.png"
            try:
                bytes_to_png(blob, png_path)
            except Exception:
                counter -= 1
                continue

            _replace_drawing_with_label(drawing_el, label)
            entries.append((label, png_path))
        except Exception:
            continue

    output_docx.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(output_docx))
    return entries
