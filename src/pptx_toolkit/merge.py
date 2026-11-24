"""Merge several .pptx files into one, preserving slide order and content.

Strategy: the first input file becomes the base presentation (so its theme,
slide size and opening slides are kept exactly). Every slide of the remaining
files is then re-created in the base deck: a new slide is added using the
layout at the same position in the base deck's first slide master, the
default placeholder shapes are removed, and the source slide's shape XML is
deep-copied in. Relationships that the copied XML points at (pictures,
media, hyperlinks, ...) are re-created on the new slide part so embedded
images travel with the slide.

Known trade-off (see README "Limitations"): because copied slides adopt the
base deck's master and layouts, a source deck built on a very different theme
can render with the base deck's fonts/colors, and exotic features such as
animations or SmartArt may not survive the copy.
"""

from __future__ import annotations

import copy

from pptx import Presentation
from pptx.opc.constants import RELATIONSHIP_TYPE as RT

# XML attributes (in the relationships namespace) that can hold an rId.
_R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_R_ATTRS = (f"{{{_R_NS}}}embed", f"{{{_R_NS}}}link", f"{{{_R_NS}}}id")


def _layout_for(dest_prs: Presentation, src_slide) -> object:
    """Pick the destination layout matching the source layout's position."""
    src_layout = src_slide.slide_layout
    try:
        idx = list(src_layout.slide_master.slide_layouts).index(src_layout)
    except ValueError:
        idx = 0
    dest_layouts = dest_prs.slide_masters[0].slide_layouts
    return dest_layouts[min(idx, len(dest_layouts) - 1)]


def _copy_slide(dest_prs: Presentation, src_slide) -> None:
    dest_slide = dest_prs.slides.add_slide(_layout_for(dest_prs, src_slide))

    # Remove the empty placeholder shapes that come with the layout.
    for shape in list(dest_slide.shapes):
        element = shape._element
        element.getparent().remove(element)

    src_part = src_slide.part
    dest_part = dest_slide.part

    # Deep-copy the shape elements first (not yet attached), so we can scan
    # which relationships they actually reference.
    new_elements = [copy.deepcopy(shape._element) for shape in src_slide.shapes]

    # Also carry over a slide background if the source slide defines one.
    src_cSld = src_slide._element.find(
        "{http://schemas.openxmlformats.org/presentationml/2006/main}cSld"
    )
    bg_element = None
    if src_cSld is not None:
        src_bg = src_cSld.find(
            "{http://schemas.openxmlformats.org/presentationml/2006/main}bg"
        )
        if src_bg is not None:
            bg_element = copy.deepcopy(src_bg)
            new_elements.append(bg_element)

    # Collect referenced rIds from the copied XML.
    referenced: set[str] = set()
    for element in new_elements:
        for node in element.iter():
            for attr in _R_ATTRS:
                value = node.get(attr)
                if value:
                    referenced.add(value)

    # Re-create those relationships on the destination slide part and build
    # an old-rId -> new-rId map. The slide-layout relationship is skipped:
    # the new slide already has one, and copied shapes never reference it.
    rid_map: dict[str, str] = {}
    for rid in referenced:
        if rid not in src_part.rels:
            continue
        rel = src_part.rels[rid]
        if rel.reltype == RT.SLIDE_LAYOUT:
            continue
        if rel.is_external:
            new_rid = dest_part.relate_to(rel.target_ref, rel.reltype, is_external=True)
        else:
            new_rid = dest_part.relate_to(rel.target_part, rel.reltype)
        rid_map[rid] = new_rid

    # Rewrite the rIds inside the copied XML.
    for element in new_elements:
        for node in element.iter():
            for attr in _R_ATTRS:
                value = node.get(attr)
                if value in rid_map:
                    node.set(attr, rid_map[value])

    # Attach shapes to the slide's shape tree; background goes to <p:cSld>.
    spTree = dest_slide.shapes._spTree
    for element in new_elements:
        if element is bg_element:
            dest_cSld = dest_slide._element.find(
                "{http://schemas.openxmlformats.org/presentationml/2006/main}cSld"
            )
            dest_cSld.insert(0, element)
        else:
            spTree.append(element)


def merge_presentations(input_paths: list[str], output_path: str) -> int:
    """Merge the decks at *input_paths* (in order) into *output_path*.

    Returns the total number of slides in the merged deck.
    """
    if len(input_paths) < 2:
        raise ValueError("merge needs at least two input files")

    dest_prs = Presentation(input_paths[0])
    for path in input_paths[1:]:
        src_prs = Presentation(path)
        for src_slide in src_prs.slides:
            _copy_slide(dest_prs, src_slide)

    dest_prs.save(output_path)
    return len(dest_prs.slides._sldIdLst)
