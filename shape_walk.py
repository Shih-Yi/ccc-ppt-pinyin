"""Walk every text-bearing shape on a slide, including shapes nested in groups.

`slide.shapes` only lists top-level shapes. A group (<p:grpSp>, what PowerPoint's
"Group" command produces) carries no text frame of its own, so a plain loop skips
it — and every text box inside it goes unprocessed, leaving the deck looking
completely untouched.

Groups also open their own coordinate space: a child's `width` is expressed in
child units, which map onto slide units by the group's ext/chExt ratio. The
walker therefore reports an *effective* width, so layout maths stays in real
slide units no matter how deeply a shape is nested.
"""
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"


def _a(tag):
    return f"{{{A_NS}}}{tag}"


def _p(tag):
    return f"{{{P_NS}}}{tag}"


GROUP_TAGS = frozenset({_p("grpSp")})


def _is_group(shape):
    # tag test rather than shape_type: python-pptx raises on a few exotic
    # shape types, and a group is unambiguous at the XML level
    return shape._element.tag in GROUP_TAGS


def _group_width_scale(group):
    """Factor converting a child's width into slide units (ext.cx / chExt.cx)."""
    xfrm = group._element.find(f"{_p('grpSpPr')}/{_a('xfrm')}")
    if xfrm is None:
        return 1.0
    ext, ch_ext = xfrm.find(_a("ext")), xfrm.find(_a("chExt"))
    if ext is None or ch_ext is None:
        return 1.0
    try:
        cx, ch_cx = int(ext.get("cx")), int(ch_ext.get("cx"))
    except (TypeError, ValueError):
        return 1.0
    return cx / ch_cx if ch_cx else 1.0


def iter_text_shapes(container, scale=1.0):
    """Yield (shape, effective_width_emu) for every shape with a text frame,
    descending into groups. `effective_width_emu` is None when the shape states
    no width at all — the caller decides what to fall back to."""
    for shape in container.shapes:
        if _is_group(shape):
            yield from iter_text_shapes(shape, scale * _group_width_scale(shape))
        elif shape.has_text_frame:
            width = shape.width
            yield shape, None if width is None else int(width) * scale
