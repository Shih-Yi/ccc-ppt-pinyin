"""Resolve the effective font size (in points) of a PPTX paragraph.

A run only carries an explicit `sz` attribute when the size was set on that
run. Decks exported from other tools — Google Slides in particular — leave
every run bare and keep the size in the shape's list style instead:

    <a:lstStyle><a:lvl1pPr><a:defRPr sz="4800"/>   <- 48pt lives here
    ...
    <a:r><a:rPr lang="en-US"/><a:t>祢是我喜樂泉源</a:t></a:r>   <- no sz

Reading the run alone reports "size unknown" for perfectly ordinary lyrics,
so this module walks the whole inheritance chain PowerPoint itself uses:

    run rPr -> paragraph defRPr -> shape lstStyle -> layout placeholder
    -> master placeholder -> master txStyles -> presentation defaultTextStyle

Every lookup is read-only; nothing here mutates the presentation.
"""
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"

MAX_LEVEL = 9  # OOXML defines lvl1pPr .. lvl9pPr

# placeholder type (the `type` attribute of <p:ph>) -> master text style.
# "obj" is the schema default, hence None maps to the body style too.
TITLE_PH_TYPES = frozenset({"title", "ctrTitle"})
BODY_PH_TYPES = frozenset({"body", "subTitle", "obj", None})
# layout placeholder type -> the master placeholder it inherits from
MASTER_PH_TYPE = {"ctrTitle": "title", "subTitle": "body", "obj": "body"}


def _a(tag):
    return f"{{{A_NS}}}{tag}"


def _p(tag):
    return f"{{{P_NS}}}{tag}"


def _sz_pt(rpr):
    """Font size in pt from an rPr/defRPr element, or None if unset."""
    if rpr is None:
        return None
    sz = rpr.get("sz")
    return int(sz) / 100 if sz else None


def _para_level(p_el):
    """Outline level of a paragraph (0-based), as stored in pPr@lvl."""
    p_pr = p_el.find(_a("pPr"))
    if p_pr is None:
        return 0
    try:
        lvl = int(p_pr.get("lvl") or 0)
    except ValueError:
        return 0
    return lvl if 0 <= lvl < MAX_LEVEL else 0


def _level_def_rpr(list_style, lvl):
    """defRPr for outline level `lvl` inside a lstStyle / txStyle element."""
    if list_style is None:
        return None
    lvl_pr = list_style.find(_a(f"lvl{lvl + 1}pPr"))
    if lvl_pr is None:
        return None
    return lvl_pr.find(_a("defRPr"))


def _run_pt(p_el):
    """Size from the first run that states one explicitly."""
    for r in p_el.findall(_a("r")):
        pt = _sz_pt(r.find(_a("rPr")))
        if pt is not None:
            return pt
    return None


def _ph(shape):
    """The <p:ph> element of a shape (it lives under nvPr), or None."""
    if shape is None:
        return None
    return shape._element.find(f".//{_p('ph')}")


def _ph_type(shape):
    """The <p:ph> type of a shape.

    Returns the sentinel "" for a shape that has no <p:ph> at all, so callers
    can tell "not a placeholder" from "placeholder with the default type"
    (which is None, meaning "obj").
    """
    ph = _ph(shape)
    return "" if ph is None else ph.get("type")


def _ph_idx(shape):
    ph = _ph(shape)
    if ph is None:
        return None
    try:
        return int(ph.get("idx") or 0)
    except ValueError:
        return None


def _shape_list_style(shape):
    """The lstStyle of a shape's txBody, if any."""
    if shape is None or not getattr(shape, "has_text_frame", False):
        return None
    return shape.text_frame._txBody.find(_a("lstStyle"))


def _matching_placeholder(container, idx, ph_type):
    """Placeholder in a layout/master matching by idx first, then by type."""
    if container is None:
        return None
    shapes = list(container.placeholders)
    if idx is not None:
        for ph in shapes:
            if _ph_idx(ph) == idx:
                return ph
    for ph in shapes:
        if _ph_type(ph) == ph_type:
            return ph
    return None


def _master_text_style(master, ph_type):
    """titleStyle / bodyStyle / otherStyle element of a slide master."""
    if master is None:
        return None
    tx_styles = master._element.find(_p("txStyles"))
    if tx_styles is None:
        return None
    if ph_type in TITLE_PH_TYPES:
        name = "titleStyle"
    elif ph_type in BODY_PH_TYPES:
        name = "bodyStyle"
    else:
        name = "otherStyle"
    return tx_styles.find(_p(name))


def _default_text_style(prs):
    """presentation.xml <p:defaultTextStyle>, the last fallback."""
    if prs is None:
        return None
    return prs._element.find(_p("defaultTextStyle"))


def _list_style_chain(shape, slide):
    """lstStyle / txStyle elements a shape inherits from, nearest first."""
    yield _shape_list_style(shape)

    ph_type = _ph_type(shape) if shape is not None else ""
    if ph_type == "" or slide is None:
        return  # a plain text box inherits nothing from layout or master

    idx = _ph_idx(shape)
    layout = slide.slide_layout
    layout_ph = _matching_placeholder(layout, idx, ph_type)
    yield _shape_list_style(layout_ph)

    master = layout.slide_master
    master_type = MASTER_PH_TYPE.get(ph_type, ph_type)
    master_ph = _matching_placeholder(master, None, master_type)
    yield _shape_list_style(master_ph)

    yield _master_text_style(master, ph_type)


def resolve_font_pt(paragraph, shape=None, slide=None, prs=None):
    """Effective font size of `paragraph` in points, or None if nothing in
    the inheritance chain states one.

    `shape`, `slide` and `prs` widen the search; without them only the
    paragraph's own runs and pPr are consulted.
    """
    p_el = paragraph._p
    lvl = _para_level(p_el)

    pt = _run_pt(p_el)
    if pt is not None:
        return pt

    p_pr = p_el.find(_a("pPr"))
    if p_pr is not None:
        pt = _sz_pt(p_pr.find(_a("defRPr")))
        if pt is not None:
            return pt

    for list_style in _list_style_chain(shape, slide):
        pt = _sz_pt(_level_def_rpr(list_style, lvl))
        if pt is not None:
            return pt

    return _sz_pt(_level_def_rpr(_default_text_style(prs), lvl))
