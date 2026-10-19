import io
import posixpath
import xml.etree.ElementTree as ET
import zipfile

NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_C = "http://schemas.openxmlformats.org/drawingml/2006/chart"
NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"
NS_R = "http://schemas.openxmlformats.org/package/2006/relationships"
NS_R_OFFICE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_CX = "http://schemas.microsoft.com/office/drawing/2014/chartex"


def open_deck(raw: bytes) -> zipfile.ZipFile | None:
    try:
        return zipfile.ZipFile(io.BytesIO(raw))
    except zipfile.BadZipFile:
        return None


def part_rels(zf: zipfile.ZipFile, part_name: str) -> dict[str, tuple[str, str | None]]:
    d, base = posixpath.split(part_name)
    rels_name = posixpath.join(d, "_rels", base + ".rels")
    if rels_name not in zf.namelist():
        return {}
    out: dict[str, tuple[str, str | None]] = {}
    for rel in ET.fromstring(zf.read(rels_name)).findall(f"{{{NS_R}}}Relationship"):
        rid = rel.get("Id")
        target = rel.get("Target", "")
        if not rid or not target:
            continue
        if rel.get("TargetMode") == "External":
            out[rid] = (target, "External")
        else:
            out[rid] = (posixpath.normpath(posixpath.join(d, target)).lstrip("/"), None)
    return out


def chart_graphic_frames(slide_root: ET.Element) -> list[ET.Element]:
    return [
        gf
        for gf in slide_root.iter(f"{{{NS_P}}}graphicFrame")
        if any(
            "chart" in (gd.get("uri") or "").lower()
            for gd in gf.iter(f"{{{NS_A}}}graphicData")
        )
    ]


def frame_chart_rid(gf: ET.Element) -> str | None:
    return next(
        (
            ch.get(f"{{{NS_R_OFFICE}}}id")
            for tag in (f"{{{NS_C}}}chart", f"{{{NS_CX}}}chart")
            for ch in gf.iter(tag)
            if ch.get(f"{{{NS_R_OFFICE}}}id")
        ),
        None,
    )


def ordered_slide_names(zf: zipfile.ZipFile) -> list[str]:
    by_file = sorted(
        (
            n
            for n in zf.namelist()
            if n.startswith("ppt/slides/slide") and n.endswith(".xml")
        ),
        key=lambda n: int(n[len("ppt/slides/slide") : -len(".xml")]),
    )
    if "ppt/presentation.xml" not in zf.namelist():
        return by_file
    lst = ET.fromstring(zf.read("ppt/presentation.xml")).find(f"{{{NS_P}}}sldIdLst")
    if lst is None:
        return by_file
    rels = part_rels(zf, "ppt/presentation.xml")
    names = [
        target[0]
        for s in lst.findall(f"{{{NS_P}}}sldId")
        if (target := rels.get(s.get(f"{{{NS_R_OFFICE}}}id", ""))) is not None
        and target[0] in zf.namelist()
    ]
    return names or by_file


def slide_chart_part_refs(zf: zipfile.ZipFile, slide_name: str) -> list[str]:
    if slide_name not in zf.namelist():
        return []
    rels = part_rels(zf, slide_name)
    refs: list[str] = []
    for gf in chart_graphic_frames(ET.fromstring(zf.read(slide_name))):
        rid = frame_chart_rid(gf)
        target = rels.get(rid) if rid else None
        if target is not None and target[1] is None and target[0] in zf.namelist():
            refs.append(target[0])
    return refs
