CITATIONS_PART_PATH = "customXml/citations.json"


def _expanded_cell(
    cell: object, sources: dict[str, object]
) -> dict[str, object] | None:
    if not isinstance(cell, dict):
        return None
    sheet, row, col = cell.get("sheet"), cell.get("row"), cell.get("col")
    keys = cell.get("source_keys")
    drive_id = cell.get("drive_id")
    if (
        not isinstance(sheet, str)
        or not isinstance(row, int)
        or not isinstance(col, int)
        or not isinstance(keys, list)
        or not all(isinstance(key, str) and key in sources for key in keys)
        or not (drive_id is None or isinstance(drive_id, str))
    ):
        return None
    body: dict[str, object] = {} if drive_id is None else {"drive_id": drive_id}
    body["sources"] = [sources[key] for key in keys]
    return {"sheet": sheet, "row": row, "col": col, "citation": body}


def cell_citation_entries(envelope: object) -> list[object]:
    if not isinstance(envelope, dict):
        return []
    inline = envelope.get("citations")
    entries: list[object] = list(inline) if isinstance(inline, list) else []
    sources = envelope.get("sources")
    cited_cells = envelope.get("cited_cells")
    if not isinstance(sources, dict) or not isinstance(cited_cells, list):
        return entries
    for cell in cited_cells:
        expanded = _expanded_cell(cell, sources)
        if expanded is not None:
            entries.append(expanded)
    return entries
