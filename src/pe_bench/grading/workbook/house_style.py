def is_date_format(code: str | None) -> bool:
    return bool(code) and any(token in code.lower() for token in ("yy", "mm", "dd"))
