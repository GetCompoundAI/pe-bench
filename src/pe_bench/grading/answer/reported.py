from pe_bench.core import Result, answer_text


def nothing_delivered(result: Result) -> bool:
    return not answer_text(result).strip() and not result.files
