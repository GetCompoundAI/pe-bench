import functools
import logging
import os
import shutil
import subprocess
import tempfile
from collections.abc import Iterator
from contextlib import contextmanager
from pathlib import Path

from pe_bench.core import Unreachable

logger = logging.getLogger(__name__)

PDFTOPPM = "pdftoppm"
PDFINFO = "pdfinfo"
PDFTOTEXT = "pdftotext"
PDFIMAGES = "pdfimages"
POPPLER_INSTALL = "apt install poppler-utils"
TIMEOUT_SECONDS = 300


@functools.cache
def _missing_on(path: str | None) -> tuple[str, ...]:
    return tuple(
        tool
        for tool in (PDFTOPPM, PDFINFO, PDFTOTEXT, PDFIMAGES)
        if shutil.which(tool, path=path) is None
    )


def require_poppler() -> None:
    missing = _missing_on(os.environ.get("PATH"))
    if missing:
        raise Unreachable(
            f"{', '.join(f'`{tool}`' for tool in missing)} not on PATH; install "
            f"Poppler with `{POPPLER_INSTALL}`"
        )


@contextmanager
def _scratch(data: bytes) -> Iterator[tuple[Path, Path]]:
    require_poppler()
    with tempfile.TemporaryDirectory(prefix="pdf-tools-") as raw:
        scratch = Path(raw)
        pdf = scratch / "in.pdf"
        pdf.write_bytes(data)
        yield pdf, scratch


def _texts_in(stdout: bytes) -> list[str]:
    pages = stdout.decode("utf-8", "replace").split("\f")
    return pages[:-1] if pages and pages[-1] == "" else pages


def _hung(argv: list[str]) -> None:
    logger.warning(
        {"type": "poppler_hung", "tool": argv[0], "seconds": TIMEOUT_SECONDS}
    )


def _ran(argv: list[str]) -> bytes | None:
    try:
        done = subprocess.run(
            argv, capture_output=True, timeout=TIMEOUT_SECONDS, check=False
        )
    except subprocess.TimeoutExpired:
        _hung(argv)
        return None
    return done.stdout if done.returncode == 0 else None


def page_texts(data: bytes, *, first_pages: int | None = None) -> list[str] | None:
    pages = [] if first_pages is None else ["-f", "1", "-l", str(first_pages)]
    with _scratch(data) as (pdf, _):
        stdout = _ran([PDFTOTEXT, "-raw", *pages, "-enc", "UTF-8", str(pdf), "-"])
    return None if stdout is None else _texts_in(stdout)
