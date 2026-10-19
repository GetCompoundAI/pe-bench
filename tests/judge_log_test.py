import hashlib
from collections.abc import Iterator
from dataclasses import dataclass, field

import pytest
from pydantic import BaseModel

from pe_bench import judge
from pe_bench.core import JudgeLog, Result, grade, judge_log_kept, rubric
from pe_bench.grading.workbook.roles import logged_answer


class _Located(BaseModel):
    cells: str


@dataclass
class _Scripted:
    replies: list[str]
    asked: list[str] = field(default_factory=list)

    async def text(self, system: str, question: str) -> str | None:
        self.asked.append(question)
        return self.replies.pop(0)


@pytest.fixture(autouse=True)
def _no_live_judge() -> Iterator[None]:
    yield
    judge._configured.clear()


def _sha(question: str) -> str:
    return hashlib.sha256(question.encode()).hexdigest()


async def test_one_grading_asks_a_repeated_locator_question_once() -> None:
    scripted = _Scripted(['{"cells": "C4:F4"}'])
    judge.configure(scripted)
    located: list[str] = []

    async def locates(result: Result) -> bool:
        answer = await logged_answer("where is revenue?", _Located, instruction="")
        located.append(answer.cells)
        return True

    first, second = rubric(weight=1)(locates), rubric(weight=1)(locates)
    verdicts = await grade([first, second], Result(answers=[], files={}))
    assert ([True, True], ["C4:F4", "C4:F4"], 1) == (
        [one.passed for one in verdicts],
        located,
        len(scripted.asked),
    )


async def test_a_stored_answer_replays_without_asking_the_judge() -> None:
    scripted = _Scripted([])
    judge.configure(scripted)
    log = JudgeLog(stored={_sha("where is revenue?"): {"cells": "C4:F4"}})
    with judge_log_kept(log):
        answer = await logged_answer("where is revenue?", _Located, instruction="")
    assert ("C4:F4", [], {_sha("where is revenue?"): {"cells": "C4:F4"}}) == (
        answer.cells,
        scripted.asked,
        log.recorded,
    )


async def test_an_unparseable_locator_reply_replays_as_no_verdict() -> None:
    scripted = _Scripted(["not json"] * judge.ATTEMPTS)
    judge.configure(scripted)
    log = JudgeLog()
    with judge_log_kept(log), pytest.raises(judge.NoVerdict):
        await logged_answer("where is revenue?", _Located, instruction="")
    judge.configure(_Scripted([]))
    with judge_log_kept(JudgeLog(stored=log.recorded)), pytest.raises(judge.NoVerdict):
        await logged_answer("where is revenue?", _Located, instruction="")
