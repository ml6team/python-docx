"""Test numbering definition identifiers."""

from typing import Sequence

import pytest

import docx.oxml.numbering as numbering_module
from docx.oxml.numbering import CT_Numbering

from ..unitutil.cxml import element


class DescribeCT_Numbering:
    @pytest.mark.parametrize(
        ("num_ids", "expected"),
        [
            ([], 1),
            ([1], 2),
            ([1, 2, 3], 4),
            ([2, 3], 1),
            ([1, 3, 4], 2),
            ([5, 2, 1, 4], 3),
            ([1, 1, 2, 4], 3),
            ([0, -1, 2, 1], 3),
            ([1000000], 1),
        ],
    )
    def it_finds_the_first_unused_positive_numId(
        self, num_ids: Sequence[int], expected: int
    ) -> None:
        numbering = element("w:numbering")
        assert isinstance(numbering, CT_Numbering)
        for num_id in num_ids:
            numbering.append(element("w:num{w:numId=%d}" % num_id))

        assert numbering._next_numId == expected

    def it_uses_the_first_gap_when_adding_numbering_definitions(self) -> None:
        numbering = element("w:numbering/(w:num{w:numId=1},w:num{w:numId=3})")
        assert isinstance(numbering, CT_Numbering)

        num = numbering.add_num(abstractNum_id=7)

        assert num.numId == 2
        assert num.abstractNumId.val == 7
        assert numbering._next_numId == 4

    def it_rejects_non_integer_numIds(self) -> None:
        numbering = element("w:numbering/w:num{w:numId=invalid}")
        assert isinstance(numbering, CT_Numbering)

        with pytest.raises(ValueError, match="invalid literal"):
            numbering._next_numId

    def it_does_not_repeat_linear_membership_scans(
        self, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        comparison_count = 0

        class CountingInt(int):
            __hash__ = int.__hash__

            def __eq__(self, other: object) -> bool:
                nonlocal comparison_count
                comparison_count += 1
                return int.__eq__(self, other)

        numbering = element("w:numbering")
        assert isinstance(numbering, CT_Numbering)
        count = 512
        for num_id in range(1, count + 1):
            numbering.append(element("w:num{w:numId=%d}" % num_id))
        monkeypatch.setattr(numbering_module, "int", CountingInt, raising=False)

        assert numbering._next_numId == count + 1
        assert comparison_count <= 2 * count
