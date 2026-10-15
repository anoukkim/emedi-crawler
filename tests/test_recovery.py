import json
import tempfile
import unittest
from pathlib import Path

from recovery import build_recovery_plan, detail_complete, load_existing_records


COLS = ["순번", "품목허가번호", "품목일련번호"]


def item(n):
    return {"순번": str(n), "품목허가번호": f"수인26-{n}호", "품목일련번호": str(1000 + n)}


def completed(n):
    return {**item(n), **{f"제품정보_항목{i}": str(i) for i in range(8)}}


class RecoveryTests(unittest.TestCase):
    def test_failed_and_new_only(self):
        old = [completed(1), {**item(2), "오류": "모달이 비어 있음"}]
        records, queue, overlap, missing = build_recovery_plan([item(1), item(2), item(3)], old, COLS)
        self.assertEqual((queue, overlap, missing), ([1, 2], 2, 0))
        self.assertEqual(records[0]["제품정보_항목0"], "0")
        self.assertEqual(records[1]["오류"], "모달이 비어 있음")
        self.assertEqual(records[2]["오류"], "상세 정보 미수집")

    def test_incomplete_modal_is_not_success(self):
        self.assertFalse(detail_complete({**item(1), "제품정보_품목명": "품목"}))
        self.assertTrue(detail_complete(completed(1)))

    def test_mismatched_search_does_not_drop_old_rows(self):
        with self.assertRaisesRegex(ValueError, "현재 검색 목록"):
            build_recovery_plan([item(1)], [completed(1), completed(2)], COLS)

    def test_load_json(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "partial.json"
            path.write_text(json.dumps([completed(1)], ensure_ascii=False), encoding="utf-8")
            self.assertEqual(load_existing_records(path)[0]["품목일련번호"], "1001")


if __name__ == "__main__":
    unittest.main()
