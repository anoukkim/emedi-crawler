import tempfile
import unittest
from pathlib import Path

from models import MODEL_COL, expand_by_model, expand_records


BASE = {"품목허가번호": "수허24-229호", "제품정보_품목명": "초음파 영상 진단장치", "제품정보_등급": "3"}


class ExpandTests(unittest.TestCase):
    def test_udi_rows_one_row_per_model(self):
        record = {**BASE, "UDI_모델명": "A; B", "UDI_UDI": "001; 002",
                  "_UDI_rows": [{"연번": "1", "모델명": "A", "UDI": "001", "포장 내 수량": "1"},
                                {"연번": "2", "모델명": "B", "UDI": "002", "포장 내 수량": "10"}]}
        rows = expand_by_model(record)
        self.assertEqual([r[MODEL_COL] for r in rows], ["A", "B"])
        self.assertEqual([r["UDI_UDI"] for r in rows], ["001", "002"])
        self.assertEqual(rows[1]["UDI_포장 내 수량"], "10")
        self.assertEqual(rows[1]["모델순번"], "2/2")
        self.assertNotIn("UDI_연번", rows[0])
        self.assertNotIn("_UDI_rows", rows[0])
        for r in rows:   # 허가증 공통 항목은 모든 행에 복사
            self.assertEqual(r["제품정보_품목명"], BASE["제품정보_품목명"])

    def test_old_files_split_joined_udi_columns(self):
        rows = expand_by_model({**BASE, "UDI_모델명": "A; B", "UDI_UDI": "001; 002", "UDI_UDI 사용중단": "; "})
        self.assertEqual([(r[MODEL_COL], r["UDI_UDI"]) for r in rows], [("A", "001"), ("B", "002")])

    def test_mismatched_udi_columns_keep_only_names(self):
        rows = expand_by_model({**BASE, "UDI_모델명": "A; B", "UDI_UDI": "001"})
        self.assertEqual([r[MODEL_COL] for r in rows], ["A", "B"])

    def test_semicolon_model_names_without_udi(self):
        rows = expand_by_model({**BASE, "제품정보_모델명": "M-1; M-2;M-3"})
        self.assertEqual([r[MODEL_COL] for r in rows], ["M-1", "M-2", "M-3"])

    def test_record_without_models_is_kept(self):
        rows = expand_records([{**BASE, "오류": "상세 정보 미수집"}])
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0][MODEL_COL], "")


class ExcelSheetTests(unittest.TestCase):
    def test_recovery_reads_permit_sheet(self):
        import pandas as pd
        from recovery import PERMIT_SHEET, load_existing_records
        record = {"품목일련번호": "1", **BASE, "UDI_모델명": "A; B", "UDI_UDI": "001; 002"}
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "crawl.xlsx"
            with pd.ExcelWriter(path, engine="openpyxl") as writer:
                pd.DataFrame(expand_records([record])).to_excel(writer, sheet_name="모델명별", index=False)
                pd.DataFrame([record]).to_excel(writer, sheet_name=PERMIT_SHEET, index=False)
            self.assertEqual(len(load_existing_records(path)), 1)


if __name__ == "__main__":
    unittest.main()
