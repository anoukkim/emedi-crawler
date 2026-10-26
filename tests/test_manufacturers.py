import unittest

from manufacturers import add_manufacturer_columns, split_manufacturers


DUAL = {
    "제조자": "Boston Scientific Corporation^Boston Scientific Medical Device (Malaysia) SDN. Bhd.",
    "수입업자_제조자 상호": "Boston Scientific Corporation, Boston Scientific Medical Device (Malaysia) SDN. Bhd.",
    "수입업자_제조자 주소": "Two Scimed Place, Maple Grove, MN 55311, USA , PMT 741, Persiaran Cassia Selatan 1, "
                         "Taman Perindustrian Batu Kawan, 14110 Bandar Cassia, Pulau Pinang, Malaysia",
}


class ManufacturerTests(unittest.TestCase):
    def test_dual_split(self):
        pairs, problem = split_manufacturers(DUAL)
        self.assertEqual(problem, "")
        self.assertEqual(pairs, [
            ("Boston Scientific Corporation", "Two Scimed Place, Maple Grove, MN 55311, USA"),
            ("Boston Scientific Medical Device (Malaysia) SDN. Bhd.",
             "PMT 741, Persiaran Cassia Selatan 1, Taman Perindustrian Batu Kawan, 14110 Bandar Cassia, "
             "Pulau Pinang, Malaysia"),
        ])

    def test_comma_in_name_and_second_address(self):
        pairs, problem = split_manufacturers({
            "제조자": "FMD Co.,Ltd.^FMD Co., Ltd. Okinawa Factory",
            "수입업자_제조자 주소": "1-166 Shimoobari Nakashima Komaki, Aichi 485-0051 Japan , "
                                 "5194-15 Katsuren Haebaru, Uruma, Okinawa, 904-2311 Japan, 5194-64 Kasturen",
        })
        self.assertEqual(problem, "")
        self.assertEqual([n for n, _ in pairs], ["FMD Co.,Ltd.", "FMD Co., Ltd. Okinawa Factory"])
        self.assertTrue(pairs[1][1].startswith("5194-15") and pairs[1][1].endswith("Kasturen"))

    def test_single_keeps_whole_address(self):
        pairs, problem = split_manufacturers({"제조자": "LEAD-LOK, Inc.", "수입업자_제조자 주소": "500 W, Sandpoint, ID"})
        self.assertEqual((pairs, problem), ([("LEAD-LOK, Inc.", "500 W, Sandpoint, ID")], ""))

    def test_count_mismatch_is_flagged(self):
        pairs, problem = split_manufacturers({"제조자": "A^B^C", "수입업자_제조자 주소": "x, y , z"})
        self.assertEqual(pairs, [("A", ""), ("B", ""), ("C", "")])
        self.assertEqual(problem, "제조자 3곳 / 주소 2개")

    def test_columns_follow_manufacturer_and_pad(self):
        rows = add_manufacturer_columns([
            {"순번": "1", "제조자": "Solo  Corp", "주소": "서울", "수입업자_제조자 주소": "1 Main St, USA"},
            {"순번": "2", "제조자": DUAL["제조자"], "주소": "서울", "수입업자_제조자 주소": ""},
        ])
        self.assertEqual(list(rows[0])[:9], ["순번", "제조자", "제조자수", "제조자1_상호", "제조자1_주소",
                                            "제조자2_상호", "제조자2_주소", "제조자_확인필요", "주소"])
        self.assertEqual(rows[0]["제조자1_상호"], "Solo Corp")
        self.assertEqual(rows[0]["제조자2_상호"], "")
        self.assertEqual((rows[1]["제조자수"], rows[1]["제조자2_주소"]), (2, ""))

    def test_rerun_replaces_old_columns(self):
        once = add_manufacturer_columns([{"제조자": "A^B^C"}])
        twice = add_manufacturer_columns([{**once[0], "제조자": "A"}])
        self.assertNotIn("제조자3_상호", twice[0])


if __name__ == "__main__":
    unittest.main()
