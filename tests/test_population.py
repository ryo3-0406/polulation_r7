"""
アプリ（model.py / app.py）で集計した人口・世帯数が、e-Statの原数値Excelを単純集計した値と一致するかのテスト。

    python -m unittest discover -s tests -v

検証側はアプリのコード（preprocess.py / config.REGIONS）を使わず、openpyxl で原数値Excelのセルを直接読んで合計する。
ローデータ出力・集計結果のExcelも、出力したファイルを openpyxl で読み戻して原数値Excelと突合する。
"""
import itertools
import random
import sys
import unittest
from io import BytesIO
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parent.parent
SRC = ROOT / "src"
sys.path.insert(0, str(SRC))

from export import build_raw_workbook, build_summary_workbook  # noqa: E402
from model import Population  # noqa: E402

RAW_AGE = ROOT / "data" / "raw" / "b02_01.xlsx"
RAW_HOUSEHOLD = ROOT / "data" / "raw" / "b01_01.xlsx"

# 検証用のエリア定義（アプリ側の config.REGIONS とは独立に記述）
EXPECTED_REGIONS = {
    "関東(1都6県)": ["茨城県", "栃木県", "群馬県", "埼玉県", "千葉県", "東京都", "神奈川県"],
    "東海(3県)": ["岐阜県", "愛知県", "三重県"],
    "関西(2府4県)": ["滋賀県", "京都府", "大阪府", "兵庫県", "奈良県", "和歌山県"],
    "福岡(2県)": ["福岡県", "佐賀県"],
}
PREFECTURES = [
    "北海道", "青森県", "岩手県", "宮城県", "秋田県", "山形県", "福島県", "茨城県", "栃木県", "群馬県",
    "埼玉県", "千葉県", "東京都", "神奈川県", "新潟県", "富山県", "石川県", "福井県", "山梨県", "長野県",
    "岐阜県", "静岡県", "愛知県", "三重県", "滋賀県", "京都府", "大阪府", "兵庫県", "奈良県", "和歌山県",
    "鳥取県", "島根県", "岡山県", "広島県", "山口県", "徳島県", "香川県", "愛媛県", "高知県", "福岡県",
    "佐賀県", "長崎県", "熊本県", "大分県", "宮崎県", "鹿児島県", "沖縄県",
]
SEX_RAW = {"男女": "0_総数", "男": "1_男", "女": "2_女"}
NAT_RAW = {1: "1_うち日本人", 0: "0_国籍総数"}
NAT_SHEETS = {"日本人のみ": 1, "外国人含む": 0}  # 画面の表示名・ローデータのシート名 → flag_japanese
SEX_SHEET = {"男女計": "男女", "男": "男", "女": "女"}  # ローデータの性別 → SEX_RAW のキー


def _num(v):
    return 0 if v == "-" or v is None else int(v)


def load_raw_age():
    """{(国籍, 男女, 地域名): {項目名: 値}}  地域名は '13000_東京都' -> '東京都'"""
    wb = openpyxl.load_workbook(RAW_AGE, read_only=True)
    rows = list(wb.active.iter_rows(values_only=True))
    wb.close()
    items = next(r for r in rows if r[3] == "項目名")
    start = next(i for i, r in enumerate(rows) if r[0] == "国籍総数か日本人") + 1
    data = {}
    for r in rows[start:]:
        if r[0] is None:
            continue
        name = r[3].split("_", 1)[1]
        data[(r[0], r[1], name)] = {item: _num(v) for item, v in zip(items, r) if isinstance(item, str) and "_" in item}
    return data


def load_raw_household():
    """{地域名: 世帯数(0_総数)}  全国・都道府県のみ"""
    wb = openpyxl.load_workbook(RAW_HOUSEHOLD, read_only=True)
    rows = list(wb.active.iter_rows(values_only=True))
    wb.close()
    tables = next(r for r in rows if r[6] == "表章項目")
    items = next(r for r in rows if r[6] == "項目名")
    col = next(i for i, (t, it) in enumerate(zip(tables, items)) if t == "世帯数" and it == "0_総数")
    start = next(i for i, r in enumerate(rows) if r[0] == "地域識別コード") + 1
    return {r[6].split("_", 1)[1]: _num(r[col]) for r in rows[start:] if r[0] == "a"}


def expand(areas):
    """エリア名を都道府県に展開（全国はそのまま）"""
    return [p for a in areas for p in EXPECTED_REGIONS.get(a, [a])]


class ExcelReference:
    """原数値Excelからの単純集計"""

    def __init__(self):
        self.age = load_raw_age()
        self.household = load_raw_household()

    def row(self, nat_flag, sex, name):
        return self.age[(NAT_RAW[nat_flag], SEX_RAW[sex], name)]

    def population(self, areas, sex, nat_flag, age_min, age_max):
        total = 0
        for name in expand(areas):
            row = self.row(nat_flag, sex, name)
            # 年齢 a 歳は項目番号 a+1（'001_0歳' ... '111_110歳以上'）
            total += sum(v for item, v in row.items() if item[:3].isdigit() and age_min + 1 <= int(item[:3]) <= age_max + 1)
        return total

    def item(self, nat_flag, sex, areas, num):
        """項目番号（0=総数, 1..111=0歳..110歳以上, 112=年齢不詳）の値を、地域を展開して合計"""
        prefix = f"{num:03d}_"
        return sum(v for name in expand(areas) for k, v in self.row(nat_flag, sex, name).items() if k.startswith(prefix))

    def unknown(self, areas, sex, nat_flag):
        return self.item(nat_flag, sex, areas, 112)

    def households(self, areas):
        return sum(self.household[name] for name in expand(areas))


def read_workbook(data):
    """{シート名: [行のタプル, ...]}"""
    wb = openpyxl.load_workbook(BytesIO(data), read_only=True)
    sheets = {ws.title: list(ws.iter_rows(values_only=True)) for ws in wb.worksheets}
    wb.close()
    return sheets


class TestPopulationVsExcel(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        cls.model = Population()
        cls.ref = ExcelReference()
        cls.areas = ["全国"] + list(EXPECTED_REGIONS) + PREFECTURES

    def test_area_list(self):
        self.assertEqual(self.model.area_list(flag_household=0), self.areas)
        self.assertEqual(self.model.area_list(flag_household=1), self.areas)

    def test_single_area_all_conditions(self):
        """全地域 × 性別 × 国籍 × 代表的な年齢範囲"""
        age_ranges = [(0, 110), (18, 59), (20, 69), (0, 0), (110, 110), (0, 14), (15, 64), (65, 110), (100, 110), (35, 49)]
        rng = random.Random(0)
        age_ranges += [tuple(sorted(rng.sample(range(111), 2))) for _ in range(10)]
        count = 0
        for area, sex, nat, (lo, hi) in itertools.product(self.areas, SEX_RAW, NAT_RAW, age_ranges):
            with self.subTest(area=area, sex=sex, nat=nat, age=(lo, hi)):
                self.assertEqual(
                    self.model.age_area_population(lo, hi, [area], sex, flag_japanese=nat),
                    self.ref.population([area], sex, nat, lo, hi),
                )
            count += 1
        print(f"\n  単一地域の人口: {count}パターン検証")

    def test_multi_area_random(self):
        """複数地域を同時選択（ランダム500パターン）"""
        rng = random.Random(42)
        for _ in range(500):
            areas = rng.sample(self.areas[1:], rng.randint(2, 8))  # 全国は除く
            sex, nat = rng.choice(list(SEX_RAW)), rng.choice(list(NAT_RAW))
            lo, hi = sorted(rng.sample(range(111), 2))
            with self.subTest(areas=areas, sex=sex, nat=nat, age=(lo, hi)):
                self.assertEqual(
                    self.model.age_area_population(lo, hi, areas, sex, flag_japanese=nat),
                    self.ref.population(areas, sex, nat, lo, hi),
                )

    def test_matches_excel_subtotal_columns(self):
        """Excelに載っている総数・再掲列（15歳未満 等）とアプリの集計が一致するか"""
        subtotals = {
            "R1_（再掲）15歳未満": (0, 14), "R2_（再掲）15～64歳": (15, 64), "R3_（再掲）65歳以上": (65, 110),
            "R4_（再掲）75歳以上": (75, 110), "R5_（再掲）85歳以上": (85, 110), "R6_（再掲）100歳以上": (100, 110),
            "R7_（再掲）20～69歳": (20, 69),
        }
        for area, sex, nat in itertools.product(["全国"] + PREFECTURES, SEX_RAW, NAT_RAW):
            row = self.ref.row(nat, sex, area)
            with self.subTest(area=area, sex=sex, nat=nat):
                for item, (lo, hi) in subtotals.items():
                    self.assertEqual(self.model.age_area_population(lo, hi, [area], sex, nat), row[item], item)
                # 総数 = 0～110歳 + 年齢不詳
                self.assertEqual(self.model.age_area_population(0, 110, [area], sex, nat) + row["112_年齢「不詳」"], row["000_総数"])

    def test_national_equals_sum_of_prefectures(self):
        for sex, nat in itertools.product(SEX_RAW, NAT_RAW):
            with self.subTest(sex=sex, nat=nat):
                self.assertEqual(
                    self.model.age_area_population(0, 110, ["全国"], sex, nat),
                    self.model.age_area_population(0, 110, PREFECTURES, sex, nat),
                )
        self.assertEqual(self.model.household_count(["全国"]), self.model.household_count(PREFECTURES))

    def test_households(self):
        for area in self.areas:
            with self.subTest(area=area):
                self.assertEqual(self.model.household_count([area]), self.ref.households([area]))
        rng = random.Random(7)
        for _ in range(200):
            areas = rng.sample(self.areas[1:], rng.randint(2, 8))
            with self.subTest(areas=areas):
                self.assertEqual(self.model.household_count(areas), self.ref.households(areas))

    def test_unknown_age(self):
        self.assertEqual(self.model.unknown_age_count(0), self.ref.row(0, "男女", "全国")["112_年齢「不詳」"])
        self.assertEqual(self.model.unknown_age_count(1), self.ref.row(1, "男女", "全国")["112_年齢「不詳」"])
        for area, sex, nat in itertools.product(self.areas, SEX_RAW, NAT_RAW):
            with self.subTest(area=area, sex=sex, nat=nat):
                self.assertEqual(self.model.unknown_age_count(nat, [area], sex), self.ref.unknown([area], sex, nat))
                self.assertEqual(self.model.total_count(nat, [area], sex), self.ref.item(nat, sex, [area], 0))
        rng = random.Random(3)
        for _ in range(200):
            areas = rng.sample(self.areas[1:], rng.randint(2, 8))
            sex, nat = rng.choice(list(SEX_RAW)), rng.choice(list(NAT_RAW))
            with self.subTest(areas=areas, sex=sex, nat=nat):
                self.assertEqual(self.model.unknown_age_count(nat, areas, sex), self.ref.unknown(areas, sex, nat))

    def test_invalid_age_range(self):
        with self.assertRaises(ValueError):
            self.model.age_area_population(60, 20, ["東京都"], "男女")
        with self.assertRaises(ValueError):
            self.model.age_area_population(0, 111, ["東京都"], "男女")

    def test_overlapping(self):
        self.assertEqual(Population.overlapping_prefectures(["関東(1都6県)", "東京都"]), ["東京都"])
        self.assertEqual(Population.overlapping_prefectures(["全国", "北海道"]), ["全国"])
        self.assertEqual(Population.overlapping_prefectures(["関東(1都6県)", "関西(2府4県)", "北海道"]), [])


class TestRawExport(unittest.TestCase):
    """ローデータ出力のExcelを読み戻して、原数値Excelと突合する"""

    @classmethod
    def setUpClass(cls):
        cls.ref = ExcelReference()
        cls.sheets = read_workbook(build_raw_workbook(Population(), "2026/01/01"))

    def test_sheets(self):
        self.assertEqual(list(self.sheets), ["日本人のみ", "外国人含む", "注記"])

    def test_header(self):
        expected = ("区分", "地域コード", "地域名", "性別", "総数", "年齢不詳", "年齢不詳を除く計")
        expected += tuple(f"{a}歳" for a in range(110)) + ("110歳以上",)
        for sheet in NAT_SHEETS:
            self.assertEqual(self.sheets[sheet][0], expected)

    def test_row_order(self):
        """性別ごとに 全国 → 47都道府県（コード順） → 主要エリア"""
        expected = [
            (kubun, code, name, sex)
            for sex in SEX_SHEET
            for kubun, code, name in (
                [("全国", "00000", "全国")]
                + [("都道府県", f"{i:02d}000", p) for i, p in enumerate(PREFECTURES, start=1)]
                + [("主要エリア", None, r) for r in EXPECTED_REGIONS]
            )
        ]
        for sheet in NAT_SHEETS:
            with self.subTest(sheet=sheet):
                self.assertEqual([(r[0], r[1] or None, r[2], r[3]) for r in self.sheets[sheet][1:]], expected)

    def test_values_match_excel(self):
        count = 0
        for sheet, nat in NAT_SHEETS.items():
            for row in self.sheets[sheet][1:]:
                name, sex = row[2], SEX_SHEET[row[3]]
                with self.subTest(sheet=sheet, area=name, sex=row[3]):
                    total, unknown, known, *ages = row[4:]
                    self.assertEqual(total, self.ref.item(nat, sex, [name], 0), "総数")
                    self.assertEqual(unknown, self.ref.item(nat, sex, [name], 112), "年齢不詳")
                    self.assertEqual(known, total - unknown, "年齢不詳を除く計")
                    self.assertEqual(ages, [self.ref.item(nat, sex, [name], a + 1) for a in range(111)], "各歳")
                count += 1
        print(f"\n  ローデータ: {count}行 × 114列 を検証")

    def test_notes(self):
        notes = dict(self.sheets["注記"][1:])
        for key in ["シート「日本人のみ」", "シート「外国人含む」", "年齢不詳", "出典", "出力日"]:
            self.assertIn(key, notes)
        self.assertEqual(notes["出力日"], "2026/01/01")


class TestSummaryExport(unittest.TestCase):

    def test_summary_workbook(self):
        from app import evaluate, summary_frame
        model, ref = Population(), ExcelReference()
        conds = [
            {"area": ["関東(1都6県)"], "nat": 1, "sex": "男女", "age_min": 18, "age_max": 59},
            {"area": ["東京都", "大阪府"], "nat": 0, "sex": "女", "age_min": 65, "age_max": 110},
        ]
        items = [(n, c, evaluate(model, "人口", c)) for n, c in enumerate(conds, start=1)]
        df, number_cols = summary_frame("人口", items)
        header, *rows = read_workbook(build_summary_workbook(df, number_cols, "2026/01/01"))["集計結果"]
        self.assertEqual(header, ("集計番号", "地域", "国籍", "性別", "年齢", "人口（人）", "年齢不詳（集計外・参考）"))
        self.assertEqual(rows, [
            (1, "関東(1都6県)", "日本人のみ", "男女計", "18～59歳", ref.population(["関東(1都6県)"], "男女", 1, 18, 59),
             ref.unknown(["関東(1都6県)"], "男女", 1)),
            (2, "東京都、大阪府", "外国人含む", "女", "65歳以上", ref.population(["東京都", "大阪府"], "女", 0, 65, 110),
             ref.unknown(["東京都", "大阪府"], "女", 0)),
        ])

    def test_clean_filename(self):
        from app import clean_filename
        self.assertEqual(clean_filename("集計", "d.xlsx"), "集計.xlsx")
        self.assertEqual(clean_filename("集計.XLSX", "d.xlsx"), "集計.XLSX")
        self.assertEqual(clean_filename("  ", "d.xlsx"), "d.xlsx")
        self.assertEqual(clean_filename("a/b:c", "d.xlsx"), "a_b_c.xlsx")


class TestStreamlitApp(unittest.TestCase):
    """実際にStreamlitアプリを動かし、画面に表示された集計結果をExcelと比較する"""

    @classmethod
    def setUpClass(cls):
        from streamlit.testing.v1 import AppTest
        cls.AppTest = AppTest
        cls.ref = ExcelReference()

    def new_app(self):
        at = self.AppTest.from_file(str(SRC / "app.py"), default_timeout=60).run()
        self.assertFalse(at.exception, at.exception)
        return at

    def run_app(self, target, conditions):
        at = self.new_app()
        at.radio(key="target").set_value(target).run()
        for _ in range(len(conditions) - 1):
            at.button(key="add_condition").click().run()
        for i, cond in enumerate(conditions):
            at.multiselect(key=f"area_{i}").set_value(cond["area"])
            if target == "人口":
                at.radio(key=f"nat_{i}").set_value(NAT_SHEETS[cond["nat"]])
                at.radio(key=f"sex_{i}").set_value(cond["sex"])
                at.number_input(key=f"age_min_{i}").set_value(cond["age"][0])
                at.number_input(key=f"age_max_{i}").set_value(cond["age"][1])
        at.run()
        self.assertFalse(at.exception, at.exception)
        return at

    def test_population_screen(self):
        conditions = [
            {"area": ["関東(1都6県)"], "nat": "日本人のみ", "sex": "男女", "age": (18, 59)},
            {"area": ["関西(2府4県)", "東海(3県)"], "nat": "外国人含む", "sex": "女", "age": (20, 34)},
            {"area": ["全国"], "nat": "外国人含む", "sex": "男女", "age": (0, 110)},
            {"area": ["福岡(2県)", "北海道", "沖縄県"], "nat": "日本人のみ", "sex": "男", "age": (65, 110)},
        ]
        at = self.run_app("人口", conditions)
        shown = [int(c.value) for c in at.code]
        expected = [self.ref.population(c["area"], c["sex"], NAT_SHEETS[c["nat"]], *c["age"]) for c in conditions]
        print(f"\n  画面表示: {shown}\n  Excel集計: {expected}")
        self.assertEqual(shown, expected)

        # 集計結果一覧（年齢不詳の参考値つき）
        summary = at.dataframe[0].value
        self.assertEqual(summary["人口（人）"].tolist(), expected)
        self.assertEqual(
            summary["年齢不詳（集計外・参考）"].tolist(),
            [self.ref.unknown(c["area"], c["sex"], NAT_SHEETS[c["nat"]]) for c in conditions],
        )

    def test_household_screen(self):
        conditions = [{"area": ["関東(1都6県)"]}, {"area": ["東京都", "大阪府"]}, {"area": ["全国"]}]
        at = self.run_app("世帯数", conditions)
        shown = [int(c.value) for c in at.code]
        expected = [self.ref.households(c["area"]) for c in conditions]
        print(f"\n  画面表示: {shown}\n  Excel集計: {expected}")
        self.assertEqual(shown, expected)
        self.assertEqual(at.dataframe[0].value["世帯数（世帯）"].tolist(), expected)

    def test_invalid_age_shows_error(self):
        at = self.run_app("人口", [{"area": ["東京都"], "nat": "日本人のみ", "sex": "男女", "age": (60, 20)}])
        self.assertTrue(at.error)
        self.assertFalse(at.code)

    def test_incomplete_condition_is_excluded_from_summary(self):
        at = self.run_app("人口", [
            {"area": ["東京都"], "nat": "日本人のみ", "sex": "男女", "age": (20, 29)},
            {"area": [], "nat": "日本人のみ", "sex": "男女", "age": (20, 29)},
        ])
        self.assertEqual(at.dataframe[0].value["集計番号"].tolist(), [1])
        self.assertTrue(any("集計 2" in i.value for i in at.info))

    def test_duplicate_remove_reset(self):
        at = self.new_app()
        at.multiselect(key="area_0").set_value(["東京都"])
        at.number_input(key="age_min_0").set_value(30).run()
        at.button(key="dup_0").click().run()
        self.assertEqual(at.session_state.cond_ids, [0, 1])
        self.assertEqual(at.multiselect(key="area_1").value, ["東京都"])
        self.assertEqual(at.number_input(key="age_min_1").value, 30)

        at.number_input(key="age_min_1").set_value(40).run()
        at.button(key="del_0").click().run()
        self.assertEqual(at.session_state.cond_ids, [1])
        self.assertEqual([int(c.value) for c in at.code], [self.ref.population(["東京都"], "男女", 1, 40, 59)])

        at.button(key="reset_conditions").click().run()
        self.assertEqual(len(at.session_state.cond_ids), 1)
        self.assertFalse(at.code)
        self.assertFalse(at.exception, at.exception)

    def test_inputs_survive_target_switch(self):
        """世帯数に切り替えて戻っても、人口の条件（年齢など）が保持される"""
        at = self.run_app("人口", [{"area": ["愛知県"], "nat": "外国人含む", "sex": "女", "age": (25, 44)}])
        at.radio(key="target").set_value("世帯数").run()
        self.assertEqual([int(c.value) for c in at.code], [self.ref.households(["愛知県"])])
        at.radio(key="target").set_value("人口").run()
        self.assertEqual(at.number_input(key="age_min_0").value, 25)
        self.assertEqual(at.radio(key="sex_0").value, "女")
        self.assertEqual([int(c.value) for c in at.code], [self.ref.population(["愛知県"], "女", 0, 25, 44)])

    def test_max_conditions(self):
        at = self.new_app()
        for _ in range(25):
            at.button(key="add_condition").click().run()
        self.assertEqual(len(at.session_state.cond_ids), 20)
        self.assertTrue(at.button(key="add_condition").disabled)


if __name__ == "__main__":
    unittest.main()
