"""
Excel出力

- 集計結果：条件別集計の一覧（集計結果シート＋注記シート）
- ローデータ：全国・47都道府県・主要エリア × 男女 × 年齢各歳の人口
  （「日本人のみ」「外国人含む」でシートを分け、年齢不詳の人数も載せる）
"""
from io import BytesIO

import pandas as pd

from config import AGE_MAX, CENSUS_URL, NAT_LABELS, REGIONS, SEX_LABELS
from model import NATIONALITY

AGES = list(range(AGE_MAX + 1))
AREA_KUBUN = {"全国": "全国", "都道府県": "都道府県", "エリア": "主要エリア"}  # 前処理CSVの区分 → 表示名（並び順も兼ねる）
RAW_KEY_COLS = ["区分", "地域コード", "地域名", "性別"]
RAW_TOTAL_COLS = ["総数", "年齢不詳", "年齢不詳を除く計"]
NOTES_SHEET = "注記"
HEADER_COLOR = "#DDEBF7"


def age_label(age):
    """0 -> '0歳'、110 -> '110歳以上'"""
    return f"{AGE_MAX}歳以上" if age == AGE_MAX else f"{age}歳"


def age_range_label(age_min, age_max):
    """(18, 59) -> '18～59歳'、(65, 110) -> '65歳以上'、(0, 110) -> '全年齢（0歳～110歳以上）'"""
    if age_min == 0 and age_max == AGE_MAX:
        return f"全年齢（0歳～{AGE_MAX}歳以上）"
    if age_max == AGE_MAX:
        return f"{age_min}歳以上"
    if age_min == age_max:
        return f"{age_min}歳"
    return f"{age_min}～{age_max}歳"


def raw_table(model, flag_japanese):
    """ローデータ1シート分：性別ごとに 全国 → 47都道府県（コード順） → 主要エリア の行を並べる"""
    df = model.df_age[model.df_age["国籍総数か日本人"] == NATIONALITY[flag_japanese]]
    df = df.assign(
        _sex=df["sex"].map({s: i for i, s in enumerate(SEX_LABELS)}),
        _kubun=df["区分"].map({k: i for i, k in enumerate(AREA_KUBUN)}),
        _pos=range(len(df)),
    ).sort_values(["_sex", "_kubun", "_pos"])

    head = pd.DataFrame({
        "区分": df["区分"].map(AREA_KUBUN),
        "地域コード": df["地域コード"].fillna(""),
        "地域名": df["地域名"],
        "性別": df["sex"].map(SEX_LABELS),
        "総数": df["all"],
        "年齢不詳": df["不詳"],
        "年齢不詳を除く計": df[AGES].sum(axis=1),
    })
    ages = df[AGES].rename(columns=age_label)
    return pd.concat([head, ages], axis=1).reset_index(drop=True)


def raw_tables(model):
    """{'日本人のみ': DataFrame, '外国人含む': DataFrame}"""
    return {label: raw_table(model, flag) for flag, label in NAT_LABELS.items()}


def _area_notes():
    return [(f"主要エリア：{name}", "、".join(prefs)) for name, prefs in REGIONS.items()]


def raw_notes(date_text):
    jp, all_ = NAT_LABELS[1], NAT_LABELS[0]
    return [
        ("内容", "全国・47都道府県・主要エリア別、男女別、年齢（各歳）別の人口"),
        (f"シート「{jp}」", "日本人の人口（e-Statの「うち日本人」）"),
        (f"シート「{all_}」", "国籍を問わない総人口（e-Statの「国籍総数」）"),
        ("行の並び", "性別（男女計 → 男 → 女）ごとに、全国 → 47都道府県（地域コード順） → 主要エリア"),
        ("総数", "年齢不詳を含む人口"),
        ("年齢不詳", "年齢が不詳の人口。各歳（0歳～110歳以上）の列には含まれない"),
        ("年齢不詳を除く計", "0歳～110歳以上の合計（＝総数－年齢不詳）。アプリの条件別集計と同じ基準"),
        (f"{AGE_MAX}歳以上", f"{AGE_MAX}歳以上の人口の合計"),
        ("区分", "全国／都道府県／主要エリア。主要エリアは都道府県の合算のため、全国や主要エリアの行を含めたまま合計すると"
               "二重に数えることになる（区分の列でフィルタして合計すること）"),
        *_area_notes(),
        ("数値", "原数値の「-」（該当なし）は0としている"),
        ("出典", "政府統計の総合窓口（e-Stat）「令和7年国勢調査 人口等基本集計」"
               "第2-1表 男女、年齢（各歳）、国籍総数か日本人別人口"),
        ("URL", CENSUS_URL),
        ("出力日", date_text),
    ]


def summary_notes(date_text):
    return [
        ("人口", "年齢不詳の人口は含まない。「年齢不詳（集計外・参考）」列は、各条件の地域・性別・国籍での年齢不詳の人数"),
        ("年齢", f"{AGE_MAX}歳は「{AGE_MAX}歳以上」の合計"),
        ("国籍", f"{NAT_LABELS[1]}＝e-Statの「うち日本人」、{NAT_LABELS[0]}＝「国籍総数」（国籍を問わない総人口）"),
        ("世帯数", "一般世帯と施設等の世帯の合計（総数）"),
        ("地域", "複数の地域を選んだ場合は合算。主要エリアの内訳は以下のとおり"),
        *_area_notes(),
        ("出典", "政府統計の総合窓口（e-Stat）「令和7年国勢調査 人口等基本集計」"
               "第1-1表（世帯数）、第2-1表（男女、年齢（各歳）、国籍総数か日本人別人口）"),
        ("URL", CENSUS_URL),
        ("出力日", date_text),
    ]


def _formats(book):
    return {
        "header": book.add_format({
            "bold": True, "bg_color": HEADER_COLOR, "border": 1,
            "align": "center", "valign": "vcenter", "text_wrap": True,
        }),
        "number": book.add_format({"num_format": "#,##0"}),
        "wrap": book.add_format({"text_wrap": True, "valign": "top"}),
    }


def _write_table(writer, fmt, sheet, df, number_cols, widths, freeze_cols=0):
    """見出しの装飾・桁区切り・列幅・ウィンドウ枠の固定・フィルタをつけて1シート書き出す"""
    df.to_excel(writer, sheet_name=sheet, index=False)
    ws = writer.sheets[sheet]
    for c, name in enumerate(df.columns):
        ws.write(0, c, name, fmt["header"])
        ws.set_column(c, c, widths.get(name, 10), fmt["number"] if name in number_cols else None)
    ws.set_row(0, 32)
    ws.freeze_panes(1, freeze_cols)
    ws.autofilter(0, 0, len(df), len(df.columns) - 1)


def _write_notes(writer, fmt, notes):
    df = pd.DataFrame(notes, columns=["項目", "内容"])
    df.to_excel(writer, sheet_name=NOTES_SHEET, index=False)
    ws = writer.sheets[NOTES_SHEET]
    for c, name in enumerate(df.columns):
        ws.write(0, c, name, fmt["header"])
    ws.set_column(0, 0, 24, fmt["wrap"])
    ws.set_column(1, 1, 110, fmt["wrap"])


def build_raw_workbook(model, date_text):
    """ローデータのExcel（bytes）"""
    number_cols = set(RAW_TOTAL_COLS) | {age_label(a) for a in AGES}
    widths = {"区分": 11, "地域コード": 10, "地域名": 14, "性別": 8, "総数": 13, "年齢不詳": 11, "年齢不詳を除く計": 13}
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        fmt = _formats(writer.book)
        for sheet, df in raw_tables(model).items():
            _write_table(writer, fmt, sheet, df, number_cols, widths, freeze_cols=len(RAW_KEY_COLS))
        _write_notes(writer, fmt, raw_notes(date_text))
    return output.getvalue()


def build_summary_workbook(df, number_cols, date_text):
    """条件別集計の一覧のExcel（bytes）"""
    widths = {"集計番号": 9, "地域": 40, "国籍": 12, "性別": 8, "年齢": 22}
    widths |= {c: 16 for c in number_cols}
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        fmt = _formats(writer.book)
        _write_table(writer, fmt, "集計結果", df, set(number_cols), widths)
        _write_notes(writer, fmt, summary_notes(date_text))
    return output.getvalue()
