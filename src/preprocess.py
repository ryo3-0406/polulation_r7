"""
e-Statの原数値Excel（data/raw）から、アプリで使う前処理済みCSV（data/preprocessed）を作成する。

    python src/preprocess.py

- 全国・47都道府県の行のみを抽出
- 「-」（該当数値なし）は 0 に変換
- 主要エリア（config.REGIONS）は都道府県を合算して作成
- 全国 = 47都道府県の合計、総数 = 各歳 + 年齢不詳 などの整合性をチェック
"""
import pandas as pd

from config import (
    AGE_FILE,
    AGE_MAX,
    HOUSEHOLD_FILE,
    PREPROCESSED_DIR,
    RAW_AGE_FILE,
    RAW_HOUSEHOLD_FILE,
    REGIONS,
)

AGE_COLS = [str(age) for age in range(AGE_MAX + 1)]
HOUSEHOLD_COLS = ["0_総数", "1_一般世帯", "2_施設等の世帯"]
SEX_MAP = {"0_総数": "男女", "1_男": "男", "2_女": "女"}


def _find_row(df_raw, col, value):
    """原数値Excelのヘッダ行の位置を、ラベル文字列から探す"""
    hits = df_raw.index[df_raw[col] == value]
    if len(hits) != 1:
        raise ValueError(f"列{col}に「{value}」の行が見つかりません（{len(hits)}件）")
    return hits[0]


def _to_int(df):
    """「-」を0にして整数化する"""
    return df.map(lambda v: 0 if v == "-" else v).astype("int64")


def _split_code_name(s):
    """'01000_北海道' -> ('01000', '北海道')"""
    code, name = s.split("_", 1)
    return code, name


def _area_order(names):
    """全国 → 主要エリア → 都道府県（コード順）の並びにする"""
    return ["全国"] + list(REGIONS) + [n for n in names if n != "全国" and n not in REGIONS]


def _check_national_total(df, value_cols, keys):
    """全国 = 47都道府県の合計 になっているか"""
    pref = df[df["区分"] == "都道府県"].groupby(keys)[value_cols].sum()
    nat = df[df["区分"] == "全国"].set_index(keys)[value_cols]
    diff = (pref.sort_index() - nat.sort_index()).abs().to_numpy().sum()
    if diff != 0:
        raise ValueError(f"全国と都道府県合計が一致しません（差分 {diff}）")


def _add_regions(df, value_cols, keys):
    """都道府県の行を合算して主要エリアの行を作る"""
    prefs = df[df["区分"] == "都道府県"]
    missing = {p for members in REGIONS.values() for p in members} - set(prefs["地域名"])
    if missing:
        raise ValueError(f"エリア定義に存在しない都道府県があります: {missing}")

    rows = []
    for region, members in REGIONS.items():
        agg = prefs[prefs["地域名"].isin(members)].groupby(keys, sort=False)[value_cols].sum().reset_index()
        agg["地域コード"] = ""
        agg["地域名"] = region
        agg["区分"] = "エリア"
        rows.append(agg)
    return pd.concat([df] + rows, ignore_index=True)


def preprocess_age():
    df_raw = pd.read_excel(RAW_AGE_FILE, header=None, dtype=object)

    item_row = _find_row(df_raw, 3, "項目名")
    header_row = _find_row(df_raw, 0, "国籍総数か日本人")
    items = df_raw.loc[item_row].tolist()
    df = df_raw.loc[header_row + 1:].reset_index(drop=True)
    df = df[df[0].notna()]

    # 列の対応付け（'001_0歳'..'111_110歳以上' -> '0'..'110'）
    col_map = {0: "国籍総数か日本人", 1: "男女", 2: "地域識別コード", 3: "地域名_raw"}
    for i, item in enumerate(items):
        if not isinstance(item, str) or "_" not in item:
            continue
        num, label = item.split("_", 1)
        if num == "000":
            col_map[i] = "all"
        elif num.isdigit() and 1 <= int(num) <= AGE_MAX + 1:
            col_map[i] = str(int(num) - 1)
        elif "不詳" in label:
            col_map[i] = "不詳"
    df = df[list(col_map)].rename(columns=col_map)

    # 全国・都道府県のみ（地域識別コード 'a'）
    df = df[df["地域識別コード"] == "a"].copy()
    df[["地域コード", "地域名"]] = df["地域名_raw"].apply(lambda s: pd.Series(_split_code_name(s)))
    df["区分"] = df["地域コード"].map(lambda c: "全国" if c == "00000" else "都道府県")
    df["sex"] = df["男女"].map(SEX_MAP)

    value_cols = ["all"] + AGE_COLS + ["不詳"]
    df[value_cols] = _to_int(df[value_cols])

    # 総数 = 各歳 + 年齢不詳
    bad = df[df["all"] != df[AGE_COLS + ["不詳"]].sum(axis=1)]
    if not bad.empty:
        raise ValueError(f"総数と各歳合計が一致しません:\n{bad[['地域名', 'sex']]}")
    if df["地域名"].nunique() != 48:
        raise ValueError(f"全国+47都道府県になっていません（{df['地域名'].nunique()}件）")

    keys = ["国籍総数か日本人", "sex"]
    _check_national_total(df, value_cols, keys)
    df = _add_regions(df, value_cols, keys)

    order = _area_order(df["地域名"].unique().tolist())
    df["地域名"] = pd.Categorical(df["地域名"], categories=order, ordered=True)
    df["sex"] = pd.Categorical(df["sex"], categories=list(SEX_MAP.values()), ordered=True)
    df = df.sort_values(["国籍総数か日本人", "sex", "地域名"])
    df["地域名"] = df["地域名"].astype(str)
    df["sex"] = df["sex"].astype(str)

    return df[["国籍総数か日本人", "sex", "区分", "地域コード", "地域名"] + value_cols].reset_index(drop=True)


def preprocess_household():
    df_raw = pd.read_excel(RAW_HOUSEHOLD_FILE, header=None, dtype=object)

    table_row = _find_row(df_raw, 6, "表章項目")
    item_row = _find_row(df_raw, 6, "項目名")
    header_row = _find_row(df_raw, 0, "地域識別コード")
    tables = df_raw.loc[table_row].tolist()
    items = df_raw.loc[item_row].tolist()
    headers = df_raw.loc[header_row].tolist()

    col_map = {
        headers.index("地域識別コード"): "地域識別コード",
        headers.index("2025年_地域コード"): "地域コード",
        headers.index("地域名"): "地域名_raw",
    }
    for i, (table, item) in enumerate(zip(tables, items)):
        if table == "世帯数" and item in HOUSEHOLD_COLS:
            col_map[i] = item
    if set(HOUSEHOLD_COLS) - set(col_map.values()):
        raise ValueError("世帯数の列が見つかりません")

    df = df_raw.loc[header_row + 1:, list(col_map)].rename(columns=col_map)
    df = df[df["地域識別コード"] == "a"].copy()
    df["地域名"] = df["地域名_raw"].map(lambda s: _split_code_name(s)[1])
    df["区分"] = df["地域コード"].map(lambda c: "全国" if c == "00000" else "都道府県")
    df[HOUSEHOLD_COLS] = _to_int(df[HOUSEHOLD_COLS])

    if df["地域名"].nunique() != 48:
        raise ValueError(f"全国+47都道府県になっていません（{df['地域名'].nunique()}件）")
    bad = df[df["0_総数"] != df["1_一般世帯"] + df["2_施設等の世帯"]]
    if not bad.empty:
        raise ValueError(f"世帯数の総数と内訳が一致しません:\n{bad['地域名']}")

    df["_key"] = 0
    _check_national_total(df, HOUSEHOLD_COLS, ["_key"])
    df = _add_regions(df, HOUSEHOLD_COLS, ["_key"])

    df["地域名"] = pd.Categorical(df["地域名"], categories=_area_order(df["地域名"].unique().tolist()), ordered=True)
    df = df.sort_values("地域名")
    df["地域名"] = df["地域名"].astype(str)

    return df[["区分", "地域コード", "地域名"] + HOUSEHOLD_COLS].reset_index(drop=True)


def main():
    PREPROCESSED_DIR.mkdir(parents=True, exist_ok=True)

    # utf-8-sig: Excelで開いても文字化けしないように
    df_age = preprocess_age()
    df_age.to_csv(AGE_FILE, index=False, encoding="utf-8-sig")
    print(f"保存: {AGE_FILE}  ({len(df_age)}行)")

    df_setai = preprocess_household()
    df_setai.to_csv(HOUSEHOLD_FILE, index=False, encoding="utf-8-sig")
    print(f"保存: {HOUSEHOLD_FILE}  ({len(df_setai)}行)")


if __name__ == "__main__":
    main()
