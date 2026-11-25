"""
令和7年国勢調査 人口・世帯数 集計アプリ

    streamlit run src/app.py
"""
import re
from copy import copy
from datetime import datetime, timedelta, timezone

import pandas as pd
import streamlit as st

from config import AGE_MAX, CENSUS_URL, ESTAT_URL, LOGO_FILE, NAT_LABELS, REGIONS, SEX_LABELS, UPDATE_HISTORY
from export import RAW_KEY_COLS, age_range_label, build_raw_workbook, build_summary_workbook, raw_tables
from model import Population

MAX_CONDITIONS = 20
UNITS = {"人口": "人", "世帯数": "世帯"}  # 集計対象 → 単位
DEFAULTS = {"area": [], "nat": 1, "sex": "男女", "age_min": 18, "age_max": 59}  # 1条件ぶんの入力欄の初期値
JST = timezone(timedelta(hours=9))
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

CSS = """
<style>
.block-container, [data-testid="stMainBlockContainer"] { padding-top: 2.2rem; }
[data-testid="stMetricValue"] { font-variant-numeric: tabular-nums; }
</style>
"""


# ---------------------------------------------------------------- データ

@st.cache_resource
def load_model():
    return Population()


@st.cache_data(show_spinner="ローデータを作成しています…")
def raw_workbook():
    return build_raw_workbook(load_model())


@st.cache_data
def raw_preview():
    return raw_tables(load_model())


# ---------------------------------------------------------------- 条件の状態管理
# 条件ごとに ID を振り、入力欄の値は session_state の "{項目}_{ID}" に持つ（追加・複製・削除しても値がずれない）

def _key(field, cid):
    return f"{field}_{cid}"


def init_state():
    ss = st.session_state
    if "cond_ids" not in ss:
        ss.cond_ids = [0]
        ss.next_id = 1
    for cid in ss.cond_ids:
        for field, default in DEFAULTS.items():
            key = _key(field, cid)
            # 世帯数モードで入力欄が非表示になっても値が消えないよう、毎回書き戻す
            ss[key] = ss[key] if key in ss else copy(default)


def add_condition():
    ss = st.session_state
    if len(ss.cond_ids) >= MAX_CONDITIONS:
        return
    ss.cond_ids.append(ss.next_id)
    ss.next_id += 1


def duplicate_condition(cid):
    ss = st.session_state
    if len(ss.cond_ids) >= MAX_CONDITIONS:
        return
    new = ss.next_id
    ss.next_id += 1
    for field in DEFAULTS:
        ss[_key(field, new)] = copy(ss[_key(field, cid)])
    ss.cond_ids.insert(ss.cond_ids.index(cid) + 1, new)


def _drop_keys(cid):
    for field in DEFAULTS:
        st.session_state.pop(_key(field, cid), None)


def remove_condition(cid):
    st.session_state.cond_ids.remove(cid)
    _drop_keys(cid)


def reset_conditions():
    ss = st.session_state
    for cid in ss.cond_ids:
        _drop_keys(cid)
    ss.cond_ids = [ss.next_id]
    ss.next_id += 1


def read_condition(cid):
    return {field: st.session_state[_key(field, cid)] for field in DEFAULTS}


# ---------------------------------------------------------------- 集計

def evaluate(model, target, cond):
    """{'value': 集計値, 'unknown': 年齢不詳の人数（人口のみ）, 'issues': 入力の不備}"""
    issues = []
    if not cond["area"]:
        issues.append("地域を選択してください")
    if target == "人口" and cond["age_min"] > cond["age_max"]:
        issues.append("年齢の下限が上限を超えています")
    if issues:
        return {"value": None, "unknown": None, "issues": issues}

    if target == "人口":
        value = model.age_area_population(cond["age_min"], cond["age_max"], cond["area"], cond["sex"], cond["nat"])
        unknown = model.unknown_age_count(cond["nat"], cond["area"], cond["sex"])
    else:
        value, unknown = model.household_count(cond["area"]), None
    return {"value": value, "unknown": unknown, "issues": []}


def summary_frame(target, items):
    """集計結果一覧（画面表示・Excel出力用）と、数値列の名前。items は [(集計番号, 条件, 結果), ...]"""
    value_col = f"{target}（{UNITS[target]}）"
    unknown_col = "年齢不詳（集計外・参考）"
    rows = []
    for n, cond, res in items:
        row = {"集計番号": n, "地域": "、".join(cond["area"]) or "（未選択）"}
        if target == "人口":
            row["国籍"] = NAT_LABELS[cond["nat"]]
            row["性別"] = SEX_LABELS[cond["sex"]]
            row["年齢"] = age_range_label(cond["age_min"], cond["age_max"])
        row[value_col] = res["value"]
        if target == "人口":
            row[unknown_col] = res["unknown"]
        rows.append(row)
    df = pd.DataFrame(rows)
    number_cols = [value_col, unknown_col] if target == "人口" else [value_col]
    df[number_cols] = df[number_cols].astype("Int64")
    return df, number_cols


def clean_filename(name, default):
    name = re.sub(r'[\\/:*?"<>|]', "_", name.strip()) or default
    return name if name.lower().endswith(".xlsx") else name + ".xlsx"


# ---------------------------------------------------------------- 画面

def render_header():
    st.image(str(LOGO_FILE), width=260)
    st.markdown("### 令和7年 人口＆世帯数 集計")
    st.caption("令和7年国勢調査（人口等基本集計）の結果から、地域・国籍・性別・年齢を指定して人口と世帯数を集計します。")


def render_area_notes(model, area):
    regions = [a for a in area if a in REGIONS]
    if regions:
        st.caption("  \n".join(f"{r}：{'、'.join(REGIONS[r])}" for r in regions))
    overlaps = model.overlapping_prefectures(area)
    if overlaps:
        st.warning(f"地域が重複しています（{'、'.join(overlaps)}）。重複分は二重に集計されます。", icon=":material/warning:")


def render_result(target, res):
    unit = UNITS[target]
    value = res["value"]
    if value is None:
        st.metric(target, "—")
        if "地域を選択してください" in res["issues"]:
            st.info("地域を選択すると、ここに結果が表示されます。", icon=":material/arrow_back:")
        return

    st.metric(target, f"{value:,} {unit}", help="下の欄の右端のアイコンで、カンマなしの数値をコピーできます")
    if value >= 10_000:
        st.caption(f"約 {value / 10_000:,.1f} 万{unit}")
    st.code(str(value), language=None)
    if res["unknown"] is not None:
        st.caption(f"年齢不詳（集計外）：{res['unknown']:,} 人")


def render_condition(model, target, n, cid, area_list, can_add, can_remove):
    cond = read_condition(cid)
    res = evaluate(model, target, cond)

    with st.container(border=True):
        title, dup, rm = st.columns([8, 1.2, 1.2], vertical_alignment="center")
        title.markdown(f"**集計 {n}**")
        dup.button("複製", key=f"dup_{cid}", icon=":material/content_copy:", on_click=duplicate_condition, args=(cid,),
                   disabled=not can_add, help="この条件をコピーして、すぐ下に追加します", use_container_width=True)
        rm.button("削除", key=f"del_{cid}", icon=":material/delete:", on_click=remove_condition, args=(cid,),
                  disabled=not can_remove, use_container_width=True)

        left, right = st.columns([3, 2], gap="large")
        with left:
            st.multiselect("地域", area_list, key=_key("area", cid), placeholder="地域を選択（複数選ぶと合算）")
            render_area_notes(model, cond["area"])
            if target == "人口":
                c1, c2 = st.columns(2)
                c1.radio("国籍", list(NAT_LABELS), format_func=NAT_LABELS.get, horizontal=True, key=_key("nat", cid))
                c2.radio("性別", list(SEX_LABELS), format_func=SEX_LABELS.get, horizontal=True, key=_key("sex", cid))
                c3, c4 = st.columns(2)
                c3.number_input("年齢（下限）", min_value=0, max_value=AGE_MAX, step=1, key=_key("age_min", cid))
                c4.number_input("年齢（上限）", min_value=0, max_value=AGE_MAX, step=1, key=_key("age_max", cid),
                                help=f"{AGE_MAX} にすると {AGE_MAX}歳以上 をすべて含みます")
                if cond["age_min"] > cond["age_max"]:
                    st.error("年齢の下限が上限を超えています。", icon=":material/error:")
        with right:
            render_result(target, res)
    return cond, res


def render_summary(target, conds, results, now):
    st.markdown("#### 集計結果一覧")
    numbered = list(zip(range(1, len(conds) + 1), conds, results))
    done = [item for item in numbered if not item[2]["issues"]]
    pending = [f"集計 {n}：{'・'.join(res['issues'])}" for n, _, res in numbered if res["issues"]]

    if not done:
        st.caption("条件を入力すると、ここに一覧が表示され、Excelに出力できます。")
        return
    if pending:
        st.info("次の条件は入力が終わっていないため、一覧とExcelには含まれません。\n\n" + "\n".join(f"- {p}" for p in pending),
                icon=":material/edit_note:")

    df, number_cols = summary_frame(target, done)
    st.dataframe(
        df, hide_index=True, use_container_width=True,
        column_config={c: st.column_config.NumberColumn(format="localized") for c in number_cols},
    )

    default_name = f"令和7年国勢調査_{target}集計_{now:%y%m%d}.xlsx"
    name_col, btn_col = st.columns([3, 1], vertical_alignment="bottom")
    filename = name_col.text_input("ファイル名", value=default_name, help="拡張子（.xlsx）は自動で付きます")
    btn_col.download_button(
        "Excelでダウンロード", data=build_summary_workbook(df, number_cols, f"{now:%Y/%m/%d}"),
        file_name=clean_filename(filename, default_name), mime=XLSX_MIME,
        type="primary", icon=":material/download:", on_click="ignore", key="download_summary",
        use_container_width=True,
    )


def render_calc_tab(model, now):
    st.caption("条件を入力すると、結果がすぐに表示されます。「条件を追加」「複製」で最大20件まで並べて比較でき、一覧はExcelに出力できます。")
    target = st.radio("集計対象", list(UNITS), horizontal=True, key="target")
    area_list = model.area_list(flag_household=0 if target == "人口" else 1)

    ids = list(st.session_state.cond_ids)
    can_add = len(ids) < MAX_CONDITIONS
    conds, results = [], []
    for n, cid in enumerate(ids, start=1):
        cond, res = render_condition(model, target, n, cid, area_list, can_add, can_remove=len(ids) > 1)
        conds.append(cond)
        results.append(res)

    add_col, reset_col, _ = st.columns([1.4, 1.4, 6])
    add_col.button("条件を追加", key="add_condition", icon=":material/add:", on_click=add_condition,
                   disabled=not can_add, use_container_width=True)
    reset_col.button("すべてリセット", key="reset_conditions", icon=":material/restart_alt:",
                     on_click=reset_conditions, type="tertiary", use_container_width=True)
    if not can_add:
        st.caption(f"条件は最大{MAX_CONDITIONS}件までです。")

    st.divider()
    render_summary(target, conds, results, now)
    st.caption("※ 年齢不詳の人口は集計に含みません。定義や出典は「注記・出典」タブをご覧ください。")


def render_raw_tab(model):
    st.markdown("#### 都道府県別・男女年齢各歳の人口（ローデータ）")
    st.markdown(
        f"全国・47都道府県・主要エリアの人口を、男女別・年齢各歳（0歳～{AGE_MAX}歳以上）で1つのExcelファイルに出力します。  \n"
        f"シートは **{NAT_LABELS[1]}** と **{NAT_LABELS[0]}** に分かれていて、どちらにも **年齢不詳** の人数を載せています。"
    )

    cols = st.columns(4)
    for col, flag in zip(cols[:2], NAT_LABELS):
        col.metric(f"総数（全国・{NAT_LABELS[flag]}）", f"{model.total_count(flag):,} 人")
    for col, flag in zip(cols[2:], NAT_LABELS):
        col.metric(f"うち年齢不詳（全国・{NAT_LABELS[flag]}）", f"{model.unknown_age_count(flag):,} 人")

    st.download_button(
        "ローデータをダウンロード（Excel）", data=raw_workbook(),
        file_name="令和7年国勢調査_都道府県別_男女年齢各歳人口.xlsx", mime=XLSX_MIME,
        type="primary", icon=":material/download:", on_click="ignore", key="download_raw",
    )

    with st.expander("出力内容（シート・列の説明）"):
        st.markdown(f"""
| シート | 内容 |
|---|---|
| {NAT_LABELS[1]} | 日本人の人口（e-Statの「うち日本人」） |
| {NAT_LABELS[0]} | 国籍を問わない総人口（e-Statの「国籍総数」） |
| 注記 | 列の定義・主要エリアの内訳・出典 |

| 列 | 内容 |
|---|---|
| 区分 | 全国／都道府県／主要エリア（主要エリアは都道府県の合算。合計するときは区分でフィルタ） |
| 地域コード・地域名・性別 | 性別（男女計 → 男 → 女）ごとに、全国 → 47都道府県 → 主要エリアの順 |
| 総数 | 年齢不詳を含む人口 |
| 年齢不詳 | 年齢が不詳の人口（各歳の列には含まれない） |
| 年齢不詳を除く計 | 0歳～{AGE_MAX}歳以上の合計（＝総数－年齢不詳。条件別集計と同じ基準） |
| 0歳 … {AGE_MAX}歳以上 | 各歳の人口（{AGE_MAX}歳以上はまとめて1列） |
""")

    st.markdown("##### プレビュー")
    c1, c2 = st.columns(2)
    flag = c1.radio("国籍", list(NAT_LABELS), format_func=NAT_LABELS.get, horizontal=True, key="preview_nat")
    sex = c2.radio("性別", list(SEX_LABELS.values()), horizontal=True, key="preview_sex")
    df = raw_preview()[NAT_LABELS[flag]]
    df = df[df["性別"] == sex]
    st.dataframe(
        df, hide_index=True, height=420, use_container_width=True,
        column_config={
            c: (st.column_config.Column(pinned=True) if c in RAW_KEY_COLS
                else st.column_config.NumberColumn(format="localized"))
            for c in df.columns
        },
    )


def render_notes_tab(model):
    region_lines = "\n".join(f"- **{name}**：{'、'.join(prefs)}" for name, prefs in REGIONS.items())
    update_lines = "\n".join(f"- {line}" for line in UPDATE_HISTORY)
    st.markdown(f"""
#### 出典
- 政府統計の総合窓口（[e-Stat]({ESTAT_URL})）「[令和7年(2025年)国勢調査 人口等基本集計]({CENSUS_URL})」
    - 第1-1表（世帯数）
    - 第2-1表（男女、年齢（各歳）、国籍総数か日本人別人口）

#### 集計の定義
- **年齢不詳**：条件別集計の人口には含みません。全国の年齢不詳は **{model.unknown_age_count(0):,}人**（うち日本人 {model.unknown_age_count(1):,}人）です。各条件の年齢不詳の人数は、結果欄と一覧に参考として表示しています。
- **年齢の上限**：{AGE_MAX}歳は「{AGE_MAX}歳以上」を表します（上限を {AGE_MAX} にすると {AGE_MAX}歳以上 の全人口を含みます）。
- **国籍**：{NAT_LABELS[1]}＝e-Statの「うち日本人」、{NAT_LABELS[0]}＝「国籍総数」（国籍を問わない総人口）。
- **世帯数**：一般世帯と施設等の世帯の合計（総数）です。
- **複数地域**：選んだ地域を合算します。主要エリアとその構成都道府県を同時に選ぶと、二重に数えます（画面に警告が出ます）。
- 原数値の「-」（該当なし）は 0 として扱っています。

#### 主要エリアの内訳
{region_lines}

#### 更新情報
{update_lines}

#### データの利用について
使用データはすべて政府統計（e-Stat）から取得しており、政府の定める利用規約に従って利用しています。詳しくは [e-Stat利用規約](https://www.e-stat.go.jp/terms-of-use) をご確認ください。
""")


def main():
    st.set_page_config(page_title="令和7年国勢調査 人口集計", page_icon="📊", layout="wide")
    st.markdown(CSS, unsafe_allow_html=True)

    model = load_model()
    init_state()
    now = datetime.now(JST)

    render_header()
    tab_calc, tab_raw, tab_notes = st.tabs([
        ":material/calculate: 条件別集計",
        ":material/table_view: ローデータ出力",
        ":material/info: 注記・出典",
    ])
    with tab_calc:
        render_calc_tab(model, now)
    with tab_raw:
        render_raw_tab(model)
    with tab_notes:
        render_notes_tab(model)


if __name__ == "__main__":
    main()
