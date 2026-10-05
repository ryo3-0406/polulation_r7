from model import Population
from config import AGE_MAX, LOGO_FILE, REGIONS
import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime, timedelta, timezone


@st.cache_resource
def load_model():
    return Population()


def get_japan_date_text():
    # JSTはUTC+9
    jst = timezone(timedelta(hours=9))
    now_jst = datetime.now(jst)
    return now_jst.strftime("%y%m%d")  # 西暦下2桁 + 月 + 日


def main():
    # ページ設定（最初のStreamlitコマンドである必要がある）
    st.set_page_config(page_title="令和7年国勢調査 人口集計", layout="wide")

    # モデルの初期化
    model = load_model()

    # メインタイトル（国勢調査2025ロゴ）
    st.image(str(LOGO_FILE), width=360)
    st.subheader("令和7年 人口＆世帯数 基本集計")
    st.sidebar.header("集計設定")

    # 集計対象の選択
    household_or_personal = st.sidebar.radio("【集計対象】", ["人口", "世帯数"], index=0)
    area_list = model.area_list(flag_household=0 if household_or_personal == "人口" else 1)

    # 集計数の入力
    number_of_aggregation = st.sidebar.number_input(
        "【集計数】",
        min_value=1,
        max_value=20,
        value=1,
        step=1
    )

    st.sidebar.markdown("---")

    # 集計ごとの条件と結果の格納リスト
    aggregations = []
    excel_rows = []  # Excel用データ
    errors = []

    for i in range(number_of_aggregation):
        with st.sidebar.expander(f"▼ 集計 {i + 1} の条件", expanded=True):
            # 各集計の条件を辞書にまとめる
            settings = {}
            selected_area = st.multiselect("地域", area_list, key=f"area_{i}")
            settings["地域"] = ", ".join(selected_area) if selected_area else "未選択"

            if not selected_area:
                errors.append(f"集計 {i + 1}：地域を1つ以上選択してください。")

            overlaps = model.overlapping_prefectures(selected_area)
            if overlaps:
                st.warning(f"地域が重複しています（{'、'.join(overlaps)}）。重複分は二重に集計されます。")

            if household_or_personal == "人口":
                flag_japanese_option = st.radio("国民種別", ["日本人のみ", "外国籍含む"], index=0, key=f"flag_japanese_{i}")
                flag_japanese = 1 if flag_japanese_option == "日本人のみ" else 0
                settings["国民種別"] = flag_japanese_option

                sex = st.radio("性別", ["男女", "男", "女"], index=0, key=f"sex_{i}")
                settings["性別"] = sex

                age_min = st.number_input("年齢（from）", min_value=0, max_value=AGE_MAX, value=18, key=f"age_min_{i}")
                age_max = st.number_input("年齢（to）", min_value=0, max_value=AGE_MAX, value=59, key=f"age_max_{i}")
                settings["年齢"] = f"{age_min}才～{age_max}才"

                if age_min > age_max:
                    errors.append(f"集計 {i + 1}：年齢（from）は年齢（to）以下にしてください。")
                    result = None
                elif selected_area:
                    result = model.age_area_population(
                        age_min=age_min,
                        age_max=age_max,
                        area=selected_area,
                        sex=sex,
                        flag_japanese=flag_japanese
                    )
                else:
                    result = None
            else:
                settings["集計対象"] = "世帯数"
                if selected_area:
                    result = model.household_count(area=selected_area)
                else:
                    result = None

            aggregations.append((i + 1, result, "人" if household_or_personal == "人口" else "世帯", settings))

            # Excel用データを追加
            excel_row = {
                '集計番号': i + 1,
                '集計対象': household_or_personal,
                '地域': settings.get('地域', '未選択'),
                '国民種別': settings.get('国民種別', '-'),
                '性別': settings.get('性別', '-'),
                '年齢範囲': settings.get('年齢', '-'),
                '結果': result
            }
            excel_rows.append(excel_row)


    # --- 結果表示 ---

    if st.button("集計する", type="primary"):
        if errors:
            for message in errors:
                st.error(message)
        else:
            for idx, value, unit, settings in aggregations:
                formatted_value = f"{value:,} {unit}"

                # 枠の外に「集計 ◯」のタイトルを出す
                st.markdown(
                    f"<h4 style='margin-top:0.1rem; margin-bottom:0.3rem;'>▼ 集計 {idx}</h4>",
                    unsafe_allow_html=True
                )

                # 結果だけをカード内にスタイリッシュに表示
                st.markdown(
                    f"""
                    <div style="
                        background-color:#f8f9fa;
                        padding: 1.5rem;
                        border-left: 6px solid #4A90E2;
                        border-radius: 8px;
                        font-size: 2.0rem;
                        font-weight: bold;
                        color: #222;
                        margin-bottom: 0.5rem;
                        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
                    ">
                        {formatted_value}
                    </div>
                    """,
                    unsafe_allow_html=True
                )

                # 設定内容を控えめに表示する例
                st.markdown(
                    "<div style='font-size:0.9rem; margin-top:0.5rem;'>設定内容:</div>",
                    unsafe_allow_html=True
                )
                for key, val in settings.items():
                    st.markdown(
                        f"<div style='font-size:0.8rem; margin-bottom:0.2rem;'>{key}: {val}</div>",
                        unsafe_allow_html=True
                    )

                st.caption("\n")
                st.caption("※コピー用※")
                st.code(f"{value}", language="python")
                st.caption("↑コピーする場合は上記ブロックの右端ボタンをクリック")

                st.markdown("---")

            # Excel出力
            st.markdown("""
            #### Excelでエクスポート
            """)
            df_excel = pd.DataFrame(excel_rows)

            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_excel.to_excel(writer, index=False, sheet_name="集計結果")

            excel_data = output.getvalue()

            date_text = get_japan_date_text()
            default_filename = "令和7年国勢調査_人口集計_" + date_text + ".xlsx"
            filename_input = st.text_input("※ファイル名を変更する場合はファイル名を入力してEnter（拡張子は.xlsxのみ有効）", value=default_filename)
            st.download_button(
                label="Excelでダウンロード",
                data=excel_data,
                file_name=f"{filename_input}",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            st.markdown("---")

    else:
        st.info("左のサイドバーで集計条件の設定が完了したら、「集計する」ボタンをクリックして下さい。")

    # 集計結果の後に備考を表示
    region_lines = "\n".join(
        f"    - **{name}**： {'、'.join(prefs)}" for name, prefs in REGIONS.items()
    )
    st.markdown(f"""
    <small>

    #### 出典
    - **政府統計の総合窓口**：[e-Stat](https://www.e-stat.go.jp/)
    - **国勢調査**： [令和7年(2025年) 人口等基本集計](https://www.e-stat.go.jp/stat-search/files?page=1&layout=datalist&toukei=00200521&tstat=000001230925&cycle=0&year=20250&month=24101210&tclass1=000001230926&tclass2=000001232187)
      **表番号**： [第1-1表（世帯数）](https://www.e-stat.go.jp/stat-search/files?page=1&layout=datalist&toukei=00200521&tstat=000001230925&cycle=0&year=20250&month=24101210&tclass1=000001230926&tclass2=000001232187) / [第2-1表（男女、年齢（各歳）、国籍総数か日本人別人口）](https://www.e-stat.go.jp/stat-search/files?page=1&layout=datalist&toukei=00200521&tstat=000001230925&cycle=0&year=20250&month=24101210&tclass1=000001230926&tclass2=000001232187)
    ---
    #### 人口と世帯数
    - 年齢不詳の**{model.unknown_age_count(flag_japanese=0):,}人**（うち日本人 {model.unknown_age_count(flag_japanese=1):,}人）は除外
    - 世帯数は「2_施設等の世帯」も含む総数で算出
    - 年齢上限は{AGE_MAX}才(年齢(to)を[{AGE_MAX}]に設定すると{AGE_MAX}才以上の全ての人口を表示)
    ---
    #### 主要エリアの内訳
{region_lines}
    ---
    #### 更新情報
    - **最終更新**：2026/10/05(月)：令和7年国勢調査の結果に更新

    </small>
    """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()
