# 令和7年 国勢調査 人口集計アプリ

令和7年（2025年）国勢調査の結果から、都道府県・主要エリア別の人口（年齢・性別・国籍条件付き）と世帯数を集計するStreamlitアプリです。

## 使用データ

e-Stat「令和7年国勢調査 人口等基本集計」の原数値Excelを `data/raw` に配置しています。

| ファイル | 表 | 用途 |
|---|---|---|
| `b01_01.xlsx` | 第1-1表 | 世帯数（`0_総数`＝一般世帯＋施設等の世帯） |
| `b02_01.xlsx` | 第2-1表 男女、年齢（各歳）、国籍総数か日本人別人口 | 人口 |

## ディレクトリ構成

```bash
py_population_r7/
├── assets/
│   └── census2025_logo.png  # 画面上部に表示するロゴ
├── data/
│   ├── raw/             # e-Statの原数値Excel（手を加えない）
│   └── preprocessed/    # preprocess.py が出力するCSV
├── src/
│   ├── config.py        # パス・主要エリアの定義
│   ├── preprocess.py    # raw → preprocessed
│   ├── model.py         # 集計ロジック
│   └── app.py           # Streamlit画面
├── tests/
│   └── test_population.py
└── requirements.txt
```

## 実行方法

```bash
pip install -r requirements.txt

# 前処理（rawのExcelを差し替えたときに実行）
python src/preprocess.py

# アプリ起動
streamlit run src/app.py

# テスト（アプリの集計値と原数値Excelの単純集計を突合）
python -m unittest discover -s tests -v
```

## データ処理

- 主要エリア（関東・東海・関西・福岡）は、`src/config.py` の `REGIONS` の定義に従って都道府県を合算して作成
- 原数値の「-」（該当なし）は0として扱う
- 前処理時に、全国＝47都道府県の合計、総数＝各歳＋年齢不詳、世帯数の総数＝内訳の合計 を確認し、不一致ならエラー
- 年齢（from）＞年齢（to）のときはエラー表示、地域の重複選択（例：関東＋東京都）は警告表示

## 注意事項

- 年齢上限を110歳に設定すると、110歳以上の全人口が含まれます。
- 年齢不詳の人口は計算に含めていません。
- 主要エリアの内訳は `src/config.py` の `REGIONS` を参照してください。

## 著作権・データ利用について

使用データは全て政府統計(e-Stat)より取得しており、政府の定める利用規約に従って利用しています。詳しくは [e-Stat利用規約](https://www.e-stat.go.jp/terms-of-use) をご確認ください。
