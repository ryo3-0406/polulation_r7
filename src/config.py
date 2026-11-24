from pathlib import Path

# ディレクトリ（実行時のカレントディレクトリに依存しないよう、このファイルの位置から解決）
ROOT_DIR = Path(__file__).resolve().parent.parent
RAW_DIR = ROOT_DIR / "data" / "raw"
PREPROCESSED_DIR = ROOT_DIR / "data" / "preprocessed"
LOGO_FILE = ROOT_DIR / "assets" / "census2025_logo.png"

# e-Statからダウンロードした原数値ファイル
RAW_HOUSEHOLD_FILE = RAW_DIR / "b01_01.xlsx"  # 第1-1表 世帯の種類別世帯数 等
RAW_AGE_FILE = RAW_DIR / "b02_01.xlsx"  # 第2-1表 男女、年齢（各歳）、国籍総数か日本人別人口

# 前処理済みファイル
AGE_FILE = PREPROCESSED_DIR / "R7_pop_age.csv"
HOUSEHOLD_FILE = PREPROCESSED_DIR / "R7_pop_setai.csv"

AGE_MAX = 110  # 110 は「110歳以上」

# 主要エリアの定義（前処理でこの都道府県を合算する）
REGIONS = {
    "関東(1都6県)": ["茨城県", "栃木県", "群馬県", "埼玉県", "千葉県", "東京都", "神奈川県"],
    "東海(3県)": ["岐阜県", "愛知県", "三重県"],
    "関西(2府4県)": ["滋賀県", "京都府", "大阪府", "兵庫県", "奈良県", "和歌山県"],
    "福岡(2県)": ["福岡県", "佐賀県"],
}

# 画面・Excelでの表示名
NAT_LABELS = {1: "日本人のみ", 0: "外国人含む"}  # flag_japanese → 表示名（0 は e-Stat の「国籍総数」）
SEX_LABELS = {"男女": "男女計", "男": "男", "女": "女"}

# 出典
ESTAT_URL = "https://www.e-stat.go.jp/"
CENSUS_URL = (
    "https://www.e-stat.go.jp/stat-search/files?page=1&layout=datalist&toukei=00200521&tstat=000001230925"
    "&cycle=0&year=20250&month=24101210&tclass1=000001230926&tclass2=000001232187"
)

# 更新情報（新しい順）
UPDATE_HISTORY = [
    "2026/11/24(火)：都道府県別・男女年齢各歳のローデータ出力を追加、画面を改善",
    "2026/10/05(月)：令和7年国勢調査の結果に更新",
]
