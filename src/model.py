import pandas as pd

from config import AGE_FILE, AGE_MAX, HOUSEHOLD_FILE, REGIONS

NATIONALITY = {1: "1_うち日本人", 0: "0_国籍総数"}


class Population:
    def __init__(self, age_file=AGE_FILE, household_file=HOUSEHOLD_FILE):
        df_age = pd.read_csv(age_file, dtype={"地域コード": str}, encoding="utf-8-sig")
        df_age = df_age.rename(columns={str(a): a for a in range(AGE_MAX + 1)})  # 年齢列は int で扱う

        df_setai = pd.read_csv(household_file, dtype={"地域コード": str}, encoding="utf-8-sig")

        self.df_age = df_age
        self.df_setai = df_setai

    def age_area_population(self, age_min, age_max, area, sex, flag_japanese=1):
        if not 0 <= age_min <= age_max <= AGE_MAX:
            raise ValueError(f"年齢の範囲が不正です: {age_min}～{age_max}")

        df = self._age_rows(area, sex, flag_japanese)

        result = df.loc[:, age_min:age_max].to_numpy().sum()

        return int(result)

    def household_count(self, area):
        df = self.df_setai
        df = df[df["地域名"].isin(area)]

        result = df["0_総数"].sum()

        return int(result)

    def unknown_age_count(self, flag_japanese=0, area=("全国",), sex="男女"):
        """年齢不詳人口（年齢別の集計からは除外される人数）"""
        return int(self._age_rows(area, sex, flag_japanese)["不詳"].sum())

    def total_count(self, flag_japanese=0, area=("全国",), sex="男女"):
        """総数（年齢不詳を含む人口）"""
        return int(self._age_rows(area, sex, flag_japanese)["all"].sum())

    def _age_rows(self, area, sex, flag_japanese):
        df = self.df_age
        df = df[df["地域名"].isin(area)]
        df = df[df["sex"] == sex]
        df = df[df["国籍総数か日本人"] == NATIONALITY[flag_japanese]]
        return df

    def area_list(self, flag_household=0):

        if flag_household:
            df = self.df_setai
        else:
            df = self.df_age

        area_list = df["地域名"].unique().tolist()

        return area_list

    @staticmethod
    def overlapping_prefectures(area):
        """選択地域の中で二重に数えられる都道府県（例：関東 と 東京都 → ['東京都']）"""
        if "全国" in area and len(area) > 1:
            return ["全国"]
        prefs = [p for a in area for p in REGIONS.get(a, [a])]
        return sorted({p for p in prefs if prefs.count(p) > 1})
