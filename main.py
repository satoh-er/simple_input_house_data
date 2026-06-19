# -*- coding: utf-8 -*-
"""
main.py
==============================================================================
代表ケースについて暖冷房負荷モデルを構築し、入力値の再現性を検証して、
heat_load_calc 入力 JSON と一覧 Excel（モデル単位）を出力する。

代表3ケース（戸建・基礎断熱／戸建・床断熱／集合住宅）に加え、入力により
室数が少なくなる縮退ケース（3.4.2 ただし書き：床面積0の室を除く）も対象とする。
==============================================================================
"""
import os
from simple_input import HeatLoadModel
import verify as vfy
import hlc_builder as hb
import export_excel as ex

OUT = "./outputs"
os.makedirs(OUT, exist_ok=True)

CASES = [
    # ---- 標準3ケース（3室 または 4室）--------------------------------------
    dict(label="戸建_基礎断熱", building_type="detached",
         A_MR=29.81, A_OR=51.35, A_A=120.08, region=6, A_env=307.51,
         UA=0.87, eta_AC=2.8, eta_AH=4.3, is_floor_ins=False),          # 4室
    dict(label="戸建_床断熱", building_type="detached",
         A_MR=29.81, A_OR=51.35, A_A=120.08, region=6, A_env=307.51,
         UA=0.87, eta_AC=2.8, eta_AH=4.3, is_floor_ins=True),           # 3室
    dict(label="集合住宅", building_type="apartment",
         A_MR=24.23, A_OR=29.75, A_A=70.00, region=6, A_env=238.22,
         UA=1.20, eta_AC=2.8, eta_AH=4.3, is_floor_ins=True),           # 3室

    # ---- 室数が少なくなる縮退ケース（3.4.2 ただし書き：床面積0の室を除く）----
    # 非居室 A_NR = max(A_A - A_MR - A_OR, 0) = 0 となり NR が除外される
    dict(label="戸建_床断熱_NR0_2室", building_type="detached",
         A_MR=60.00, A_OR=60.08, A_A=120.08, region=6, A_env=307.51,
         UA=0.87, eta_AC=2.8, eta_AH=4.3, is_floor_ins=True),           # MR/OR の2室
    dict(label="戸建_基礎断熱_NR0_3室", building_type="detached",
         A_MR=60.00, A_OR=60.08, A_A=120.08, region=6, A_env=307.51,
         UA=0.87, eta_AC=2.8, eta_AH=4.3, is_floor_ins=False),          # MR/OR/UF の3室
    # その他居室 A_OR = 0 で OR が除外される
    dict(label="集合住宅_OR0_2室", building_type="apartment",
         A_MR=40.00, A_OR=0.00, A_A=70.00, region=6, A_env=238.22,
         UA=1.20, eta_AC=2.8, eta_AH=4.3, is_floor_ins=True),           # MR/NR の2室
    # 主たる居室のみ（OR=0 かつ A_A=A_MR で NR=0）
    dict(label="戸建_床断熱_MR単室_1室", building_type="detached",
         A_MR=120.08, A_OR=0.00, A_A=120.08, region=6, A_env=307.51,
         UA=0.87, eta_AC=2.8, eta_AH=4.3, is_floor_ins=True),           # MR の1室
    # 主たる居室 A_MR = 0 で MR が除外される（内壁床は生存室へ付け替え: 3.4.3.9.2）
    dict(label="戸建_基礎断熱_MR0_3室", building_type="detached",
         A_MR=0.00, A_OR=51.35, A_A=120.08, region=6, A_env=307.51,
         UA=0.87, eta_AC=2.8, eta_AH=4.3, is_floor_ins=False),          # OR/NR/UF の3室
    dict(label="戸建_床断熱_MR0_2室", building_type="detached",
         A_MR=0.00, A_OR=51.35, A_A=120.08, region=6, A_env=307.51,
         UA=0.87, eta_AC=2.8, eta_AH=4.3, is_floor_ins=True),           # OR/NR の2室
]


def run():
    for c in CASES:
        label = c.pop("label")
        m = HeatLoadModel(**c)
        results = vfy.verify(m, tol=0.5)
        print("=" * 80)
        print(f"■ {label}")
        print(vfy.format_report(results))

        hlc = hb.build(m)
        jpath = os.path.join(OUT, f"model_{label}.json")
        hb.to_json(hlc, jpath)
        xpath = os.path.join(OUT, f"model_{label}.xlsx")
        ex.export(m, results, xpath, label)
        rooms = [r["name"] for r in hlc["rooms"]]
        print(f"  境界数={len(hlc['boundaries'])}, 室数={len(hlc['rooms'])} {rooms}")
        print(f"  出力: {os.path.basename(jpath)}, {os.path.basename(xpath)}")
        c["label"] = label


if __name__ == "__main__":
    run()
