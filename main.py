# -*- coding: utf-8 -*-
"""
main.py
==============================================================================
本スクリプトは2つのフェーズを実行する。

 フェーズ1（成果物生成）:
   代表3ケースについて暖冷房負荷モデルを構築し、入力値の再現性を検証して、
   heat_load_calc 入力 JSON と一覧 Excel（モデル単位）を出力する。

 フェーズ2（室数設定の検証 / 3.4.2）:
   test_room_count.py のテストケース（標準3ケース＋床面積0による室除外の
   縮退ケース）を実行し、room_id 連番・容積>0・床面積保存・入力再現性を確認する。
==============================================================================
"""
import os
from simple_input import HeatLoadModel
import verify as vfy
import hlc_builder as hb
import export_excel as ex
import test_room_count as trc

OUT = "/mnt/user-data/outputs"
os.makedirs(OUT, exist_ok=True)

# 成果物（JSON/Excel）を生成する代表ケース
CASES = [
    dict(label="戸建_基礎断熱", building_type="detached",
         A_MR=29.81, A_OR=51.35, A_A=120.08, region=6, A_env=307.51,
         UA=0.87, eta_AC=2.8, eta_AH=4.3, is_floor_ins=False),
    dict(label="戸建_床断熱", building_type="detached",
         A_MR=29.81, A_OR=51.35, A_A=120.08, region=6, A_env=307.51,
         UA=0.87, eta_AC=2.8, eta_AH=4.3, is_floor_ins=True),
    dict(label="集合住宅", building_type="apartment",
         A_MR=24.23, A_OR=29.75, A_A=70.00, region=6, A_env=238.22,
         UA=1.20, eta_AC=2.8, eta_AH=4.3, is_floor_ins=True),
]


def build_and_export():
    """フェーズ1: 代表3ケースの構築・再現性検証・JSON/Excel 出力。"""
    print("#" * 80)
    print("# フェーズ1: 代表ケースの構築・検証・出力")
    print("#" * 80)
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
        print(f"  境界数={len(hlc['boundaries'])}, 室数={len(hlc['rooms'])}")
        print(f"  出力: {os.path.basename(jpath)}, {os.path.basename(xpath)}")
        c["label"] = label


def run():
    # フェーズ1: 成果物生成
    build_and_export()
    # フェーズ2: 室数設定（3.4.2）の検証ケースを実行
    print()
    print("#" * 80)
    print("# フェーズ2: 室数設定（3.4.2）の検証")
    print("#" * 80)
    all_ok = trc.run()
    return all_ok


if __name__ == "__main__":
    run()
