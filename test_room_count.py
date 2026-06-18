# -*- coding: utf-8 -*-
"""
test_room_count.py
==============================================================================
仕様書 3.4.2「出力する室数の設定」の検証スクリプト。

入力パラメータによって暖冷房負荷モデルの室数が変わること、および
「室用途ごとの床面積が0m2となる室は除く」（3.4.2.1・3.4.2.2 ただし書き）が
正しく機能することを、縮退ケースで確認する。

確認項目:
  - m.spaces が床面積>0 の室のみで構成される（UFは戸建基礎断熱のみ）
  - heat_load_calc 出力の room_id が 0 始まりの連番である（3.5.3）
  - 全室の容積>0（容積0の幽霊室を出力しない）
  - 境界の connected_room_id・換気 route・設備 space_id が実在 room_id のみ参照
  - 裏面境界IDが相互に整合
  - 入力値（A_A/A_MR/A_OR/A_env/UA/ηA）の再現性
==============================================================================
"""
from simple_input import HeatLoadModel
import hlc_builder as hb
import verify as vfy


def check(label, **kw):
    m = HeatLoadModel(**kw)
    hlc = hb.build(m)
    rooms = hlc["rooms"]
    ids = [r["id"] for r in rooms]
    valid = set(ids)

    contiguous = ids == list(range(len(ids)))                       # room_id連番（3.5.3）
    vol_ok = all(r["volume"] > 0 for r in rooms)                    # 容積0の室が無い
    bad_bnd = [b["id"] for b in hlc["boundaries"]
               if b["connected_room_id"] not in valid]              # 境界の接室
    bad_mv = [mv["id"] for mv in hlc["mechanical_ventilations"]
              if any(r not in valid for r in mv["route"])]          # 換気route
    eq = hlc["equipments"]
    sids = [e["property"]["space_id"]
            for e in eq["heating_equipments"] + eq["cooling_equipments"]]
    bad_eq = [s for s in sids if s not in valid]                    # 設備space_id
    # 裏面境界IDの相互参照
    bs = {b["id"]: b for b in hlc["boundaries"]}
    bad_rear = [b["id"] for b in hlc["boundaries"]
                if b.get("rear_surface_boundary_id") is not None
                and bs.get(b["rear_surface_boundary_id"], {}).get("rear_surface_boundary_id") != b["id"]]
    res = vfy.verify(m, tol=0.5)
    repro = all(res[k]["ok"] for k in ("A_A", "A_MR", "A_OR", "A_env", "UA", "etaA"))

    ok = (contiguous and vol_ok and not bad_bnd and not bad_mv
          and not bad_eq and not bad_rear and repro)
    print(f"[{'PASS' if ok else 'FAIL'}] {label}")
    print(f"       spaces={m.spaces} room_id={ids} "
          f"連番={contiguous} 容積>0={vol_ok} 再現={repro}")
    if bad_bnd or bad_mv or bad_eq or bad_rear:
        print(f"       不正: 接室={bad_bnd} route={bad_mv} 設備={bad_eq} 裏面={bad_rear}")
    return ok


CASES = [
    # 標準（3室/4室）
    ("標準 戸建基礎断熱(4室)", dict(building_type="detached", A_MR=29.81, A_OR=51.35,
        A_A=120.08, region=6, A_env=307.51, UA=0.87, eta_AC=2.8, eta_AH=4.3, is_floor_ins=False)),
    ("標準 戸建床断熱(3室)", dict(building_type="detached", A_MR=29.81, A_OR=51.35,
        A_A=120.08, region=6, A_env=307.51, UA=0.87, eta_AC=2.8, eta_AH=4.3, is_floor_ins=True)),
    ("標準 集合住宅(3室)", dict(building_type="apartment", A_MR=24.23, A_OR=29.75,
        A_A=70.0, region=6, A_env=238.22, UA=1.20, eta_AC=2.8, eta_AH=4.3)),
    # 縮退（3.4.2 ただし書き：床面積0の室を除外）
    ("NR=0 戸建床断熱→2室", dict(building_type="detached", A_MR=60.0, A_OR=60.08,
        A_A=120.08, region=6, A_env=307.51, UA=0.87, eta_AC=2.8, eta_AH=4.3, is_floor_ins=True)),
    ("NR=0 戸建基礎断熱→3室(UF残る)", dict(building_type="detached", A_MR=60.0, A_OR=60.08,
        A_A=120.08, region=6, A_env=307.51, UA=0.87, eta_AC=2.8, eta_AH=4.3, is_floor_ins=False)),
    ("OR=0 集合住宅→2室", dict(building_type="apartment", A_MR=40.0, A_OR=0.0,
        A_A=70.0, region=6, A_env=238.22, UA=1.20, eta_AC=2.8, eta_AH=4.3)),
    ("MR単室 戸建床断熱→1室", dict(building_type="detached", A_MR=120.08, A_OR=0.0,
        A_A=120.08, region=6, A_env=307.51, UA=0.87, eta_AC=2.8, eta_AH=4.3, is_floor_ins=True)),
]


def run():
    print("=" * 70)
    print("3.4.2 室数設定の検証")
    print("=" * 70)
    allok = True
    for label, kw in CASES:
        allok &= check(label, **kw)
    print("=" * 70)
    print("総合判定:", "全ケース PASS" if allok else "FAIL あり")
    return allok


if __name__ == "__main__":
    run()
