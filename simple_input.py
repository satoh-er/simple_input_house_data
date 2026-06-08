"""簡易入力（UA・ηAC・ηAH・室用途別床面積・総外皮面積・1階床断熱位置）から
動的熱負荷計算（heat_load_calc）用の入力 JSON を構築するモジュール。

仕様書「簡易な入力方法の検討（改訂版）」3.4.2〜3.4.8 に準拠。
最終成果物は Excel ではなく、heat_load_calc の入力 JSON 仕様
(https://hc-energy.readthedocs.io/ja/latest/contents/02_02_spec_input.html)
に合致した dict（json.dump 可能）。

主なエントリポイント:
    estimate(...) -> dict   入力 JSON 辞書を返す
"""

from __future__ import annotations

import math
from typing import List, Tuple, Optional


# ==============================================================================
# 0. 方位インデックス
#   neu_c / neu_h の方位順: 0:上面 1:北 2:北東 3:東 4:南東 5:南 6:南西 7:西 8:北西 9:下面
# ==============================================================================
IDX_TOP, IDX_N, IDX_E, IDX_S, IDX_W, IDX_BTM = 0, 1, 3, 5, 7, 9


# ==============================================================================
# 1. 温度差係数 (改訂版 表13)
#   方位 vert / top / btm = 1.0 / 1.0 / 0.7
#   外気に接しない部位（界壁・界床・戸境・同用途内壁床）は 0.0 とする。
# ==============================================================================
H_TOP = 1.0
H_VERT = 1.0
H_BTM = 0.7


# ==============================================================================
# 2. 仕様基準熱貫流率 (表16) [W/m2K]、土間床外周部は線熱貫流率 [W/(m・K)]
#   地域→列の対応: {1,2}->0, 3->1, 4->2, {5,6,7}->3, 8->4
#   ※表16で 8 地域が空欄の部位は直近列（5,6,7 地域）の値で補完している。
# ==============================================================================
def _region_col(region: int) -> int:
    return {1: 0, 2: 0, 3: 1, 4: 2, 5: 3, 6: 3, 7: 3, 8: 4}[region]


# (戸建/集合, 部位) -> 5列の値
_SPEC_U_TABLE = {
    ("戸建住宅", "roof"):  [0.17, 0.24, 0.24, 0.24, 0.99],
    ("戸建住宅", "wall"):  [0.35, 0.53, 0.53, 0.53, 0.53],
    ("戸建住宅", "floor"): [0.24, 0.24, 0.34, 0.34, 0.34],
    ("戸建住宅", "base_wall"): [0.27, 0.27, 0.52, 0.52, 0.52],  # 土間床等の外周部分の基礎壁
    ("戸建住宅", "base_hb"):   [1.01, 1.01, 1.05, 1.05, 1.05],  # 土間床等の外周部分（線熱貫流率）
    ("共同住宅", "roof"):  [0.38, 0.55, 0.75, 0.92, 1.18],
    ("共同住宅", "wall"):  [0.47, 0.70, 0.97, 0.97, 0.97],
    ("共同住宅", "floor"): [0.44, 0.61, 0.81, 0.98, 0.98],
    # 共通（窓・ドア）。8 地域は表16で空欄のため 4.7 で補完。
    ("共通", "win"):  [2.3, 2.3, 3.5, 4.7, 4.7],
    ("共通", "door"): [2.3, 2.3, 3.5, 4.7, 4.7],
}


def get_spec_u(tatekata: str, part: str, region: int) -> float:
    key = ("共通", part) if part in ("win", "door") else (tatekata, part)
    return _SPEC_U_TABLE[key][_region_col(region)]


# ==============================================================================
# 3. 壁体構成（表17 戸建 / 表18 集合 / 表19 共通）
#   各部位を次の形式で保持する:
#     fixed:     室内側→室外側の固定層リスト [(name, R[m2K/W], C[kJ/m2K]), ...]
#     ins_index: 断熱層を挿入する位置（fixed リスト中の index）。None なら断熱層なし
#     lambda_ins: 断熱材の熱伝導率 λ [W/mK]
#     c_ins:     断熱材の容積比熱 c [kJ/m3K]
#     r_noins:   断熱なし熱抵抗の合計 [m2K/W]（表面熱伝達抵抗・中空層を含む）
#
#   断熱層が必要な部位では、推定された熱貫流率 U_ex から
#       R_ins = max(1/U_ex - r_noins, 0)
#       C_ins = c_ins * R_ins * lambda_ins   [kJ/m2K]
#   として断熱層を生成し、ins_index の位置に差し込む。
# ==============================================================================
_WALL_SPECS = {
    # ---- 表17 戸建住宅 ----
    ("戸建住宅", "roof"): {
        "fixed": [("plaster_board", 0.047, 8.638)],
        "ins_index": 1, "lambda_ins": 0.050, "c_ins": 8.0, "r_noins": 0.227,
    },
    ("戸建住宅", "wall"): {
        "fixed": [
            ("plaster_board", 0.047, 8.638),
            ("air_gap", 0.071, 0.0),
            ("plywood", 0.075, 8.640),
            ("cement_board", 0.088, 13.235),
        ],
        "ins_index": 2, "lambda_ins": 0.045, "c_ins": 13.0, "r_noins": 0.431,
    },
    ("戸建住宅", "floor"): {
        "fixed": [("plywood", 0.075, 8.640)],
        "ins_index": 1, "lambda_ins": 0.050, "c_ins": 8.0, "r_noins": 0.375,
    },
    ("戸建住宅", "base_wall"): {  # 基礎壁（室内側に断熱材）
        "fixed": [("concrete", 0.075, 240.0)],
        "ins_index": 0, "lambda_ins": 0.022, "c_ins": 77.0, "r_noins": 0.225,
    },
    ("戸建住宅", "inner_floor"): {  # 内壁床（断熱層なし、表17 内壁床）
        "fixed": [
            ("plywood", 0.138, 15.84),
            ("air_gap", 0.07, 0.0),
            ("plaster_board", 0.432, 78.85),
        ],
        "ins_index": None, "lambda_ins": None, "c_ins": None, "r_noins": 0.880,
    },
    # ---- 表18 集合住宅 ----
    ("共同住宅", "roof"): {
        "fixed": [("concrete", 0.094, 300.0)],
        "ins_index": 0, "lambda_ins": 0.023, "c_ins": 60.0, "r_noins": 0.274,
    },
    ("共同住宅", "roof_in"): {  # 外気に接しない屋根（断熱なし）
        "fixed": [("concrete", 0.094, 300.0)],
        "ins_index": None, "lambda_ins": None, "c_ins": None, "r_noins": 0.274,
    },
    ("共同住宅", "wall"): {
        "fixed": [("concrete", 0.084, 270.0)],
        "ins_index": 0, "lambda_ins": 0.034, "c_ins": 61.0, "r_noins": 0.234,
    },
    ("共同住宅", "wall_in"): {  # 外気に接しない外壁（界壁相当・断熱なし）
        "fixed": [("concrete", 0.084, 270.0)],
        "ins_index": None, "lambda_ins": None, "c_ins": None, "r_noins": 0.234,
    },
    ("共同住宅", "floor_in"): {  # 界床（断熱なし）
        "fixed": [("concrete", 0.094, 300.0)],
        "ins_index": None, "lambda_ins": None, "c_ins": None, "r_noins": 0.394,
    },
    ("共同住宅", "inner_floor"): {  # 内壁床＝界床相当（集合は内壁床=界床構成を流用）
        "fixed": [("concrete", 0.094, 300.0)],
        "ins_index": None, "lambda_ins": None, "c_ins": None, "r_noins": 0.394,
    },
    # ---- 表19 共通 ----
    ("共通", "partition"): {  # 間仕切り
        "fixed": [
            ("plaster_board", 0.0555, 9.960),
            ("air_gap", 0.07, 0.0),
            ("plaster_board", 0.0555, 9.960),
        ],
        "ins_index": None, "lambda_ins": None, "c_ins": None, "r_noins": 0.401,
    },
    # ---- 地盤（土間床）を覆う材料（表に明示が無いため一般的なコンクリート土間で代用）----
    ("共通", "ground"): {
        "fixed": [("concrete", 0.075, 240.0)],
        "ins_index": None, "lambda_ins": None, "c_ins": None, "r_noins": 0.075,
    },
}


def _spec_key(tatekata: str, part: str):
    if part in ("partition", "ground"):
        return ("共通", part)
    return (tatekata, part)


def get_r_noins(tatekata: str, part: str) -> float:
    return _WALL_SPECS[_spec_key(tatekata, part)]["r_noins"]


def build_layers(tatekata: str, part: str, u_ex: Optional[float] = None) -> List[dict]:
    """部位の層構成（layer 要素のリスト）を生成する。室内側→室外側の順。

    断熱層を持つ部位では u_ex（推定熱貫流率）から断熱層の熱抵抗・熱容量を求めて挿入する。
    熱抵抗が 0 以下の層は生成しない（heat_load_calc でエラーになるため）。
    """
    spec = _WALL_SPECS[_spec_key(tatekata, part)]
    layers = [
        {"name": n, "thermal_resistance": r, "thermal_capacity": c}
        for (n, r, c) in spec["fixed"] if r > 0.0
    ]
    if spec["ins_index"] is not None and u_ex is not None and u_ex > 0.0:
        r_ins = max(1.0 / u_ex - spec["r_noins"], 0.0)
        if r_ins > 0.0:
            lam = spec["lambda_ins"]
            c_ins = spec["c_ins"] * r_ins * lam  # kJ/m2K
            ins_layer = {
                "name": "insulation",
                "thermal_resistance": r_ins,
                "thermal_capacity": c_ins,
            }
            # fixed の中で r>0 の層だけ残しているので挿入位置を補正する
            insert_at = sum(1 for (n, r, c) in spec["fixed"][: spec["ins_index"]] if r > 0.0)
            layers.insert(insert_at, ins_layer)
    return layers


# ==============================================================================
# 4. 暖冷房期間の日数（表4）
# ==============================================================================
def get_master_days(region: int) -> Tuple[int, int]:
    """(暖房期間日数 n_heating, 冷房期間日数 n_cooling) を返す。"""
    return (
        (257, 53), (252, 48), (244, 53), (242, 53),
        (218, 57), (169, 117), (122, 152), (0, 265),
    )[region - 1]


# ==============================================================================
# 5. 方位係数（表14 暖房期 / 表15 冷房期）
#   返り値はいずれも [上面,北,北東,東,南東,南,南西,西,北西,下面] の 10 要素
# ==============================================================================
def get_neu_avg(region: int) -> Tuple[List[float], List[float]]:
    neu_c = [
        [1.0, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0],
        [0.329, 0.341, 0.335, 0.322, 0.373, 0.341, 0.307, 0.325],
        [0.430, 0.412, 0.390, 0.426, 0.437, 0.431, 0.415, 0.414],
        [0.545, 0.503, 0.468, 0.518, 0.500, 0.512, 0.509, 0.515],
        [0.560, 0.527, 0.487, 0.508, 0.500, 0.498, 0.490, 0.528],
        [0.502, 0.507, 0.476, 0.437, 0.472, 0.434, 0.412, 0.480],
        [0.526, 0.548, 0.550, 0.481, 0.520, 0.491, 0.479, 0.517],
        [0.508, 0.529, 0.553, 0.481, 0.518, 0.504, 0.495, 0.505],
        [0.411, 0.428, 0.447, 0.401, 0.442, 0.427, 0.406, 0.411],
        [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0],
    ]
    neu_h = [
        [1.0, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0, 0.0],
        [0.260, 0.263, 0.284, 0.256, 0.238, 0.261, 0.227, 0.000],
        [0.333, 0.341, 0.348, 0.330, 0.310, 0.325, 0.281, 0.000],
        [0.564, 0.554, 0.540, 0.531, 0.568, 0.579, 0.543, 0.000],
        [0.823, 0.766, 0.751, 0.724, 0.846, 0.833, 0.843, 0.000],
        [0.935, 0.856, 0.851, 0.815, 0.983, 0.936, 1.023, 0.000],
        [0.790, 0.753, 0.750, 0.723, 0.815, 0.763, 0.848, 0.000],
        [0.535, 0.544, 0.542, 0.527, 0.538, 0.523, 0.548, 0.000],
        [0.325, 0.341, 0.351, 0.326, 0.297, 0.317, 0.284, 0.000],
        [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0],
    ]
    col = region - 1
    return [row[col] for row in neu_c], [row[col] for row in neu_h]


# ==============================================================================
# 6. 参照住戸の面積（表5〜表12）
# ==============================================================================
def get_floor_area_ref(tatekata: str) -> Tuple[float, float, float]:
    if tatekata == "共同住宅":
        return 24.23, 29.75, 16.02
    if tatekata == "戸建住宅":
        return 29.81, 51.35, 38.93
    raise ValueError(tatekata)


def get_area_table_ref(tatekata: str) -> Tuple[Tuple[float, ...], ...]:
    """参照住戸の面積一覧。
    行: 外皮(上面/北/東/南/西/下面) / 窓(北/東/南/西) / ドア(北/西)
    列: 主たる居室 / その他の居室 / 非居室 / 床下空間
    """
    if tatekata == "共同住宅":
        return (
            (24.23, 29.75, 16.02, 0.00),   # 外皮-上面
            (0.00, 11.80, 4.16, 0.00),     # 外皮-北
            (0.00, 21.59, 8.05, 0.00),     # 外皮-東
            (9.52, 6.45, 0.00, 0.00),      # 外皮-南
            (17.21, 10.06, 2.37, 0.00),    # 外皮-西
            (24.23, 29.75, 16.02, 0.00),   # 外皮-下面
            (0.00, 2.53, 0.00),            # 窓-北
            (0.00, 0.00, 0.00),            # 窓-東
            (4.52, 3.24, 0.00),            # 窓-南
            (0.00, 0.00, 0.00),            # 窓-西
            (0.00, 0.00, 1.76),            # ドア-北
            (0.00, 0.00, 0.00),            # ドア-西
        )
    if tatekata == "戸建住宅":
        return (
            (0.00, 34.79, 17.40, 0.00),    # 外皮-上面
            (5.12, 6.77, 39.08, 2.81),     # 外皮-北
            (17.20, 8.74, 4.36, 3.28),     # 外皮-東
            (14.21, 29.26, 0.00, 2.91),    # 外皮-南
            (0.00, 17.48, 13.20, 3.28),    # 外皮-西
            (29.81, 16.56, 21.53, 55.48),  # 外皮-下面
            (0.00, 4.59, 3.15),            # 窓-北
            (3.13, 0.66, 0.00),            # 窓-東
            (6.94, 8.17, 0.00),            # 窓-南
            (0.00, 0.99, 1.08),            # 窓-西
            (1.62, 0.00, 1.76),            # ドア-北
            (0.00, 0.00, 1.89),            # ドア-西
        )
    raise ValueError(tatekata)


def get_uf_perimeter_ref(tatekata: str) -> dict:
    """戸建の床下空間の土間床等外周部の長さ（表8）。方位別 [m]。"""
    if tatekata == "戸建住宅":
        return {"north": 10.47, "east": 7.28, "south": 10.47, "west": 7.28}
    return {"north": 0.0, "east": 0.0, "south": 0.0, "west": 0.0}


def get_partition_table_ref(tatekata: str) -> Tuple[float, float, float]:
    """参照住戸の間仕切り面積 (MR-OR, MR-NO, OR-NO)。"""
    if tatekata == "共同住宅":
        return (12.53, 16.19, 40.51)
    if tatekata == "戸建住宅":
        return (8.64, 17.20, 29.51)
    raise ValueError(tatekata)


def get_partition_bottom_table_ref(tatekata: str) -> Tuple[float, ...]:
    """参照住戸の内壁床面積。順序:
    MR->MR, MR->OR, MR->NO, MR->UF, OR->MR, OR->OR, OR->NO, OR->UF,
    NO->MR, NO->OR, NO->NO, NO->UF
    """
    if tatekata == "共同住宅":
        return (0.0,) * 12
    if tatekata == "戸建住宅":
        return (0.0, 0.0, 0.0, 29.81, 21.53, 13.25, 0.0, 16.56, 4.14, 0.0, 12.42, 21.53)
    raise ValueError(tatekata)


# ==============================================================================
# 7. JSON 構築用ヘルパ（heat_load_calc の境界形状に合わせる）
# ==============================================================================
def _h_c(direction: str) -> float:
    if direction in ("s", "sw", "w", "nw", "n", "ne", "e", "se"):
        return 2.5
    if direction == "bottom":
        return 0.7
    if direction == "top":
        return 5.0
    raise ValueError(direction)


def _outside_r(direction: str, temp_dif_coef: float) -> float:
    is_parting = (temp_dif_coef != 1.0)
    if direction in ("s", "sw", "w", "nw", "n", "ne", "e", "se"):
        return 0.04 if not is_parting else 0.11
    if direction == "bottom":
        return 0.15
    if direction == "top":
        return 0.04 if not is_parting else 0.09
    raise ValueError(direction)


def create_equipments(eq_id: int, space_id: int, a_floor: float) -> Tuple[dict, dict]:
    """ルームエアコン（RAC）の能力等を床面積から推定する（既存実装を踏襲）。"""
    q_rtd_c = 190.5 * a_floor + 45.6
    q_rtd_h = 1.2090 * q_rtd_c - 85.1
    q_max_c = max(0.8462 * q_rtd_c + 1205.9, q_rtd_c)
    q_max_h = max(1.7597 * q_max_c - 413.7, q_rtd_h)
    q_min_c = q_min_h = 500
    v_max_c = 11.076 * (q_rtd_c / 1000.0) ** 0.3432
    v_max_h = 11.076 * (q_rtd_h / 1000.0) ** 0.3432
    v_min_c = v_max_c * 0.55
    v_min_h = v_max_h * 0.55
    cooling = {
        "id": eq_id, "name": f"cooling_equipment no.{eq_id}", "equipment_type": "rac",
        "property": {"space_id": space_id, "q_min": q_min_c, "q_max": q_max_c,
                     "v_min": v_min_c, "v_max": v_max_c, "bf": 0.2},
    }
    heating = {
        "id": eq_id, "name": f"heating_equipment no.{eq_id}", "equipment_type": "rac",
        "property": {"space_id": space_id, "q_min": q_min_h, "q_max": q_max_h,
                     "v_min": v_min_h, "v_max": v_max_h, "bf": 0.2},
    }
    return cooling, heating


# ==============================================================================
# 8. メイン: 簡易入力 → 入力 JSON 辞書
# ==============================================================================
def estimate(
    region: int,
    total_floor_area: float,
    main_floor_area: float,
    other_floor_area: float,
    A_env: float,
    ua: float,
    eta_ac: float,
    eta_ah: float,
    tatekata: str,
    structure: str,
    *,
    has_vertical_internal: str = "有",
    ac_method: str = "ot",
    c_value: float = 2.0,
    inside_pressure: str = "negative",
    natural_vent_ach: float = 5.0,
    include_debug: bool = False,
) -> dict:
    """簡易入力から heat_load_calc 用の入力 JSON 辞書を返す。

    Args:
        region: 地域の区分 1-8
        total_floor_area: 床面積の合計 [m2]
        main_floor_area: 主たる居室の床面積 [m2]
        other_floor_area: その他の居室の床面積 [m2]
        A_env: 外皮の部位の面積の合計 [m2]
        ua: 外皮平均熱貫流率 [W/m2K]
        eta_ac: 冷房期平均日射熱取得率（×100 表示値。例 2.8）
        eta_ah: 暖房期平均日射熱取得率（×100 表示値。例 4.3）
        tatekata: "戸建住宅" または "共同住宅"
        structure: "基礎断熱" / "床断熱"（戸建のみ意味を持つ）
        ac_method: 運転モード決定方法（"ot" 等）
        c_value: 相当隙間面積 [cm2/m2]
        inside_pressure: "negative" / "positive" / "balanced"
        natural_vent_ach: 居室の自然風利用換気回数 [1/h]
    """
    if tatekata not in ("戸建住宅", "共同住宅"):
        raise ValueError(tatekata)
    is_kiso = (tatekata == "戸建住宅" and structure == "基礎断熱")

    # ---- 室用途別床面積 ----
    A_MR = main_floor_area
    A_OR = other_floor_area
    A_NO = max(total_floor_area - main_floor_area - other_floor_area, 0.0)  # 3.4.2.1

    # ---- 参照住戸の面積 ----
    A_MR_ref, A_OR_ref, A_NO_ref = get_floor_area_ref(tatekata)
    at = get_area_table_ref(tatekata)
    (A_top_MR_r, A_top_OR_r, A_top_NO_r, A_top_UF_r) = at[0]
    (A_n_MR_r, A_n_OR_r, A_n_NO_r, A_n_UF_r) = at[1]
    (A_e_MR_r, A_e_OR_r, A_e_NO_r, A_e_UF_r) = at[2]
    (A_s_MR_r, A_s_OR_r, A_s_NO_r, A_s_UF_r) = at[3]
    (A_w_MR_r, A_w_OR_r, A_w_NO_r, A_w_UF_r) = at[4]
    (A_btm_MR_r, A_btm_OR_r, A_btm_NO_r, A_btm_UF_r) = at[5]
    (A_win_n_MR_r, A_win_n_OR_r, A_win_n_NO_r) = at[6]
    (A_win_e_MR_r, A_win_e_OR_r, A_win_e_NO_r) = at[7]
    (A_win_s_MR_r, A_win_s_OR_r, A_win_s_NO_r) = at[8]
    (A_win_w_MR_r, A_win_w_OR_r, A_win_w_NO_r) = at[9]
    (A_door_n_MR_r, A_door_n_OR_r, A_door_n_NO_r) = at[10]
    (A_door_w_MR_r, A_door_w_OR_r, A_door_w_NO_r) = at[11]

    # 断熱方法による参照住戸下面・床下面積の読み替え（表8 脚注）
    if is_kiso:
        A_btm_MR_r = A_btm_OR_r = A_btm_NO_r = 0.0
    else:  # 床断熱・集合
        A_btm_UF_r = A_n_UF_r = A_e_UF_r = A_s_UF_r = A_w_UF_r = 0.0

    # 参照住戸 空間ごと垂直外皮面積
    A_vert_MR_r = A_s_MR_r + A_e_MR_r + A_n_MR_r + A_w_MR_r
    A_vert_OR_r = A_s_OR_r + A_e_OR_r + A_n_OR_r + A_w_OR_r
    A_vert_NO_r = A_s_NO_r + A_e_NO_r + A_n_NO_r + A_w_NO_r
    A_vert_UF_r = A_s_UF_r + A_e_UF_r + A_n_UF_r + A_w_UF_r

    # 参照住戸 空間ごと開口部面積
    A_op_MR_r = (A_win_n_MR_r + A_win_e_MR_r + A_win_s_MR_r + A_win_w_MR_r
                 + A_door_n_MR_r + A_door_w_MR_r)
    A_op_OR_r = (A_win_n_OR_r + A_win_e_OR_r + A_win_s_OR_r + A_win_w_OR_r
                 + A_door_n_OR_r + A_door_w_OR_r)
    A_op_NO_r = (A_win_n_NO_r + A_win_e_NO_r + A_win_s_NO_r + A_win_w_NO_r
                 + A_door_n_NO_r + A_door_w_NO_r)

    # ------------------------------------------------------------------
    # 3.4.2.2 水平外皮面積（上面・下面）
    # ------------------------------------------------------------------
    A_top_MR = A_top_MR_r * A_MR / A_MR_ref
    A_top_OR = A_top_OR_r * A_OR / A_OR_ref
    A_top_NO = A_top_NO_r * A_NO / A_NO_ref
    A_top_UF = 0.0

    A_btm_MR = A_btm_MR_r * A_MR / A_MR_ref
    A_btm_OR = A_btm_OR_r * A_OR / A_OR_ref
    A_btm_NO = A_btm_NO_r * A_NO / A_NO_ref
    A_btm_UF = (A_btm_UF_r * (A_MR + A_OR + A_NO) / (A_MR_ref + A_OR_ref + A_NO_ref)
                if A_btm_UF_r > 0.0 else 0.0)

    # ------------------------------------------------------------------
    # 3.4.2.3 垂直外皮面積
    # ------------------------------------------------------------------
    # 床下空間の総垂直外皮面積
    A_vert_UF = (A_vert_UF_r * A_btm_UF / A_btm_UF_r) if A_btm_UF_r > 0.0 else 0.0

    # 居室の総垂直外皮面積（総外皮から水平・床下垂直を減じて床面積案分）
    A_vert = max(A_env - (A_top_MR + A_top_OR + A_top_NO)
                 - (A_btm_MR + A_btm_OR + A_btm_NO) - A_vert_UF, 0.0)
    A_vert_MR = A_vert * A_MR / total_floor_area
    A_vert_OR = A_vert * A_OR / total_floor_area
    A_vert_NO = A_vert * A_NO / total_floor_area

    def _split_vert(A_vert_space, ref_dir, ref_total):
        return A_vert_space * ref_dir / ref_total if ref_total > 0.0 else 0.0

    # 空間ごと方位ごと垂直外皮面積
    A_s_MR = _split_vert(A_vert_MR, A_s_MR_r, A_vert_MR_r)
    A_e_MR = _split_vert(A_vert_MR, A_e_MR_r, A_vert_MR_r)
    A_n_MR = _split_vert(A_vert_MR, A_n_MR_r, A_vert_MR_r)
    A_w_MR = _split_vert(A_vert_MR, A_w_MR_r, A_vert_MR_r)
    A_s_OR = _split_vert(A_vert_OR, A_s_OR_r, A_vert_OR_r)
    A_e_OR = _split_vert(A_vert_OR, A_e_OR_r, A_vert_OR_r)
    A_n_OR = _split_vert(A_vert_OR, A_n_OR_r, A_vert_OR_r)
    A_w_OR = _split_vert(A_vert_OR, A_w_OR_r, A_vert_OR_r)
    A_s_NO = _split_vert(A_vert_NO, A_s_NO_r, A_vert_NO_r)
    A_e_NO = _split_vert(A_vert_NO, A_e_NO_r, A_vert_NO_r)
    A_n_NO = _split_vert(A_vert_NO, A_n_NO_r, A_vert_NO_r)
    A_w_NO = _split_vert(A_vert_NO, A_w_NO_r, A_vert_NO_r)
    A_s_UF = _split_vert(A_vert_UF, A_s_UF_r, A_vert_UF_r)
    A_e_UF = _split_vert(A_vert_UF, A_e_UF_r, A_vert_UF_r)
    A_n_UF = _split_vert(A_vert_UF, A_n_UF_r, A_vert_UF_r)
    A_w_UF = _split_vert(A_vert_UF, A_w_UF_r, A_vert_UF_r)

    sum_top = A_top_MR + A_top_OR + A_top_NO
    sum_btm = A_btm_MR + A_btm_OR + A_btm_NO
    sum_s = A_s_MR + A_s_OR + A_s_NO
    sum_n = A_n_MR + A_n_OR + A_n_NO
    sum_e = A_e_MR + A_e_OR + A_e_NO
    sum_w = A_w_MR + A_w_OR + A_w_NO

    # ------------------------------------------------------------------
    # 3.4.2.4.1 外気に接する総外皮面積
    # ------------------------------------------------------------------
    if tatekata == "戸建住宅":
        r_env_ex = 1.0
    else:
        r_min = (sum_e + sum_w) / A_env
        r_max = (sum_top + sum_s + sum_n + sum_e + sum_w) / A_env
        r_logit = 1.0 / (1.0 + math.exp(-9.10907512 * (ua - 1.05204145)))
        r_env_ex = min(max(r_logit, r_min), r_max)
    A_env_ex = A_env * r_env_ex

    # ------------------------------------------------------------------
    # 3.4.2.4.2 空間ごと方位ごとの外気に接する外皮面積（表3）
    # ------------------------------------------------------------------
    if tatekata == "戸建住宅":
        r_top = r_btm = r_s = r_n = r_e = r_w = 1.0
        r_uf = 1.0  # 床下空間（基礎壁・土間）も外気に接する
    else:
        A_var = sum_top + sum_s + sum_n
        if A_var > 0.0:
            r_var = min(max((A_env_ex - sum_e - sum_w) / A_var, 0.0), 1.0)
        else:
            r_var = 0.0
        r_top = r_var
        r_s = r_var
        r_n = r_var
        r_e = 1.0
        r_w = 1.0
        r_btm = 0.0   # 下面（界床）は外気に接しない
        r_uf = 0.0

    A_top_MR_ex, A_top_OR_ex, A_top_NO_ex = A_top_MR * r_top, A_top_OR * r_top, A_top_NO * r_top
    A_btm_MR_ex, A_btm_OR_ex, A_btm_NO_ex = A_btm_MR * r_btm, A_btm_OR * r_btm, A_btm_NO * r_btm
    A_s_MR_ex, A_s_OR_ex, A_s_NO_ex = A_s_MR * r_s, A_s_OR * r_s, A_s_NO * r_s
    A_n_MR_ex, A_n_OR_ex, A_n_NO_ex = A_n_MR * r_n, A_n_OR * r_n, A_n_NO * r_n
    A_e_MR_ex, A_e_OR_ex, A_e_NO_ex = A_e_MR * r_e, A_e_OR * r_e, A_e_NO * r_e
    A_w_MR_ex, A_w_OR_ex, A_w_NO_ex = A_w_MR * r_w, A_w_OR * r_w, A_w_NO * r_w
    A_s_UF_ex, A_e_UF_ex = A_s_UF * r_uf, A_e_UF * r_uf
    A_n_UF_ex, A_w_UF_ex = A_n_UF * r_uf, A_w_UF * r_uf

    # 床断熱戸建では 1 階床（下面）が外気に接する床
    if tatekata == "戸建住宅" and not is_kiso:
        pass  # r_btm=1.0 のまま（A_btm_*_ex に既に反映）

    # 空間×方位の外気接外皮面積（窓・ドアの上限に使用）
    ex_dir = {
        ("MR", "s"): A_s_MR_ex, ("MR", "e"): A_e_MR_ex, ("MR", "n"): A_n_MR_ex, ("MR", "w"): A_w_MR_ex,
        ("OR", "s"): A_s_OR_ex, ("OR", "e"): A_e_OR_ex, ("OR", "n"): A_n_OR_ex, ("OR", "w"): A_w_OR_ex,
        ("NO", "s"): A_s_NO_ex, ("NO", "e"): A_e_NO_ex, ("NO", "n"): A_n_NO_ex, ("NO", "w"): A_w_NO_ex,
    }

    # ------------------------------------------------------------------
    # 3.4.2.5 開口部面積
    # ------------------------------------------------------------------
    # 総開口部面積
    r_env_op = 1.0 / (1.0 + math.exp(-0.129 * (eta_ac - 13.35)))
    A_op = A_env_ex * r_env_op

    # 空間ごと開口部面積
    denom_op = (A_MR * A_op_MR_r / A_MR_ref + A_OR * A_op_OR_r / A_OR_ref
                + A_NO * A_op_NO_r / A_NO_ref)
    if denom_op > 0.0:
        A_op_MR = A_op * (A_MR * A_op_MR_r / A_MR_ref) / denom_op
        A_op_OR = A_op * (A_OR * A_op_OR_r / A_OR_ref) / denom_op
        A_op_NO = A_op * (A_NO * A_op_NO_r / A_NO_ref) / denom_op
    else:
        A_op_MR = A_op_OR = A_op_NO = 0.0

    def _win(A_op_space, A_win_ref, A_op_ref, ex):
        if A_op_ref <= 0.0:
            return 0.0
        return min(A_op_space * A_win_ref / A_op_ref, ex)

    def _door(A_op_space, A_door_ref, A_op_ref, ex, win):
        if A_op_ref <= 0.0:
            return 0.0
        return min(A_op_space * A_door_ref / A_op_ref, max(ex - win, 0.0))

    # 窓（空間×方位）
    win = {
        ("MR", "s"): _win(A_op_MR, A_win_s_MR_r, A_op_MR_r, ex_dir[("MR", "s")]),
        ("MR", "e"): _win(A_op_MR, A_win_e_MR_r, A_op_MR_r, ex_dir[("MR", "e")]),
        ("MR", "n"): _win(A_op_MR, A_win_n_MR_r, A_op_MR_r, ex_dir[("MR", "n")]),
        ("MR", "w"): _win(A_op_MR, A_win_w_MR_r, A_op_MR_r, ex_dir[("MR", "w")]),
        ("OR", "s"): _win(A_op_OR, A_win_s_OR_r, A_op_OR_r, ex_dir[("OR", "s")]),
        ("OR", "e"): _win(A_op_OR, A_win_e_OR_r, A_op_OR_r, ex_dir[("OR", "e")]),
        ("OR", "n"): _win(A_op_OR, A_win_n_OR_r, A_op_OR_r, ex_dir[("OR", "n")]),
        ("OR", "w"): _win(A_op_OR, A_win_w_OR_r, A_op_OR_r, ex_dir[("OR", "w")]),
        ("NO", "s"): _win(A_op_NO, A_win_s_NO_r, A_op_NO_r, ex_dir[("NO", "s")]),
        ("NO", "e"): _win(A_op_NO, A_win_e_NO_r, A_op_NO_r, ex_dir[("NO", "e")]),
        ("NO", "n"): _win(A_op_NO, A_win_n_NO_r, A_op_NO_r, ex_dir[("NO", "n")]),
        ("NO", "w"): _win(A_op_NO, A_win_w_NO_r, A_op_NO_r, ex_dir[("NO", "w")]),
    }
    # ドア（空間×方位、北・西のみ）
    door = {
        ("MR", "n"): _door(A_op_MR, A_door_n_MR_r, A_op_MR_r, ex_dir[("MR", "n")], win[("MR", "n")]),
        ("MR", "w"): _door(A_op_MR, A_door_w_MR_r, A_op_MR_r, ex_dir[("MR", "w")], win[("MR", "w")]),
        ("OR", "n"): _door(A_op_OR, A_door_n_OR_r, A_op_OR_r, ex_dir[("OR", "n")], win[("OR", "n")]),
        ("OR", "w"): _door(A_op_OR, A_door_w_OR_r, A_op_OR_r, ex_dir[("OR", "w")], win[("OR", "w")]),
        ("NO", "n"): _door(A_op_NO, A_door_n_NO_r, A_op_NO_r, ex_dir[("NO", "n")], win[("NO", "n")]),
        ("NO", "w"): _door(A_op_NO, A_door_w_NO_r, A_op_NO_r, ex_dir[("NO", "w")], win[("NO", "w")]),
    }

    # ------------------------------------------------------------------
    # 3.4.2.6 外気に接する外壁等面積（外皮接面積 − 窓 − ドア）
    # ------------------------------------------------------------------
    def _wall_ex(ex, sp, d):
        return max(ex - win.get((sp, d), 0.0) - door.get((sp, d), 0.0), 0.0)

    wall_ex = {}
    for sp in ("MR", "OR", "NO"):
        for d in ("s", "e", "n", "w"):
            wall_ex[(sp, d)] = _wall_ex(ex_dir[(sp, d)], sp, d)
    # 上面（屋根）はそのまま外気接外皮＝外壁等
    wall_ex[("MR", "top")] = A_top_MR_ex
    wall_ex[("OR", "top")] = A_top_OR_ex
    wall_ex[("NO", "top")] = A_top_NO_ex
    # 下面（床断熱戸建の 1 階床）
    wall_ex[("MR", "btm")] = A_btm_MR_ex
    wall_ex[("OR", "btm")] = A_btm_OR_ex
    wall_ex[("NO", "btm")] = A_btm_NO_ex
    # 床下空間の基礎壁（外気接の垂直）
    wall_ex[("UF", "s")] = A_s_UF_ex
    wall_ex[("UF", "e")] = A_e_UF_ex
    wall_ex[("UF", "n")] = A_n_UF_ex
    wall_ex[("UF", "w")] = A_w_UF_ex

    # ------------------------------------------------------------------
    # 3.4.2.7 外気に接しない外壁等面積（界壁・界床・戸境）
    # ------------------------------------------------------------------
    in_dir = {}
    for sp, A_full in (
        ("MR", {"s": A_s_MR, "e": A_e_MR, "n": A_n_MR, "w": A_w_MR, "top": A_top_MR, "btm": A_btm_MR}),
        ("OR", {"s": A_s_OR, "e": A_e_OR, "n": A_n_OR, "w": A_w_OR, "top": A_top_OR, "btm": A_btm_OR}),
        ("NO", {"s": A_s_NO, "e": A_e_NO, "n": A_n_NO, "w": A_w_NO, "top": A_top_NO, "btm": A_btm_NO}),
    ):
        for d, A_d in A_full.items():
            ex = wall_ex.get((sp, d), 0.0) if d in ("top", "btm") else ex_dir.get((sp, d), 0.0)
            in_dir[(sp, d)] = max(A_d - ex, 0.0)

    # ------------------------------------------------------------------
    # 3.4.2.8 土間床等外周部の長さ（基礎断熱戸建のみ）
    # ------------------------------------------------------------------
    L_uf_ref = get_uf_perimeter_ref(tatekata)
    L_uf_ref_total = sum(L_uf_ref.values())
    A_uf_vert_ref_total = A_vert_UF_r
    if is_kiso and A_uf_vert_ref_total > 0.0:
        # 基礎壁面積比で外周長を案分（方位別には算出しない）
        L_env_uf_ex = L_uf_ref_total * A_vert_UF / A_uf_vert_ref_total
    else:
        L_env_uf_ex = 0.0

    # ------------------------------------------------------------------
    # 3.4.2.9.1 間仕切り面積
    # ------------------------------------------------------------------
    A_part_MR_OR_r, A_part_MR_NO_r, A_part_OR_NO_r = get_partition_table_ref(tatekata)
    if has_vertical_internal == "有":
        def _part(ref, v1, v2, v1r, v2r):
            denom = v1r + v2r
            return ref * (v1 + v2) / denom if denom > 0.0 else 0.0
        A_part_MR_OR = _part(A_part_MR_OR_r, A_vert_MR, A_vert_OR, A_vert_MR_r, A_vert_OR_r)
        A_part_MR_NO = _part(A_part_MR_NO_r, A_vert_MR, A_vert_NO, A_vert_MR_r, A_vert_NO_r)
        A_part_OR_NO = _part(A_part_OR_NO_r, A_vert_OR, A_vert_NO, A_vert_OR_r, A_vert_NO_r)
    elif has_vertical_internal == "無":
        A_part_MR_OR = A_part_MR_NO = A_part_OR_NO = 0.0
    else:
        raise ValueError(has_vertical_internal)

    # ------------------------------------------------------------------
    # 3.4.2.9.2 内壁床面積
    # ------------------------------------------------------------------
    pbt = get_partition_bottom_table_ref(tatekata)
    (pb_MR_MR_r, pb_MR_OR_r, pb_MR_NO_r, pb_MR_UF_r,
     pb_OR_MR_r, pb_OR_OR_r, pb_OR_NO_r, pb_OR_UF_r,
     pb_NO_MR_r, pb_NO_OR_r, pb_NO_NO_r, pb_NO_UF_r) = pbt
    if not is_kiso:
        pb_MR_UF_r = pb_OR_UF_r = pb_NO_UF_r = 0.0

    A_pb_MR = max(A_MR - A_btm_MR, 0.0)
    A_pb_OR = max(A_OR - A_btm_OR, 0.0)
    A_pb_NO = max(A_NO - A_btm_NO, 0.0)
    A_pb_MR_ref = pb_MR_MR_r + pb_MR_OR_r + pb_MR_NO_r + pb_MR_UF_r
    A_pb_OR_ref = pb_OR_MR_r + pb_OR_OR_r + pb_OR_NO_r + pb_OR_UF_r
    A_pb_NO_ref = pb_NO_MR_r + pb_NO_OR_r + pb_NO_NO_r + pb_NO_UF_r

    def _pb(A_pb, ref, total):
        return A_pb * ref / total if total > 0.0 else 0.0

    pb = {
        ("MR", "MR"): _pb(A_pb_MR, pb_MR_MR_r, A_pb_MR_ref),
        ("MR", "OR"): _pb(A_pb_MR, pb_MR_OR_r, A_pb_MR_ref),
        ("MR", "NO"): _pb(A_pb_MR, pb_MR_NO_r, A_pb_MR_ref),
        ("MR", "UF"): _pb(A_pb_MR, pb_MR_UF_r, A_pb_MR_ref),
        ("OR", "MR"): _pb(A_pb_OR, pb_OR_MR_r, A_pb_OR_ref),
        ("OR", "OR"): _pb(A_pb_OR, pb_OR_OR_r, A_pb_OR_ref),
        ("OR", "NO"): _pb(A_pb_OR, pb_OR_NO_r, A_pb_OR_ref),
        ("OR", "UF"): _pb(A_pb_OR, pb_OR_UF_r, A_pb_OR_ref),
        ("NO", "MR"): _pb(A_pb_NO, pb_NO_MR_r, A_pb_NO_ref),
        ("NO", "OR"): _pb(A_pb_NO, pb_NO_OR_r, A_pb_NO_ref),
        ("NO", "NO"): _pb(A_pb_NO, pb_NO_NO_r, A_pb_NO_ref),
        ("NO", "UF"): _pb(A_pb_NO, pb_NO_UF_r, A_pb_NO_ref),
    }

    # ------------------------------------------------------------------
    # 3.4.3 室容積
    # ------------------------------------------------------------------
    V_MR = 2.4 * A_MR
    V_OR = 2.4 * A_OR
    V_NO = 2.4 * A_NO
    V_UF = 0.4 * A_btm_UF

    # ------------------------------------------------------------------
    # 3.4.4 熱貫流率の推定（仕様基準 U × 比例配分）
    # ------------------------------------------------------------------
    # 外気接部位の集計面積
    A_roof_ex = wall_ex[("MR", "top")] + wall_ex[("OR", "top")] + wall_ex[("NO", "top")]
    A_wall_vert_ex = sum(wall_ex[(sp, d)] for sp in ("MR", "OR", "NO") for d in ("s", "e", "n", "w"))
    A_floor_ex = wall_ex[("MR", "btm")] + wall_ex[("OR", "btm")] + wall_ex[("NO", "btm")]
    A_base_ex = wall_ex[("UF", "s")] + wall_ex[("UF", "e")] + wall_ex[("UF", "n")] + wall_ex[("UF", "w")]
    A_win_all = sum(win.values())
    A_door_all = sum(door.values())

    U_roof_spec = get_spec_u(tatekata, "roof", region)
    U_wall_spec = get_spec_u(tatekata, "wall", region)
    U_floor_spec = get_spec_u(tatekata, "floor", region)
    U_win_spec = get_spec_u(tatekata, "win", region)
    U_door_spec = get_spec_u(tatekata, "door", region)

    # 3.4.4.1 床の外気接面積:
    #   床断熱 -> 居室下面（1階床）。基礎断熱 -> 床下空間に接する内壁床（1階床）の合計。
    if is_kiso:
        A_floor_ex_q = pb[("MR", "UF")] + pb[("OR", "UF")] + pb[("NO", "UF")]
    else:
        A_floor_ex_q = A_floor_ex

    q_spec_roof = U_roof_spec * A_roof_ex * H_TOP
    q_spec_wall = U_wall_spec * A_wall_vert_ex * H_VERT
    q_spec_floor = U_floor_spec * A_floor_ex_q * H_BTM
    q_spec_win = U_win_spec * A_win_all * H_VERT
    q_spec_door = U_door_spec * A_door_all * H_VERT
    q_spec_all = q_spec_roof + q_spec_wall + q_spec_floor + q_spec_win + q_spec_door

    q_target_all = ua * A_env
    if q_spec_all > 0.0:
        q_t_roof = q_target_all * q_spec_roof / q_spec_all
        q_t_wall = q_target_all * q_spec_wall / q_spec_all
        q_t_floor = q_target_all * q_spec_floor / q_spec_all
        q_t_win = q_target_all * q_spec_win / q_spec_all
        q_t_door = q_target_all * q_spec_door / q_spec_all
    else:
        q_t_roof = q_t_wall = q_t_floor = q_t_win = q_t_door = 0.0

    U_roof_ex = q_t_roof / (H_TOP * A_roof_ex) if A_roof_ex > 0.0 else 0.0
    U_wall_ex = q_t_wall / (H_VERT * A_wall_vert_ex) if A_wall_vert_ex > 0.0 else 0.0
    U_win = q_t_win / (H_VERT * A_win_all) if A_win_all > 0.0 else 0.0
    U_door = q_t_door / (H_VERT * A_door_all) if A_door_all > 0.0 else 0.0

    # 床／基礎の熱貫流率
    if is_kiso:
        # 3.4.4.2.2 基礎断熱: 床の目標熱損失を基礎壁と土間床外周部に配分
        U_base_spec = get_spec_u(tatekata, "base_wall", region)
        psi_base_spec = get_spec_u(tatekata, "base_hb", region)
        q_spec_uf_wall = U_base_spec * A_base_ex * H_VERT
        q_spec_uf_hb = psi_base_spec * L_env_uf_ex * H_VERT
        denom_uf = q_spec_uf_wall + q_spec_uf_hb
        q_t_uf_wall = q_t_floor * q_spec_uf_wall / denom_uf if denom_uf > 0.0 else 0.0
        q_t_uf_hb = q_t_floor * q_spec_uf_hb / denom_uf if denom_uf > 0.0 else 0.0
        U_base = q_t_uf_wall / (H_VERT * A_base_ex) if A_base_ex > 0.0 else 0.0
        psi_base = q_t_uf_hb / (H_VERT * L_env_uf_ex) if L_env_uf_ex > 0.0 else 0.0
        U_floor_ex = 0.0
    else:
        # 3.4.4.2.1 床断熱: 床の熱貫流率（温度差係数 0.7）
        U_floor_ex = q_t_floor / (H_BTM * A_floor_ex) if A_floor_ex > 0.0 else 0.0
        U_base = 0.0
        psi_base = 0.0

    # ------------------------------------------------------------------
    # 3.4.6 平均日射熱取得率目標値
    # ------------------------------------------------------------------
    n_h, n_c = get_master_days(region)
    eta_avg = (eta_ac * n_c + eta_ah * n_h) / (n_c + n_h)  # ×100 表示値
    eta_A_target = eta_avg / 100.0
    m_total = eta_A_target * A_env

    # ------------------------------------------------------------------
    # 3.4.7 窓の日射熱取得率の推定
    # ------------------------------------------------------------------
    neu_c, neu_h = get_neu_avg(region)

    def _nu(idx):
        return (neu_c[idx] * n_c + neu_h[idx] * n_h) / (n_c + n_h)

    nu_top, nu_n, nu_e, nu_s, nu_w = _nu(IDX_TOP), _nu(IDX_N), _nu(IDX_E), _nu(IDX_S), _nu(IDX_W)

    # 不透明部位（屋根・外壁・ドア）の日射熱取得量 m_model_wall
    sum_wall_s = sum(wall_ex[(sp, "s")] for sp in ("MR", "OR", "NO"))
    sum_wall_e = sum(wall_ex[(sp, "e")] for sp in ("MR", "OR", "NO"))
    sum_wall_n = sum(wall_ex[(sp, "n")] for sp in ("MR", "OR", "NO"))
    sum_wall_w = sum(wall_ex[(sp, "w")] for sp in ("MR", "OR", "NO"))
    sum_door_n = sum(door.get((sp, "n"), 0.0) for sp in ("MR", "OR", "NO"))
    sum_door_w = sum(door.get((sp, "w"), 0.0) for sp in ("MR", "OR", "NO"))

    m_model_wall = 0.034 * (
        U_roof_ex * A_roof_ex * nu_top
        + U_wall_ex * (sum_wall_s * nu_s + sum_wall_e * nu_e
                       + sum_wall_n * nu_n + sum_wall_w * nu_w)
        + U_door * (sum_door_n * nu_n + sum_door_w * nu_w)
    )

    sum_win_s = sum(win[(sp, "s")] for sp in ("MR", "OR", "NO"))
    sum_win_e = sum(win[(sp, "e")] for sp in ("MR", "OR", "NO"))
    sum_win_n = sum(win[(sp, "n")] for sp in ("MR", "OR", "NO"))
    sum_win_w = sum(win[(sp, "w")] for sp in ("MR", "OR", "NO"))
    win_nu_sum = (sum_win_s * nu_s + sum_win_e * nu_e + sum_win_n * nu_n + sum_win_w * nu_w)

    if win_nu_sum > 1e-9:
        eta_win_temp = (m_total - m_model_wall) / win_nu_sum
    else:
        # 窓を配置できる外気接方位が無い（縮退ケース）。窓の日射熱取得は評価不能。
        eta_win_temp = 0.0

    # ------------------------------------------------------------------
    # 3.4.8 窓の日射熱取得率・窓面積・窓の熱貫流率の補正
    # ------------------------------------------------------------------
    ETA_WIN_MIN, ETA_WIN_MAX = 0.10, 0.73
    if eta_win_temp < ETA_WIN_MIN:
        eta_win = ETA_WIN_MIN
        win_scale = 1.0
    elif eta_win_temp > ETA_WIN_MAX:
        eta_win = ETA_WIN_MAX
        win_scale = eta_win_temp / ETA_WIN_MAX
    else:
        eta_win = max(eta_win_temp, 1e-8)
        win_scale = 1.0

    # 窓面積・熱貫流率の補正（面積を win_scale 倍、熱損失保存のため U を A_win/A_mod 倍）。
    # 面積拡大時に外気接面積を超えないようにクランプし、外壁等面積を再計算する。
    # クランプにより総窓面積が名目どおり拡大できない場合でも、窓の熱損失
    #   U_mod * ΣA_mod = U_win,ex * ΣA_win
    # が保たれるよう、U は総面積比で補正する（仕様 3.4.8 の趣旨）。
    if win_scale != 1.0:
        A_win_before = sum(win.values())
        for sp in ("MR", "OR", "NO"):
            for d in ("s", "e", "n", "w"):
                w0 = win[(sp, d)]
                if w0 <= 0.0:
                    continue
                w_new = min(w0 * win_scale, max(ex_dir[(sp, d)] - door.get((sp, d), 0.0), 0.0))
                win[(sp, d)] = w_new
                wall_ex[(sp, d)] = max(ex_dir[(sp, d)] - w_new - door.get((sp, d), 0.0), 0.0)
        A_win_after = sum(win.values())
        if A_win_after > 0.0:
            U_win = U_win * A_win_before / A_win_after
        # 集計面積を更新
        A_win_all = A_win_after
        A_wall_vert_ex = sum(wall_ex[(s, dd)] for s in ("MR", "OR", "NO") for dd in ("s", "e", "n", "w"))

    # ==================================================================
    # 9. 入力 JSON 辞書の組み立て
    # ==================================================================
    DIR_KEY = {"s": "s", "e": "e", "n": "n", "w": "w", "top": "top", "btm": "bottom"}
    ROOM_ID = {"MR": 0, "OR": 1, "NO": 2, "UF": 3}
    has_uf = is_kiso and A_btm_UF > 0.0

    boundaries: List[dict] = []
    _counter = {"i": 0}

    def _next_id() -> int:
        i = _counter["i"]
        _counter["i"] += 1
        return i

    def add_general(name, room, direction, area, part, temp_dif_coef, is_floor):
        """external_general_part を追加。"""
        if area <= 1e-6:
            return
        d = DIR_KEY[direction]
        u_ex = {
            "roof": U_roof_ex, "roof_in": None, "wall": U_wall_ex, "wall_in": None,
            "floor": U_floor_ex, "floor_in": None, "base_wall": U_base,
        }.get(part)
        sun = (temp_dif_coef == 1.0)
        boundaries.append({
            "id": _next_id(),
            "name": name, "sub_name": "",
            "connected_room_id": room,
            "boundary_type": "external_general_part",
            "area": area,
            "h_c": _h_c(d),
            "is_solar_absorbed_inside": bool(is_floor),
            "is_floor": bool(is_floor),
            "layers": build_layers(tatekata, part, u_ex),
            "solar_shading_part": {"existence": False},
            "is_sun_striked_outside": sun,
            "direction": d,
            "outside_emissivity": 0.9,
            "outside_heat_transfer_resistance": _outside_r(d, temp_dif_coef),
            "outside_solar_absorption": 0.8,
            "temp_dif_coef": temp_dif_coef,
        })

    def add_opaque(name, room, direction, area, u_value):
        if area <= 1e-6:
            return
        d = DIR_KEY[direction]
        boundaries.append({
            "id": _next_id(),
            "name": name, "sub_name": "",
            "connected_room_id": room,
            "boundary_type": "external_opaque_part",
            "area": area,
            "h_c": _h_c(d),
            "is_solar_absorbed_inside": False,
            "is_floor": False,
            "solar_shading_part": {"existence": False},
            "is_sun_striked_outside": True,
            "direction": d,
            "outside_emissivity": 0.9,
            "outside_heat_transfer_resistance": _outside_r(d, 1.0),
            "u_value": u_value,
            "inside_heat_transfer_resistance": 0.11,
            "outside_solar_absorption": 0.8,
            "temp_dif_coef": 1.0,
        })

    def add_transparent(name, room, direction, area, u_value, eta_value):
        if area <= 1e-6:
            return
        d = DIR_KEY[direction]
        boundaries.append({
            "id": _next_id(),
            "name": name, "sub_name": "",
            "connected_room_id": room,
            "boundary_type": "external_transparent_part",
            "area": area,
            "h_c": _h_c(d),
            "is_solar_absorbed_inside": False,
            "is_floor": False,
            "solar_shading_part": {"existence": False},
            "is_sun_striked_outside": True,
            "direction": d,
            "outside_emissivity": 0.9,
            "outside_heat_transfer_resistance": _outside_r(d, 1.0),
            "u_value": u_value,
            "inside_heat_transfer_resistance": 0.11,
            "eta_value": eta_value,
            "incident_angle_characteristics": "multiple",
            "glass_area_ratio": 0.72,
            "temp_dif_coef": 1.0,
        })

    def add_internal_pair(name, room_a, room_b, area, part, orientation):
        """間仕切り・内壁床を室Aと室Bの両面で追加（rear_surface 相互参照）。

        orientation: "vertical"（間仕切り壁）/ "floor"（room_a が上階＝床、room_b が下階＝天井）
        """
        if area <= 1e-6:
            return
        id_a = _next_id()
        id_b = _next_id()
        layers_a = build_layers(tatekata, part)
        layers_b = list(reversed(layers_a))
        if orientation == "vertical":
            hc_a = hc_b = 2.5
            isf_a = isf_b = False
        else:  # floor
            hc_a, isf_a = 0.7, True    # 上階の床
            hc_b, isf_b = 5.0, False   # 下階の天井
        boundaries.append({
            "id": id_a, "name": f"{name}_a", "sub_name": "",
            "connected_room_id": room_a, "boundary_type": "internal",
            "area": area, "h_c": hc_a,
            "is_solar_absorbed_inside": isf_a, "is_floor": isf_a,
            "layers": layers_a, "solar_shading_part": {"existence": False},
            "rear_surface_boundary_id": id_b,
        })
        boundaries.append({
            "id": id_b, "name": f"{name}_b", "sub_name": "",
            "connected_room_id": room_b, "boundary_type": "internal",
            "area": area, "h_c": hc_b,
            "is_solar_absorbed_inside": isf_b, "is_floor": isf_b,
            "layers": layers_b, "solar_shading_part": {"existence": False},
            "rear_surface_boundary_id": id_a,
        })

    def add_ground(name, room, area):
        if area <= 1e-6:
            return
        boundaries.append({
            "id": _next_id(), "name": name, "sub_name": "",
            "connected_room_id": room, "boundary_type": "ground",
            "area": area,
            "h_c": _h_c("bottom"),
            "is_solar_absorbed_inside": True, "is_floor": True,
            "layers": build_layers(tatekata, "ground"),
            "solar_shading_part": {"existence": False},
        })

    # 外気に接しない外壁の部位名（集合は界壁/界床/外気に接しない屋根、戸建は基本発生しない）
    wall_in_part = "wall_in" if tatekata == "共同住宅" else "wall"
    roof_in_part = "roof_in" if tatekata == "共同住宅" else "roof"
    floor_in_part = "floor_in" if tatekata == "共同住宅" else "floor"

    for sp in ("MR", "OR", "NO"):
        rid = ROOM_ID[sp]
        # --- 外気に接する外壁等 ---
        add_general(f"{sp}_roof_ex", rid, "top", wall_ex[(sp, "top")], "roof", H_TOP, is_floor=False)
        for d in ("s", "e", "n", "w"):
            add_general(f"{sp}_wall_{d}_ex", rid, d, wall_ex[(sp, d)], "wall", H_VERT, is_floor=False)
        if not is_kiso:
            # 床断熱戸建・集合（集合は r_btm=0 のため面積0）：1階床（外気に接する床）
            add_general(f"{sp}_floor_ex", rid, "btm", wall_ex[(sp, "btm")], "floor", H_BTM, is_floor=True)
        # --- 外気に接しない外壁等（界壁・界床・戸境）temp_dif_coef=0 ---
        add_general(f"{sp}_roof_in", rid, "top", in_dir[(sp, "top")], roof_in_part, 0.0, is_floor=False)
        for d in ("s", "e", "n", "w"):
            add_general(f"{sp}_wall_{d}_in", rid, d, in_dir[(sp, d)], wall_in_part, 0.0, is_floor=False)
        add_general(f"{sp}_floor_in", rid, "btm", in_dir[(sp, "btm")], floor_in_part, 0.0, is_floor=True)
        # --- 窓・ドア ---
        for d in ("s", "e", "n", "w"):
            add_transparent(f"{sp}_win_{d}", rid, d, win[(sp, d)], U_win, eta_win)
        for d in ("n", "w"):
            add_opaque(f"{sp}_door_{d}", rid, d, door.get((sp, d), 0.0), U_door)

    # --- 床下空間（基礎断熱戸建）---
    if has_uf:
        rid_uf = ROOM_ID["UF"]
        for d in ("s", "e", "n", "w"):
            add_general(f"UF_base_{d}", rid_uf, d, wall_ex[("UF", d)], "base_wall", H_VERT, is_floor=False)
        add_ground("UF_ground", rid_uf, A_btm_UF)

    # --- 間仕切り壁（内部・垂直）---
    add_internal_pair("part_MR_OR", ROOM_ID["MR"], ROOM_ID["OR"], A_part_MR_OR, "partition", "vertical")
    add_internal_pair("part_MR_NO", ROOM_ID["MR"], ROOM_ID["NO"], A_part_MR_NO, "partition", "vertical")
    add_internal_pair("part_OR_NO", ROOM_ID["OR"], ROOM_ID["NO"], A_part_OR_NO, "partition", "vertical")

    # --- 内壁床 ---
    #  同じ室用途同士（MR-MR, OR-OR, NO-NO）は「温度差係数0の外気に接する床」として扱う
    for sp in ("MR", "OR", "NO"):
        add_general(f"innerfloor_{sp}_{sp}", ROOM_ID[sp], "btm", pb[(sp, sp)],
                    "inner_floor", 0.0, is_floor=True)
    #  異なる室用途間（上階→下階）は内部境界（両面）として扱う
    cross_pairs = [
        ("MR", "OR"), ("MR", "NO"),
        ("OR", "MR"), ("OR", "NO"),
        ("NO", "MR"), ("NO", "OR"),
    ]
    for a, b in cross_pairs:
        add_internal_pair(f"innerfloor_{a}_{b}", ROOM_ID[a], ROOM_ID[b], pb[(a, b)],
                          "inner_floor", "floor")
    #  床下空間に接する内壁床（基礎断熱戸建）
    if has_uf:
        for sp in ("MR", "OR", "NO"):
            add_internal_pair(f"innerfloor_{sp}_UF", ROOM_ID[sp], ROOM_ID["UF"], pb[(sp, "UF")],
                              "inner_floor", "floor")

    # ------------------------------------------------------------------
    # rooms
    # ------------------------------------------------------------------
    rooms = [
        {
            "id": 0, "name": "main_occupant_room", "sub_name": "",
            "floor_area": A_MR, "volume": V_MR,
            "ventilation": {"natural": natural_vent_ach * V_MR},
            "furniture": {"input_method": "default"},
            "schedule": {"name": "main_occupant_room"},
        },
        {
            "id": 1, "name": "other_occupant_room", "sub_name": "",
            "floor_area": A_OR, "volume": V_OR,
            "ventilation": {"natural": natural_vent_ach * V_OR},
            "furniture": {"input_method": "default"},
            "schedule": {"name": "other_occupant_room"},
        },
        {
            "id": 2, "name": "non_occupant_room", "sub_name": "",
            "floor_area": A_NO, "volume": V_NO,
            "ventilation": {"natural": natural_vent_ach * V_NO},
            "furniture": {"input_method": "default"},
            "schedule": {"name": "non_occupant_room"},
        },
    ]
    if has_uf:
        rooms.append({
            "id": 3, "name": "underfloor", "sub_name": "",
            "floor_area": A_btm_UF, "volume": V_UF,
            "ventilation": {"natural": 0.0},
            "furniture": {"input_method": "default"},
            "schedule": {"name": "non_occupant_room"},
        })

    # ------------------------------------------------------------------
    # mechanical_ventilations（第3種、0.5回/h を NO へ分配）
    # ------------------------------------------------------------------
    vent_rate = 0.5
    v_vent_MR = vent_rate * (V_MR + V_NO * V_MR / (V_MR + V_OR)) if (V_MR + V_OR) > 0 else 0.0
    v_vent_OR = vent_rate * (V_OR + V_NO * V_OR / (V_MR + V_OR)) if (V_MR + V_OR) > 0 else 0.0
    mechanical_ventilations = [
        # NOTE: 入力 JSON 仕様(readthedocs)では経路キーは "route" だが、
        #       既存の検証済みパイプラインでは "root" が使われていた。
        #       エンジン側の受け口に合わせて _ROUTE_KEY を切り替えること。
        {"id": 0, "root_type": "type3", "volume": v_vent_MR, "root": [0, 2]},
        {"id": 1, "root_type": "type3", "volume": v_vent_OR, "root": [1, 2]},
    ]

    # ------------------------------------------------------------------
    # equipments（MR, OR に RAC）
    # ------------------------------------------------------------------
    eq_c_MR, eq_h_MR = create_equipments(0, 0, A_MR)
    eq_c_OR, eq_h_OR = create_equipments(1, 1, A_OR)
    equipments = {
        "heating_equipments": [eq_h_MR, eq_h_OR],
        "cooling_equipments": [eq_c_MR, eq_c_OR],
    }

    common = {
        "ac_method": ac_method,
        "weather": {"method": "ees", "region": region},
    }
    building = {
        "infiltration": {
            "method": "balance_residential",
            "story": 2 if tatekata == "戸建住宅" else 1,
            "c_value_estimate": "specify",
            "c_value": c_value,
            "inside_pressure": inside_pressure,
        }
    }

    result = {
        "common": common,
        "building": building,
        "rooms": rooms,
        "boundaries": boundaries,
        "mechanical_ventilations": mechanical_ventilations,
        "equipments": equipments,
    }

    if include_debug:
        # 参考情報（計算には未使用・検証用）。仕様準拠の出力には含めない。
        result["_debug"] = {
            "A_NO": A_NO, "A_env_ex": A_env_ex, "r_env_ex": r_env_ex,
            "A_op": A_op, "r_env_op": r_env_op,
            "U_roof_ex": U_roof_ex, "U_wall_ex": U_wall_ex,
            "U_floor_ex": U_floor_ex, "U_base": U_base, "psi_base": psi_base,
            "U_win": U_win, "U_door": U_door,
            "eta_win_temp": eta_win_temp, "eta_win": eta_win,
            "L_env_uf_ex": L_env_uf_ex,
            "sum_area_check": {
                "A_env_input": A_env,
                "A_top": sum_top, "A_btm": sum_btm + A_btm_UF,
                "A_vert": A_vert + A_vert_UF,
            },
        }
    return result


if __name__ == "__main__":
    import json

    result = estimate(
        region=3,
        total_floor_area=83.38,
        main_floor_area=29.225,
        other_floor_area=34.47,
        A_env=264.12,
        ua=0.87,
        eta_ac=2.8,
        eta_ah=4.3,
        tatekata="戸建住宅",
        structure="基礎断熱",
    )
    print(json.dumps(result, ensure_ascii=False, indent=2))
