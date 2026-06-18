# -*- coding: utf-8 -*-
"""
reference_data.py
==============================================================================
data/ 配下の CSV を読み込み、暖冷房負荷モデル構築に必要な参照データを
構造化して供給するローダ。計算ロジックはここを通してのみ参照データに触れる。

床断熱/基礎断熱（is_floor_ins）による参照住戸の 0 埋め（表8注記・表12注記）も
本モジュールで適用する。
==============================================================================
"""
import csv
import os

DATA = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data")

SPACES = ["MR", "OR", "NR"]            # 居室（UF以外）
ALL_SPACES = ["MR", "OR", "NR", "UF"]  # 床下空間を含む
VERT_DIRS = ["north", "east", "south", "west"]
ENV_DIRS = ["top"] + VERT_DIRS + ["bottom"]


def _rows(name):
    with open(os.path.join(DATA, name), encoding="utf-8-sig") as f:
        return list(csv.DictReader(f))


def _f(x):
    return float(x) if x not in ("", None) else None


# ---- 読み込み（モジュール初期化時に1度）------------------------------------
_floor = _rows("floor_area.csv")
_env = _rows("envelope_area.csv")
_ufp = _rows("uf_perimeter.csv")
_part = _rows("partition.csv")
_inf = _rows("inner_floor.csv")
_tdc = {r["position"]: float(r["value"]) for r in _rows("temp_diff_coef.csv")}
_dir = _rows("direction_factor.csv")
_days = {int(r["region"]): (int(r["heating_days"]), int(r["cooling_days"]))
         for r in _rows("heating_cooling_days.csv")}
_specu = _rows("spec_u_value.csv")
_layers = _rows("wall_layers.csv")
_hc = {r["part_type"]: float(r["h_c"]) for r in _rows("inside_convective_htc.csv")}


class Reference:
    """1つの建て方・断熱方式に対する参照住戸データを保持する。"""

    def __init__(self, building_type, is_floor_ins):
        self.bt = building_type
        self.is_floor_ins = is_floor_ins  # True=床断熱, False=基礎断熱
        self._load()

    def _load(self):
        bt = self.bt
        # 床面積
        self.floor = {r["space"]: float(r["area"]) for r in _floor if r["building_type"] == bt}
        self.A_A = sum(self.floor.values())
        # 外皮/窓/ドア面積  area[part][direction][space]
        self.area = {"env": {}, "win": {}, "door": {}}
        for r in _env:
            if r["building_type"] != bt:
                continue
            self.area[r["part"]].setdefault(r["direction"], {})[r["space"]] = float(r["area"])
        # 床断熱/基礎断熱による 0 埋め（表8注記）
        if bt == "detached":
            if self.is_floor_ins:
                # 床断熱: 床下空間の n/e/s/w/bottom を 0
                for d in VERT_DIRS + ["bottom"]:
                    self.area["env"][d]["UF"] = 0.0
            else:
                # 基礎断熱: 居室(MR/OR/NR)の bottom を 0
                for s in SPACES:
                    self.area["env"]["bottom"][s] = 0.0
        # 土間床外周長
        self.uf_perimeter = {r["direction"]: float(r["length"]) for r in _ufp
                             if r["building_type"] == bt}
        # 間仕切り（無向）
        self.partition = {}
        for r in _part:
            if r["building_type"] == bt:
                self.partition[(r["space1"], r["space2"])] = float(r["area"])
        # 内壁床（有向 r1->r2）
        self.inner_floor = {}
        for r in _inf:
            if r["building_type"] != bt:
                continue
            v = float(r["area"])
            # 床断熱時は *->UF を 0（表12注記）
            if self.is_floor_ins and r["space2"] == "UF":
                v = 0.0
            self.inner_floor[(r["space1"], r["space2"])] = v

    # --- アクセサ ---
    def env(self, d, s):
        return self.area["env"].get(d, {}).get(s, 0.0)

    def win(self, d, s):
        return self.area["win"].get(d, {}).get(s, 0.0)

    def door(self, d, s):
        return self.area["door"].get(d, {}).get(s, 0.0)

    def vert_total(self, s):
        return sum(self.env(d, s) for d in VERT_DIRS)

    def partition_area(self, r1, r2):
        return self.partition.get((r1, r2), self.partition.get((r2, r1), 0.0))

    def inner_floor_area(self, r1, r2):
        return self.inner_floor.get((r1, r2), 0.0)

    def inner_floor_total(self, r1):
        return sum(v for (a, b), v in self.inner_floor.items() if a == r1)


def get_reference(building_type, is_floor_ins=True):
    return Reference(building_type, is_floor_ins)


# ---- 温度差係数 ------------------------------------------------------------
def temp_diff(position):
    return _tdc[position]


# ---- 方位係数（暖冷房日数加重平均）----------------------------------------
def direction_factor(region):
    """region に対する {direction: nu_weighted} を返す（top/n/e/s/w/bottom）。"""
    nh, nc = _days[region]
    out = {}
    seasons = {"heating": {}, "cooling": {}}
    for r in _dir:
        if int(r["region"]) == region:
            seasons[r["season"]][r["direction"]] = float(r["value"])
    for d in ENV_DIRS:
        vh = seasons["heating"].get(d, 0.0)
        vc = seasons["cooling"].get(d, 0.0)
        out[d] = (vc * nc + vh * nh) / (nc + nh)
    return out


def hc_days(region):
    nh, nc = _days[region]
    return nh, nc


# ---- 仕様基準熱貫流率 ------------------------------------------------------
def spec_u(building_type, part, region):
    for r in _specu:
        if r["building_type"] == building_type and r["part"] == part and int(r["region"]) == region:
            return float(r["u_value"])
    raise KeyError(f"spec_u not found: {building_type}/{part}/region{region}")


# ---- 壁体構成 --------------------------------------------------------------
def layers(building_type, part):
    """指定部位の層リスト（室内側→室外側、順序付き）を返す。"""
    out = []
    for r in _layers:
        if r["building_type"] == building_type and r["part"] == part:
            out.append({
                "order": int(r["order"]), "name": r["name"], "role": r["role"],
                "d": _f(r["d"]), "lambda": _f(r["lambda"]), "c": _f(r["c"]),
                "R": _f(r["R"]), "C": _f(r["C"]),
            })
    return sorted(out, key=lambda x: x["order"])


def r_noins(building_type, part):
    """断熱材を除いた熱抵抗合計 R_noins [m2K/W]。"""
    return sum(l["R"] for l in layers(building_type, part)
               if l["role"] != "insulation" and l["R"] is not None)


def insulation_layer(building_type, part):
    """可変断熱材層（lambda, c）を返す。無ければ None。"""
    for l in layers(building_type, part):
        if l["role"] == "insulation":
            return l
    return None


def hc(part_type):
    return _hc[part_type]
