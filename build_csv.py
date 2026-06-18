# -*- coding: utf-8 -*-
"""
build_csv.py
==============================================================================
仕様書「簡易入力からの暖冷房負荷モデルの構築」の表5〜表19 等に記載された
参照住戸データ・仕様基準値を data/ 配下の CSV に書き出す。

このファイルは「参照データの唯一の出所」。数値を変えたい場合は CSV を直接
編集すればよく、計算ロジック（reference_data.py / simple_input.py 等）は
一切ハードコードを持たない。再シードしたいときのみ本スクリプトを実行する。

各 CSV の出所は仕様書の表番号をコメントで明記する。
==============================================================================
"""
import csv
import os

DATA = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data")
os.makedirs(DATA, exist_ok=True)


def write(name, header, rows):
    with open(os.path.join(DATA, name), "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(header)
        w.writerows(rows)


# --- 表5・表6 参照住戸の床面積 [m2] -----------------------------------------
# building_type: detached(戸建) / apartment(集合)
write("floor_area.csv", ["building_type", "space", "area"], [
    ["apartment", "MR", 24.23], ["apartment", "OR", 29.75], ["apartment", "NR", 16.02],
    ["detached",  "MR", 29.81], ["detached",  "OR", 51.35], ["detached",  "NR", 38.93],
])

# --- 表7・表8 参照住戸の外皮/窓/ドア面積 [m2] -------------------------------
# part: env(外皮) / win(窓) / door(ドア)
# env の direction: top,north,east,south,west,bottom
# win/door の direction: north,east,south,west
# 注: 表8の床下空間 bottom=55.48 は「基礎断熱時」の値。床断熱/基礎断熱の
#     0埋めは reference_data.py 側で is_floor_ins により行う（表8注記）。
env_rows = []


def er(bt, part, d, mr, orr, nr, uf=None):
    env_rows.append([bt, part, d, "MR", mr])
    env_rows.append([bt, part, d, "OR", orr])
    env_rows.append([bt, part, d, "NR", nr])
    if uf is not None:
        env_rows.append([bt, part, d, "UF", uf])


# 集合住宅（表7）
er("apartment", "env", "top",    24.23, 29.75, 16.02, 0.00)
er("apartment", "env", "north",   0.00, 11.80,  4.16, 0.00)
er("apartment", "env", "east",    0.00, 21.59,  8.05, 0.00)
er("apartment", "env", "south",   9.52,  6.45,  0.00, 0.00)
er("apartment", "env", "west",   17.21, 10.06,  2.37, 0.00)
er("apartment", "env", "bottom", 24.23, 29.75, 16.02, 0.00)
er("apartment", "win", "north",   0.00,  2.53,  0.00)
er("apartment", "win", "east",    0.00,  0.00,  0.00)
er("apartment", "win", "south",   4.52,  3.24,  0.00)
er("apartment", "win", "west",    0.00,  0.00,  0.00)
er("apartment", "door", "north",  0.00,  0.00,  1.76)
er("apartment", "door", "east",   0.00,  0.00,  0.00)
er("apartment", "door", "south",  0.00,  0.00,  0.00)
er("apartment", "door", "west",   0.00,  0.00,  0.00)

# 戸建住宅（表8）
er("detached", "env", "top",     0.00, 34.79, 17.40,  0.00)
er("detached", "env", "north",   5.12,  6.77, 39.08,  2.81)
er("detached", "env", "east",   17.20,  8.74,  4.36,  3.28)
er("detached", "env", "south",  14.21, 29.26,  0.00,  2.91)
er("detached", "env", "west",    0.00, 17.48, 13.20,  3.28)
er("detached", "env", "bottom", 29.81, 16.56, 21.53, 55.48)
er("detached", "win", "north",   0.00,  4.59,  3.15)
er("detached", "win", "east",    3.13,  0.66,  0.00)
er("detached", "win", "south",   6.94,  8.17,  0.00)
er("detached", "win", "west",    0.00,  0.99,  1.08)
er("detached", "door", "north",  1.62,  0.00,  1.76)
er("detached", "door", "east",   0.00,  0.00,  0.00)
er("detached", "door", "south",  0.00,  0.00,  0.00)
er("detached", "door", "west",   0.00,  0.00,  1.89)
write("envelope_area.csv", ["building_type", "part", "direction", "space", "area"], env_rows)

# --- 表8 土間床等外周部の長さ [m]（戸建のみ） -------------------------------
write("uf_perimeter.csv", ["building_type", "direction", "length"], [
    ["detached", "north", 10.47], ["detached", "east", 7.28],
    ["detached", "south", 10.47], ["detached", "west", 7.28],
])

# --- 表9・表10 参照住戸の間仕切り面積 [m2] ----------------------------------
write("partition.csv", ["building_type", "space1", "space2", "area"], [
    ["apartment", "MR", "OR", 12.53], ["apartment", "MR", "NR", 16.19], ["apartment", "OR", "NR", 40.51],
    ["detached",  "MR", "OR",  8.64], ["detached",  "MR", "NR", 17.20], ["detached",  "OR", "NR", 29.51],
])

# --- 表11・表12 参照住戸の内壁床面積 [m2] -----------------------------------
# 注: 床断熱時の *→UF は reference_data.py 側で 0 埋め（表12注記）
write("inner_floor.csv", ["building_type", "space1", "space2", "area"], [
    ["apartment", "MR", "MR", 0.00], ["apartment", "MR", "OR", 0.00], ["apartment", "MR", "NR", 0.00], ["apartment", "MR", "UF", 0.00],
    ["apartment", "OR", "MR", 0.00], ["apartment", "OR", "OR", 0.00], ["apartment", "OR", "NR", 0.00], ["apartment", "OR", "UF", 0.00],
    ["apartment", "NR", "MR", 0.00], ["apartment", "NR", "OR", 0.00], ["apartment", "NR", "NR", 0.00], ["apartment", "NR", "UF", 0.00],
    ["detached", "MR", "MR", 0.00], ["detached", "MR", "OR", 0.00], ["detached", "MR", "NR", 0.00], ["detached", "MR", "UF", 29.81],
    ["detached", "OR", "MR", 21.53], ["detached", "OR", "OR", 13.25], ["detached", "OR", "NR", 0.00], ["detached", "OR", "UF", 16.56],
    ["detached", "NR", "MR", 4.14], ["detached", "NR", "OR", 0.00], ["detached", "NR", "NR", 12.42], ["detached", "NR", "UF", 21.53],
])

# --- 表13 温度差係数 --------------------------------------------------------
# position: vert / top / floor_ne_uf_btm(r≠UF下面) / in_vert / in_btm / uf_btm
write("temp_diff_coef.csv", ["position", "value"], [
    ["vert", 1.0], ["top", 1.0], ["floor_ne_uf_btm", 0.7],
    ["in_vert", 0.0], ["in_btm", 0.0], ["uf_btm", 0.0],
])

# --- 表14・表15 方位係数（暖房期/冷房期） -----------------------------------
# model が使う方位: top, north, east, south, west, bottom
# top=1.0 / bottom=0.0（全地域）。8地域の暖房期方位係数は規定なし（暖房日数0
# のため加重に寄与しない）→ 0 を入れておく。
dir_rows = []
heat = {  # region -> {dir: value}（表14 北/東/南/西のみ抽出）
    1: dict(north=0.260, east=0.564, south=0.935, west=0.535),
    2: dict(north=0.263, east=0.554, south=0.856, west=0.544),
    3: dict(north=0.284, east=0.540, south=0.851, west=0.542),
    4: dict(north=0.256, east=0.531, south=0.815, west=0.527),
    5: dict(north=0.238, east=0.568, south=0.983, west=0.538),
    6: dict(north=0.261, east=0.579, south=0.936, west=0.523),
    7: dict(north=0.227, east=0.543, south=1.023, west=0.548),
    8: dict(north=0.0,   east=0.0,   south=0.0,   west=0.0),  # 暖房日数0
}
cool = {  # 表15
    1: dict(north=0.329, east=0.545, south=0.502, west=0.508),
    2: dict(north=0.341, east=0.503, south=0.507, west=0.529),
    3: dict(north=0.335, east=0.468, south=0.476, west=0.553),
    4: dict(north=0.322, east=0.518, south=0.437, west=0.481),
    5: dict(north=0.373, east=0.500, south=0.472, west=0.518),
    6: dict(north=0.341, east=0.512, south=0.434, west=0.504),
    7: dict(north=0.307, east=0.509, south=0.412, west=0.495),
    8: dict(north=0.325, east=0.515, south=0.480, west=0.505),
}
for season, table in [("heating", heat), ("cooling", cool)]:
    for region in range(1, 9):
        dir_rows.append([season, region, "top", 1.0])
        for d in ("north", "east", "south", "west"):
            dir_rows.append([season, region, d, table[region][d]])
        dir_rows.append([season, region, "bottom", 0.0])
write("direction_factor.csv", ["season", "region", "direction", "value"], dir_rows)

# --- 表4 暖冷房期間日数 -----------------------------------------------------
write("heating_cooling_days.csv", ["region", "heating_days", "cooling_days"], [
    [1, 257, 53], [2, 252, 48], [3, 244, 53], [4, 242, 53],
    [5, 218, 57], [6, 169, 117], [7, 122, 152], [8, 0, 265],
])

# --- 表16 仕様基準熱貫流率・線熱貫流率 --------------------------------------
# 地域グルーピング（1&2 / 3 / 4 / 5,6,7 / 8）を地域1〜8に展開
# part: roof / wall / floor / uf_wall(基礎壁) / uf_perimeter(土間床外周ψ)
#       roof_ex / roof_in / wall_ex / wall_in（集合）
def expand(g12, g3, g4, g567, g8):
    return {1: g12, 2: g12, 3: g3, 4: g4, 5: g567, 6: g567, 7: g567, 8: g8}


spec_u = {
    ("detached", "roof"):         expand(0.17, 0.24, 0.24, 0.24, 0.99),
    ("detached", "wall"):         expand(0.35, 0.53, 0.53, 0.53, 2.323),
    ("detached", "floor"):        expand(0.24, 0.24, 0.34, 0.34, 2.673),
    ("detached", "uf_wall"):      expand(0.27, 0.27, 0.52, 0.52, 4.443),
    ("detached", "uf_perimeter"): expand(1.01, 1.01, 1.05, 1.05, 1.05),
    ("apartment", "roof_ex"):     expand(0.38, 0.55, 0.75, 0.92, 1.18),
    ("apartment", "roof_in"):     expand(3.653, 3.653, 3.653, 3.653, 3.653),
    ("apartment", "wall_ex"):     expand(0.47, 0.70, 0.97, 0.97, 4.273),
    ("apartment", "wall_in"):     expand(4.273, 4.273, 4.273, 4.273, 4.273),
    # 集合の床は全て界床のため8地域は規定なし→無断熱床 2.540 を仮置き（通常未到達）
    ("apartment", "floor"):       expand(0.44, 0.61, 0.81, 0.98, 2.540),
}
# 集合の roof/wall は外気接側を roof/wall として参照（命名統一のため別名も付与）
spec_u[("apartment", "roof")] = spec_u[("apartment", "roof_ex")]
spec_u[("apartment", "wall")] = spec_u[("apartment", "wall_ex")]
# 窓・ドア（共通, 1&2 / 3 / 4 / 5,6,7 / 8 → 2.3,2.3,3.5,4.7,6.516）
for bt in ("detached", "apartment"):
    spec_u[(bt, "window")] = expand(2.3, 2.3, 3.5, 4.7, 6.516)
    spec_u[(bt, "door")] = expand(2.3, 2.3, 3.5, 4.7, 6.516)

u_rows = []
for (bt, part), d in spec_u.items():
    for region in range(1, 9):
        u_rows.append([bt, part, region, d[region]])
write("spec_u_value.csv", ["building_type", "part", "region", "u_value"], u_rows)

# --- 表17・表18・表19 壁体構成 ----------------------------------------------
# role: surface(表面熱伝達抵抗 Ri/Ro), air(中空層), material(材料), insulation(可変断熱材)
# 列: building_type, part, order, name, role, d[m], lambda[W/mK], c[kJ/m3K],
#     R[m2K/W], C[kJ/m2K]
# - surface/air: R,C を直接指定（d,lambda,c は空欄）
# - material  : d,lambda,c を指定（R=d/λ, C=d*c は reference_data.py で算出可だが
#               仕様書記載値をそのまま保持）
# - insulation: lambda,c を指定、d は U値から逆算（CSV上は空欄）
L = []  # rows


def lay(bt, part, order, name, role, d="", lam="", c="", R="", C=""):
    L.append([bt, part, order, name, role, d, lam, c, R, C])


# 表17 戸建住宅
# 外気に接する外壁（断熱なしR合計0.431, U=2.320）
lay("detached", "wall", 0, "Ri", "surface", R=0.110, C=0.0)
lay("detached", "wall", 1, "gypsum_board", "material", d=0.010, lam=0.220, c=830, R=0.047, C=8.638)
lay("detached", "wall", 2, "air_gap", "air", R=0.070, C=0.0)
lay("detached", "wall", 3, "glasswool_16K", "insulation", lam=0.045, c=13)
lay("detached", "wall", 4, "plywood", "material", d=0.012, lam=0.160, c=720, R=0.075, C=8.640)
lay("detached", "wall", 5, "cement_board", "material", d=0.013, lam=0.150, c=1000, R=0.088, C=13.235)
lay("detached", "wall", 6, "Ro", "surface", R=0.040, C=0.0)
# 外気に接する屋根（0.227, U=4.405）
lay("detached", "roof", 0, "Ri", "surface", R=0.090, C=0.0)
lay("detached", "roof", 1, "gypsum_board", "material", d=0.010, lam=0.220, c=830, R=0.047, C=8.638)
lay("detached", "roof", 2, "glasswool_10K", "insulation", lam=0.050, c=8)
lay("detached", "roof", 3, "Ro", "surface", R=0.090, C=0.0)
# 外気に接する床（0.375, U=2.667）
lay("detached", "floor", 0, "Ri", "surface", R=0.150, C=0.0)
lay("detached", "floor", 1, "plywood", "material", d=0.012, lam=0.160, c=720, R=0.075, C=8.640)
lay("detached", "floor", 2, "glasswool_16K", "insulation", lam=0.045, c=13)
lay("detached", "floor", 3, "Ro", "surface", R=0.150, C=0.0)
# 内壁床（0.880, U=1.136）
lay("detached", "inner_floor", 0, "Ri", "surface", R=0.150, C=0.0)
lay("detached", "inner_floor", 1, "plywood", "material", d=0.022, lam=0.160, c=720, R=0.138, C=15.84)
lay("detached", "inner_floor", 2, "air_gap", "air", R=0.070, C=0.0)
lay("detached", "inner_floor", 3, "gypsum_board", "material", d=0.095, lam=0.220, c=830, R=0.432, C=78.85)
lay("detached", "inner_floor", 4, "Ro", "surface", R=0.090, C=0.0)
# 基礎壁（0.225, U=4.444）
lay("detached", "uf_wall", 0, "Ri", "surface", R=0.110, C=0.0)
lay("detached", "uf_wall", 1, "phenolic_foam", "insulation", lam=0.022, c=77)
lay("detached", "uf_wall", 2, "concrete", "material", d=0.120, lam=1.600, c=2000, R=0.075, C=240.0)
lay("detached", "uf_wall", 3, "Ro", "surface", R=0.040, C=0.0)
# 土間床中央部（地盤境界, R=0.150+0.075）
lay("detached", "ground_floor", 0, "Ri", "surface", R=0.150, C=0.0)
lay("detached", "ground_floor", 1, "concrete", "material", d=0.120, lam=1.600, c=2000, R=0.075, C=240.0)

# 表18 集合住宅
lay("apartment", "wall", 0, "Ri", "surface", R=0.110, C=0.0)
lay("apartment", "wall", 1, "urethane_foam_A1", "insulation", lam=0.034, c=61)
lay("apartment", "wall", 2, "concrete", "material", d=0.135, lam=1.600, c=2000, R=0.084, C=270.0)
lay("apartment", "wall", 3, "Ro", "surface", R=0.040, C=0.0)
# 外気に接しない外壁（固定, U=4.267）
lay("apartment", "wall_in", 0, "Ri", "surface", R=0.110, C=0.0)
lay("apartment", "wall_in", 1, "concrete", "material", d=0.135, lam=1.600, c=2000, R=0.084, C=270.0)
lay("apartment", "wall_in", 2, "Ro", "surface", R=0.040, C=0.0)
# 外気に接する屋根（0.274, U=3.653）
lay("apartment", "roof", 0, "Ri", "surface", R=0.090, C=0.0)
lay("apartment", "roof", 1, "urethane_foam_2_1", "insulation", lam=0.023, c=60)
lay("apartment", "roof", 2, "concrete", "material", d=0.150, lam=1.600, c=2000, R=0.094, C=300.0)
lay("apartment", "roof", 3, "Ro", "surface", R=0.090, C=0.0)
# 外気に接しない屋根（固定, U=3.653）
lay("apartment", "roof_in", 0, "Ri", "surface", R=0.090, C=0.0)
lay("apartment", "roof_in", 1, "concrete", "material", d=0.150, lam=1.600, c=2000, R=0.094, C=300.0)
lay("apartment", "roof_in", 2, "Ro", "surface", R=0.090, C=0.0)
# 外気に接しない床（界床, 断熱なしR=0.394, U=2.540）
lay("apartment", "floor_in", 0, "Ri", "surface", R=0.150, C=0.0)
lay("apartment", "floor_in", 1, "concrete", "material", d=0.150, lam=1.600, c=2000, R=0.094, C=300.0)
lay("apartment", "floor_in", 2, "Ro", "surface", R=0.150, C=0.0)
# 集合住宅の外気接床は基本的に存在しない（界床）が、念のため界床と同構成を floor にも割当
lay("apartment", "floor", 0, "Ri", "surface", R=0.150, C=0.0)
lay("apartment", "floor", 1, "concrete", "material", d=0.150, lam=1.600, c=2000, R=0.094, C=300.0)
lay("apartment", "floor", 2, "Ro", "surface", R=0.150, C=0.0)

# 表19 共通：間仕切り（両 building_type 共通, R=0.401, U=2.494）
for bt in ("detached", "apartment"):
    lay(bt, "partition", 0, "Ri", "surface", R=0.110, C=0.0)
    lay(bt, "partition", 1, "gypsum_board", "material", d=0.012, lam=0.220, c=830, R=0.0555, C=9.960)
    lay(bt, "partition", 2, "air_gap", "air", R=0.070, C=0.0)
    lay(bt, "partition", 3, "gypsum_board", "material", d=0.012, lam=0.220, c=830, R=0.0555, C=9.960)
    lay(bt, "partition", 4, "Ro", "surface", R=0.110, C=0.0)

write("wall_layers.csv",
      ["building_type", "part", "order", "name", "role", "d", "lambda", "c", "R", "C"], L)

# --- 3.11 室内側対流熱伝達率 [W/m2K] ----------------------------------------
# part_type: wall(壁・窓・ドア) / roof(屋根・天井) / floor(床)
write("inside_convective_htc.csv", ["part_type", "h_c"], [
    ["wall", 2.5], ["roof", 5.0], ["floor", 0.7],
])

print("CSV を", DATA, "に書き出しました。")
for fn in sorted(os.listdir(DATA)):
    print("  -", fn)
