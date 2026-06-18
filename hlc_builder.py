# -*- coding: utf-8 -*-
"""
hlc_builder.py
==============================================================================
HeatLoadModel から Heat Load Calc の入力 Dictionary / JSON を生成する。
仕様書 3.5（共通出力事項）および hc-energy 02_02_spec_input に準拠。

外皮一般部位は layers で構成（表面熱伝達抵抗は含めず、Ro は
outside_heat_transfer_resistance=0.04 で別途指定。中空層は layers に含める）。
窓・ドアは u_value / eta_value を直接指定する。
==============================================================================
"""
import json
import reference_data as ref

VERT = ref.VERT_DIRS
DIR_HLC = {"north": "n", "east": "e", "south": "s", "west": "w",
           "top": "top", "bottom": "bottom"}
ROOM_ID = {"MR": 0, "OR": 1, "NR": 2, "UF": 3}  # 3.5.3 正準ID（除外が無い場合の基準）
ROOM_ORDER = ["MR", "OR", "NR", "UF"]            # room_id を割り当てる正準順
ROOM_NAME = {"MR": "main_occupant_room", "OR": "other_occupant_room",
             "NR": "non_occupant_room", "UF": "under_floor"}
SCHEDULE = {"MR": "main_occupant_room", "OR": "other_occupant_room",
            "NR": "non_occupant_room", "UF": "zero"}
HC_WALL, HC_ROOF, HC_FLOOR = None, None, None  # 初期化時にCSVから


def room_id_map(m):
    """モデルに実在する室(m.spaces)へ 0 始まりの連番 room_id を割り当てて返す。

    3.4.2 の「室数の設定」により、床面積0の室は m.spaces から除外されている
    （入力パラメータによって室数は3室/4室、さらに除外により2室等にもなる）。
    heat_load_calc は room_id が 0 始まりの連番であることを前提とするため、
    正準順 (MR→OR→NR→UF) を保ったまま詰め直して連番化する（3.5.3）。
      例) NR が除外された戸建基礎断熱 → {"MR": 0, "OR": 1, "UF": 2}
      例) OR が除外された集合住宅     → {"MR": 0, "NR": 1}
    境界の connected_room_id・換気 route・設備 space_id は、すべて本マップを
    通して採番すること（固定の ROOM_ID を直接使わない）。
    """
    ordered = [s for s in ROOM_ORDER if s in m.spaces]
    return {s: i for i, s in enumerate(ordered)}

R_SO = 0.04   # 3.5.4.1 室外側熱伝達抵抗
R_SI_WIN = 0.11  # 3.5.4.2 窓・ドアの室内側熱抵抗
EMISS = 0.9   # 3.5.4.5
SOLAR_ABS = 0.8  # 3.5.4.4


def _hc(part_type):
    return ref.hc(part_type)


def _layers_for(bt, part, d_ins=None):
    """HLC用 layer リスト（室内側→室外側、表面熱伝達抵抗を除く）を返す。"""
    out = []
    for l in ref.layers(bt, part):
        if l["role"] == "surface":
            continue  # Ri/Ro は含めない
        if l["role"] == "insulation":
            if d_ins is None or d_ins <= 0:
                continue
            R = d_ins / l["lambda"]
            C = d_ins * l["c"]  # c[kJ/m3K]*d[m] = kJ/m2K
        else:  # material / air
            R = l["R"]
            C = l["C"] if l["C"] is not None else 0.0
        out.append({"name": l["name"],
                    "thermal_resistance": round(R, 6),
                    "thermal_capacity": round(C, 4)})
    return out


def build(m):
    bt = m.bt
    # 3.4.2/3.5.3：実在する室にのみ 0始まりの連番 room_id を割り当てる
    rid = room_id_map(m)
    boundaries = []
    bid = [0]

    def next_id():
        i = bid[0]
        bid[0] += 1
        return i

    def add(b):
        boundaries.append(b)
        return b["id"]

    # ---- 一般外皮（外気に接する: 屋根/外壁/床/基礎壁）-----------------------
    def gen_ext(s, d, area, part, htype):
        if area <= 1e-9:
            return
        d_ins = m.d_ins.get(part)
        add({
            "id": next_id(),
            "name": f"{ROOM_NAME[s]}_{part}_{d}",
            "sub_name": "",
            "connected_room_id": rid[s],
            "boundary_type": "external_general_part",
            "area": round(area, 4),
            "is_sun_striked_outside": True,
            "temp_dif_coef": _tdc_for(d, exterior=True),
            "is_solar_absorbed_inside": (htype == "floor"),
            "is_floor": (htype == "floor"),
            "direction": DIR_HLC[d],
            "h_c": _hc(htype),
            "outside_emissivity": EMISS,
            "inside_emissivity": EMISS,
            "outside_heat_transfer_resistance": R_SO,
            "outside_solar_absorption": SOLAR_ABS,
            "layers": _layers_for(bt, part, d_ins),
            "solar_shading_part": {"existence": False},
        })

    for s in m.spaces:
        # 屋根（上面）
        gen_ext(s, "top", m.A_wall_ex[s]["top"], "roof", "roof")
        # 外壁（垂直, UFは基礎壁）
        for d in VERT:
            part = "uf_wall" if s == "UF" else "wall"
            gen_ext(s, d, m.A_wall_ex[s][d], part, "wall")
        # 床（下面）: 床断熱/集合のみ external_general_part として外気床
        if s != "UF" and not m.has_uf:
            gen_ext(s, "bottom", m.A_wall_ex[s]["bottom"], "floor", "floor")

    # ---- 外気に接しない外皮（界壁・界床: temp_dif_coef 0）------------------
    def gen_nonext(s, d, area, part, htype):
        if area <= 1e-9:
            return
        add({
            "id": next_id(),
            "name": f"{ROOM_NAME[s]}_nonext_{part}_{d}",
            "sub_name": "",
            "connected_room_id": rid[s],
            "boundary_type": "external_general_part",
            "area": round(area, 4),
            "is_sun_striked_outside": False,
            "temp_dif_coef": 0.0,
            "is_solar_absorbed_inside": (htype == "floor"),
            "is_floor": (htype == "floor"),
            "h_c": _hc(htype),
            "inside_emissivity": EMISS,
            "layers": _layers_for(bt, part),
            "solar_shading_part": {"existence": False},
        })

    if bt == "apartment":
        for s in m.spaces:
            gen_nonext(s, "top", m.A_in[s]["top"], "roof_in", "roof")
            for d in VERT:
                gen_nonext(s, d, m.A_in[s][d], "wall_in", "wall")
            gen_nonext(s, "bottom", m.A_in[s]["bottom"], "floor_in", "floor")

    # ---- 窓（透明開口部）---------------------------------------------------
    for s in m.spaces:
        for d in VERT:
            area = m.A_win_mod[s][d]
            if area <= 1e-9:
                continue
            add({
                "id": next_id(),
                "name": f"{ROOM_NAME[s]}_window_{d}",
                "sub_name": "",
                "connected_room_id": rid[s],
                "boundary_type": "external_transparent_part",
                "area": round(area, 4),
                "is_sun_striked_outside": True,
                "temp_dif_coef": 1.0,
                "is_solar_absorbed_inside": False,
                "is_floor": False,
                "direction": DIR_HLC[d],
                "h_c": _hc("wall"),
                "outside_emissivity": EMISS,
                "outside_heat_transfer_resistance": R_SO,
                "eta_value": round(m.eta_win, 4),
                "u_value": round(m.U_win_mod[s][d], 4),
                "inside_heat_transfer_resistance": R_SI_WIN,
                "glass_area_ratio": m.glass_area_ratio,
                "incident_angle_characteristics": "multiple",
                "solar_shading_part": {"existence": False},
            })

    # ---- ドア（非透明開口部）----------------------------------------------
    for s in m.spaces:
        for d in VERT:
            area = m.A_door[s][d]
            if area <= 1e-9:
                continue
            add({
                "id": next_id(),
                "name": f"{ROOM_NAME[s]}_door_{d}",
                "sub_name": "",
                "connected_room_id": rid[s],
                "boundary_type": "external_opaque_part",
                "area": round(area, 4),
                "is_sun_striked_outside": True,
                "temp_dif_coef": 1.0,
                "is_solar_absorbed_inside": False,
                "is_floor": False,
                "direction": DIR_HLC[d],
                "h_c": _hc("wall"),
                "outside_emissivity": EMISS,
                "outside_heat_transfer_resistance": R_SO,
                "u_value": round(m.U["door"], 4),
                "inside_heat_transfer_resistance": R_SI_WIN,
                "outside_solar_absorption": SOLAR_ABS,
                "solar_shading_part": {"existence": False},
            })

    # ---- 間仕切り（両面ペア internal）-------------------------------------
    for (r1, r2), area in m.partition.items():
        if area <= 1e-9:
            continue
        # 3.4.2 で除外された室への間仕切りは生成しない（防御的ガード）
        if r1 not in rid or r2 not in rid:
            continue
        id_a = next_id()
        id_b = next_id()
        lay = _layers_for(bt, "partition")
        add({"id": id_a, "name": f"partition_{r1}_to_{r2}", "sub_name": "",
             "connected_room_id": rid[r1], "boundary_type": "internal",
             "area": round(area, 4), "rear_surface_boundary_id": id_b,
             "is_solar_absorbed_inside": False, "is_floor": False,
             "h_c": _hc("wall"), "inside_emissivity": EMISS, "layers": lay})
        add({"id": id_b, "name": f"partition_{r2}_to_{r1}", "sub_name": "",
             "connected_room_id": rid[r2], "boundary_type": "internal",
             "area": round(area, 4), "rear_surface_boundary_id": id_a,
             "is_solar_absorbed_inside": False, "is_floor": False,
             "h_c": _hc("wall"), "inside_emissivity": EMISS, "layers": list(reversed(lay))})

    # ---- 内壁床 ------------------------------------------------------------
    # 同一室用途間(MR→MR等): 温度差係数0の外気に接する床(external_general_part)
    # 異室用途間 / 居室↔UF: 間仕切り床(internal 両面ペア)
    for (r1, r2), area in m.inner_floor.items():
        if area <= 1e-9:
            continue
        # 3.4.2 で除外された室に接する内壁床は生成しない（防御的ガード）
        if r1 not in rid or r2 not in rid:
            continue
        if r1 == r2:
            add({"id": next_id(), "name": f"inner_floor_{r1}_self", "sub_name": "",
                 "connected_room_id": rid[r1], "boundary_type": "external_general_part",
                 "area": round(area, 4), "is_sun_striked_outside": False,
                 "temp_dif_coef": 0.0, "is_solar_absorbed_inside": True, "is_floor": True,
                 "h_c": _hc("floor"), "inside_emissivity": EMISS,
                 "layers": _layers_for(bt, "inner_floor"),
                 "solar_shading_part": {"existence": False}})
        else:
            # r1 の床（下面）が r2 に接する。裏面は r2 側の天井。
            id_a = next_id()
            id_b = next_id()
            lay = _layers_for(bt, "inner_floor")
            add({"id": id_a, "name": f"inner_floor_{r1}_to_{r2}", "sub_name": "",
                 "connected_room_id": rid[r1], "boundary_type": "internal",
                 "area": round(area, 4), "rear_surface_boundary_id": id_b,
                 "is_solar_absorbed_inside": True, "is_floor": True,
                 "h_c": _hc("floor"), "inside_emissivity": EMISS, "layers": lay})
            add({"id": id_b, "name": f"inner_floor_{r2}_to_{r1}", "sub_name": "",
                 "connected_room_id": rid[r2], "boundary_type": "internal",
                 "area": round(area, 4), "rear_surface_boundary_id": id_a,
                 "is_solar_absorbed_inside": False, "is_floor": False,
                 "h_c": _hc("wall"), "inside_emissivity": EMISS, "layers": list(reversed(lay))})

    # ---- 土間床中央部（基礎断熱: ground 境界）-----------------------------
    if m.has_uf and m.A_UF > 1e-9:
        add({"id": next_id(), "name": "under_floor_earth", "sub_name": "",
             "connected_room_id": rid["UF"], "boundary_type": "ground",
             "area": round(m.A_UF, 4), "is_solar_absorbed_inside": True, "is_floor": True,
             "h_c": _hc("floor"), "inside_emissivity": EMISS,
             "layers": _layers_for(bt, "ground_floor")})

    # ---- rooms -------------------------------------------------------------
    # 3.4.2 で実在する室のみを、rid の連番 room_id で出力する。
    rooms = []
    for s in m.spaces:
        rooms.append({
            "id": rid[s], "name": ROOM_NAME[s], "sub_name": "",
            "floor_area": round(m.A.get(s, m.A_UF if s == "UF" else 0.0), 4),
            "volume": round(m.V[s], 4),
            "ventilation": {"natural": round(m.Q_ntrl[s], 4)},
            "furniture": {"input_method": "default"},
            "schedule": {"name": SCHEDULE[s]},
        })

    # ---- mechanical_ventilations（3.5.5）----------------------------------
    # 仕様3.5.5は MR起点(route 0,2) と OR起点(route 1,2) の第3種換気2系統を定義
    # （いずれも非居室NRを経由して排気）。3.4.2 で居室が除外され得るため、
    #   ・給気起点(MR/OR)が実在する系統のみ生成
    #   ・経由先の NR が実在する場合のみ route に NR を含める
    # とする。room_id は rid で採番し、給気量は当該室容積の0.5倍（仕様3.5.5）。
    # ※起点室や経由NRが除外された場合の扱いは仕様に明記が無いため、上記は
    #   「存在する室のみで経路を張る」という保守的な解釈（要確認）。
    mvs = []
    mv_id = 0
    for src in ("MR", "OR"):
        if src not in rid:
            continue
        route = [rid[src]] + ([rid["NR"]] if "NR" in rid else [])
        mvs.append({"id": mv_id, "root_type": "type3",
                    "volume": round(m.V[src] * 0.5, 4), "route": route})
        mv_id += 1

    # ---- equipments（3.5.6）-----------------------------------------------
    import math

    def heat_qmax(A):
        return 1.2090 * (190.5 * A + 45.6) - 85.1

    def cool_qmax(A):
        return 190.5 * A + 45.6

    def vmin(qmax):
        return 7.8574 * math.exp(0.0537 * qmax / 1000.0)

    def vmax(qmax):
        return 11.076 * (qmax / 1000.0) ** 0.3432

    def equip(name, sid, qmax):
        return {"id": sid, "name": name, "equipment_type": "rac",
                "property": {"space_id": sid, "q_min": 500.0, "q_max": round(qmax, 2),
                             "v_min": round(vmin(qmax), 4), "v_max": round(vmax(qmax), 4),
                             "bf": 0.2}}

    # 仕様3.5.6は MR・OR に暖房機器/冷房機器を各1台定義（space_id=各室のID）。
    # 3.4.2 で MR/OR が0m2により除外された場合は、その室向け機器を生成しない。
    # space_id は rid で採番（仕様3.5.6のOR冷房 space_id は誤記と判断し当該室IDを使用）。
    equipments = {"heating_equipments": [], "cooling_equipments": []}
    for src in ("MR", "OR"):
        if src not in rid:
            continue
        sid = rid[src]
        equipments["heating_equipments"].append(
            equip(f"heating_equipment for {src}", sid, heat_qmax(m.A[src])))
        equipments["cooling_equipments"].append(
            equip(f"cooling_equipment for {src}", sid, cool_qmax(m.A[src])))

    common = {
        "ac_method": "air_temperature",
        "weather": {"method": "ees", "region": str(m.region)},
        "mutual_radiation_method": "Nagata",
    }
    building = {"infiltration": {
        "method": "balance_residential",
        "story": 2 if bt == "detached" else 1,
        "c_value_estimate": "specify", "c_value": 0.0,
        "inside_pressure": "negative",
    }}

    return {
        "common": common, "building": building, "rooms": rooms,
        "boundaries": boundaries, "mechanical_ventilations": mvs,
        "equipments": equipments,
    }


def _tdc_for(d, exterior):
    if not exterior:
        return 0.0
    if d == "top":
        return ref.temp_diff("top")          # 1.0
    if d == "bottom":
        return ref.temp_diff("floor_ne_uf_btm")  # 0.7
    return ref.temp_diff("vert")             # 1.0


def to_json(model_dict, path):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(model_dict, f, ensure_ascii=False, indent=2)
