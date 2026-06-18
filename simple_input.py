# -*- coding: utf-8 -*-
"""
simple_input.py
==============================================================================
仕様書「簡易入力からの暖冷房負荷モデルの構築」3.4 に準拠し、少数の設計住戸
入力パラメータから、動的熱負荷計算エンジン（Heat Load Calc）が必要とする
詳細な暖冷房負荷モデルを構築するモジュール。

メソッドは仕様書の節番号に対応させている（例: _s3_4_3_2 = 3.4.3.2）。
参照データは reference_data 経由で CSV から取得し、本ファイルに数値の
ハードコードは置かない（窓η上下限など仕様本文中の固定値を除く）。
==============================================================================
"""
import math
import reference_data as ref

VERT = ref.VERT_DIRS                 # ["north","east","south","west"]
SPACES = ref.SPACES                  # ["MR","OR","NR"]

# 3.4.10 窓の垂直入射時日射熱取得率の上下限（仕様本文の固定値）
ETA_WIN_MIN = 0.10
ETA_WIN_MAX = 0.73
GLASS_AREA_RATIO = 0.8
# 3.4.9 不透明部位の日射取得換算係数（α・R_so 相当の固定値）
M_OPAQUE_COEF = 0.034
# 集合住宅 r_env,ex ロジスティック係数（3.4.3.4.1）
REX_A = 9.10907512
REX_B = 1.05204145
# 開口部比 r_env,op ロジスティック係数（3.4.3.5.1）
ROP_A = 0.129
ROP_B = 13.35


class HeatLoadModel:
    def __init__(self, building_type, A_MR, A_OR, A_A, region, A_env, UA,
                 eta_AC, eta_AH, is_floor_ins=True, n_r=None):
        self.bt = building_type            # "detached" / "apartment"
        self.region = int(region)
        self.A_env_input = float(A_env)
        self.UA = float(UA)
        self.eta_AC = float(eta_AC)
        self.eta_AH = float(eta_AH)
        # 集合住宅は基礎断熱の概念を持たない（床=界床）。UF ゾーンは戸建基礎断熱のみ。
        self.is_floor_ins = True if building_type == "apartment" else bool(is_floor_ins)
        self.has_uf = (building_type == "detached" and not self.is_floor_ins)
        self.ref = ref.get_reference(building_type, self.is_floor_ins)

        # ---- 非居室の床面積（3.4.3.1）------------------------------------
        # 非居室 A_NR = max(A_A - A_MR - A_OR, 0)
        # （主たる居室・その他居室を除いた残りを非居室とみなす。負値は0で頭打ち）
        A_NR = max(A_A - A_MR - A_OR, 0.0)
        self.A = {"MR": float(A_MR), "OR": float(A_OR), "NR": A_NR}
        self.A_A = float(A_A)

        # ---- 出力する室数の設定（3.4.2）----------------------------------
        # 入力パラメータによって暖冷房負荷モデルの室数が変わる。
        #   3.4.2.1 戸建＋床断熱 / 集合住宅 → 主たる居室・その他居室・非居室の3室
        #   3.4.2.2 戸建＋基礎断熱        → 上記3室＋床下空間(UF)の4室
        # かつ、いずれの場合も「室用途ごとの床面積が0m2となる室は除く」
        # （3.4.2.1・3.4.2.2 のただし書き）。
        #   occ_spaces : 床面積>0 の居室(MR/OR/NR)のみを残した「有効居室」リスト
        #   spaces     : occ_spaces に、戸建基礎断熱のときだけ UF を加えた全空間リスト
        # 以降の各推定メソッド(_s3_4_*)と境界生成(hlc_builder)は、すべて
        # self.spaces / self.occ_spaces を基準に回し、除外された室は一切生成しない。
        # UF(床下空間)は has_uf(戸建基礎断熱)のときのみ存在し、その床面積は常に>0
        # のため、UFは床面積0による除外対象としない。
        self.occ_spaces = [s for s in SPACES if self.A[s] > 0.0]
        self.spaces = list(self.occ_spaces) + (["UF"] if self.has_uf else [])
        self.nu = ref.direction_factor(self.region)
        self.n_h, self.n_c = ref.hc_days(self.region)
        # 自然風換気回数（未指定時の既定値: 居室5回/h, 床下0回/h）
        default_n = {s: 5.0 for s in SPACES}
        default_n["UF"] = 0.0
        self.n_r = {**default_n, **(n_r or {})}

        self._build()

    # 温度差係数ショートカット
    def _H(self, key):
        return ref.temp_diff(key)

    def _build(self):
        self._s3_4_3_2()   # 水平外皮（上面・下面）
        self._s3_4_3_3()   # 垂直外皮
        self._s3_4_3_4()   # 外気に接する外皮
        self._s3_4_3_5()   # 開口部（窓・ドア）
        self._s3_4_3_6_7() # 外壁等・外気に接しない外皮
        self._s3_4_3_8()   # 土間床外周長
        self._s3_4_3_9()   # 間仕切り・内壁床
        self._s3_4_4()     # 室容積
        self._s3_4_5()     # 自然風換気量
        self._s3_4_6()     # 熱貫流率
        self._s3_4_7()     # 断熱材厚さ
        self._s3_4_8_9_10() # 平均日射熱取得率目標・窓η・補正

    # ---- 3.4.3.2 水平外皮（上面・下面）------------------------------------
    def _s3_4_3_2(self):
        self.A_top = {}
        self.A_btm = {}
        for s in self.spaces:
            Ar = self.A.get(s, 0.0)
            Aref = self.ref.floor.get(s, 0.0)
            # 上面（3.4.3.2.1）: UF は 0
            if s == "UF":
                self.A_top[s] = 0.0
            else:
                self.A_top[s] = self.ref.env("top", s) * (Ar / Aref) if Aref > 0 else 0.0
            # 下面（3.4.3.2.2）
            if s == "UF":
                self.A_btm[s] = self.ref.env("bottom", "UF") * (self.A_A / self.ref.A_A)
            else:
                self.A_btm[s] = self.ref.env("bottom", s) * (Ar / Aref) if Aref > 0 else 0.0

    # ---- 3.4.3.3 垂直外皮 --------------------------------------------------
    def _s3_4_3_3(self):
        # 3.4.3.3.1 床下空間の方位別垂直外皮（基礎壁）
        self.A_vert = {s: {d: 0.0 for d in VERT} for s in self.spaces}
        if self.has_uf:
            ref_uf_btm = self.ref.env("bottom", "UF")
            ratio = (self.A_btm["UF"] / ref_uf_btm) if ref_uf_btm > 0 else 0.0
            for d in VERT:
                self.A_vert["UF"][d] = self.ref.env(d, "UF") * ratio
        # 3.4.3.3.2 居室の総垂直外皮
        sum_top = sum(self.A_top.values())
        sum_btm = sum(self.A_btm.values())
        sum_uf_vert = sum(self.A_vert["UF"].values()) if self.has_uf else 0.0
        self.A_env_vert_total = max(self.A_env_input - sum_top - sum_btm - sum_uf_vert, 0.0)
        # 居室(MR/OR/NR)のうち有効なものだけを対象とする（3.4.2 で除外済みの
        # 0m2居室は self.A_vert に枠が無いため、必ず occ_spaces で回す）。
        for s in self.occ_spaces:
            A_vert_s = self.A_env_vert_total * (self.A[s] / self.A_A) if self.A_A > 0 else 0.0
            ref_vert_sum = self.ref.vert_total(s)
            # 3.4.3.3.3 方位別
            for d in VERT:
                self.A_vert[s][d] = (A_vert_s * self.ref.env(d, s) / ref_vert_sum
                                     if ref_vert_sum > 0 else 0.0)

    def _vert_total_model(self, s):
        return sum(self.A_vert[s].values())

    # 空間ごと方位ごとの「外皮面積」A_env,r,d（top/vert/bottom）
    def _A_env_rd(self, s, d):
        if d == "top":
            return self.A_top[s]
        if d == "bottom":
            return self.A_btm[s]
        return self.A_vert[s][d]

    # ---- 3.4.3.4 外気に接する外皮 ------------------------------------------
    def _s3_4_3_4(self):
        sum_east = sum(self.A_vert[s]["east"] for s in self.spaces)
        sum_west = sum(self.A_vert[s]["west"] for s in self.spaces)
        sum_top = sum(self.A_top.values())
        sum_vert = sum(self._vert_total_model(s) for s in self.spaces)
        Aenv = self.A_env_input
        r_min = (sum_east + sum_west) / Aenv if Aenv > 0 else 0.0
        r_max = (sum_top + sum_vert) / Aenv if Aenv > 0 else 0.0
        if self.bt == "apartment":
            logistic = 1.0 / (1.0 + math.exp(-REX_A * (self.UA - REX_B)))
            r_ex = min(max(logistic, r_min), r_max)
        else:
            r_ex = 1.0
        self.r_env_ex = r_ex
        self.A_env_ex = Aenv * r_ex

        # 3.4.3.4.2 空間ごと方位ごとの外気接外皮
        A_var = (sum_top
                 + sum(self.A_vert[s]["south"] for s in self.spaces)
                 + sum(self.A_vert[s]["north"] for s in self.spaces))
        if self.bt == "apartment":
            if A_var > 0:
                r_var = min(max((self.A_env_ex - sum_east - sum_west) / A_var, 0.0), 1.0)
            else:
                r_var = 0.0
        else:
            r_var = 1.0
        self.r_var = r_var

        # 表3 の外気接割合 r_env,ex,r,d
        self.A_env_rd_ex = {s: {} for s in self.spaces}
        for s in self.spaces:
            for d in ["top"] + VERT + ["bottom"]:
                A_rd = self._A_env_rd(s, d)
                if s == "UF":
                    # 集合にUFは無い。戸建基礎断熱のUFは全方位1.0
                    frac = 1.0 if self.bt == "detached" else 0.0
                else:
                    if self.bt == "detached":
                        frac = 1.0
                    else:  # apartment 居室
                        if d in ("east", "west"):
                            frac = 1.0
                        elif d == "bottom":
                            frac = 0.0
                        else:  # top, south, north
                            frac = r_var
                self.A_env_rd_ex[s][d] = A_rd * frac

    # ---- 3.4.3.5 開口部 ----------------------------------------------------
    def _s3_4_3_5(self):
        # 3.4.3.5.1 総開口部面積
        r_op = 1.0 / (1.0 + math.exp(-ROP_A * (self.eta_AC - ROP_B)))
        self.r_env_op = r_op
        self.A_env_op = self.A_env_ex * r_op

        # 3.4.3.5.2 空間ごと開口部面積
        def op_ref(s):  # 参照住戸の空間別 窓+ドア合計
            return (sum(self.ref.win(d, s) for d in VERT)
                    + sum(self.ref.door(d, s) for d in VERT))
        weight = {}
        # 開口部は居室にのみ配分する。3.4.2 で除外された0m2居室は対象外。
        for s in self.occ_spaces:
            Aref = self.ref.floor.get(s, 0.0)
            weight[s] = self.A[s] * (op_ref(s) / Aref) if Aref > 0 else 0.0
        wsum = sum(weight.values())
        self.A_env_op_r = {s: (self.A_env_op * weight[s] / wsum if wsum > 0 else 0.0)
                           for s in self.occ_spaces}

        # 3.4.3.5.3 窓面積（方位別）
        self.A_win = {s: {d: 0.0 for d in VERT} for s in self.spaces}
        self.A_door = {s: {d: 0.0 for d in VERT} for s in self.spaces}
        for s in self.occ_spaces:
            denom = (sum(self.ref.win(d, s) for d in VERT)
                     + sum(self.ref.door(d, s) for d in VERT))
            for d in VERT:
                if denom > 0:
                    raw = self.A_env_op_r[s] * self.ref.win(d, s) / denom
                else:
                    raw = 0.0
                self.A_win[s][d] = min(raw, self.A_env_rd_ex[s][d])
            # 3.4.3.5.4 ドア面積（窓の後）
            for d in VERT:
                if denom > 0:
                    raw = self.A_env_op_r[s] * self.ref.door(d, s) / denom
                else:
                    raw = 0.0
                cap = max(self.A_env_rd_ex[s][d] - self.A_win[s][d], 0.0)
                self.A_door[s][d] = min(raw, cap)

    # ---- 3.4.3.6 / 3.4.3.7 外壁等・外気に接しない外皮 ----------------------
    def _s3_4_3_6_7(self):
        # 外気に接する外壁等（方位別, top/vert/bottom）
        self.A_wall_ex = {s: {} for s in self.spaces}
        for s in self.spaces:
            for d in ["top"] + VERT + ["bottom"]:
                win = self.A_win[s].get(d, 0.0) if d in VERT else 0.0
                door = self.A_door[s].get(d, 0.0) if d in VERT else 0.0
                self.A_wall_ex[s][d] = max(self.A_env_rd_ex[s][d] - win - door, 0.0)
        # 外気に接しない外皮（方位別）
        self.A_in = {s: {} for s in self.spaces}
        for s in self.spaces:
            for d in ["top"] + VERT + ["bottom"]:
                self.A_in[s][d] = max(self._A_env_rd(s, d) - self.A_env_rd_ex[s][d], 0.0)

    # ---- 3.4.3.8 土間床外周長 ----------------------------------------------
    def _s3_4_3_8(self):
        if not self.has_uf:
            self.L_uf_ex = 0.0
            return
        total = 0.0
        for d in VERT:
            ref_len = self.ref.uf_perimeter.get(d, 0.0)
            ref_wall = self.ref.env(d, "UF")          # 基礎壁(外気接=ref値)
            model_wall_ex = self.A_env_rd_ex["UF"][d]  # 戸建基礎断熱は frac=1.0
            if ref_wall > 0:
                total += ref_len * (model_wall_ex / ref_wall)
        self.L_uf_ex = total

    # ---- 3.4.3.9 間仕切り・内壁床 ------------------------------------------
    def _s3_4_3_9(self):
        # ---- 3.4.3.9.1 間仕切り壁 ----------------------------------------
        # 参照住戸の間仕切り面積に、暖冷房負荷モデルと参照住戸の両室合計垂直外皮
        # 面積比を乗じて求める。【3.4.3.9.1 ただし書き】空間 r1 または r2 が除外
        # されている場合は当該間仕切りはないものとするため、両室とも有効居室
        # (occ_spaces)であるペアのみを対象とする。
        self.partition = {}
        cand_pairs = [("MR", "OR"), ("MR", "NR"), ("OR", "NR")]
        pairs = [(r1, r2) for (r1, r2) in cand_pairs
                 if r1 in self.occ_spaces and r2 in self.occ_spaces]
        for (r1, r2) in pairs:
            ref_p = self.ref.partition_area(r1, r2)
            ref_v = self.ref.vert_total(r1) + self.ref.vert_total(r2)
            mdl_v = self._vert_total_model(r1) + self._vert_total_model(r2)
            self.partition[(r1, r2)] = ref_p * (mdl_v / ref_v) if ref_v > 0 else 0.0

        # ---- 3.4.3.9.2 内壁床 --------------------------------------------
        # 空間ごと総内壁床 = 設計住戸の床面積 - 暖冷房負荷モデルの下面外皮面積。
        # （A_part,btm,r = max(A_r - A_env,r,btm, 0)）
        self.A_inner_floor_space = {}  # 空間ごと総内壁床
        for s in self.spaces:
            self.A_inner_floor_space[s] = max(self.A.get(s, 0.0) - self.A_btm.get(s, 0.0), 0.0)
        # 参照住戸の (r1→r2) 内壁床比で r1 の総内壁床を按分する。
        # 【3.4.3.9.2 ただし書き】接続先 r2 が 3.4.2 で除外されている場合は、その面積を
        #   捨てず、除外されていない室 r1 の床面積が担保されるように r1 自身へ接続する
        #   内壁床（＝同一室用途間と同じ「温度差係数0の外気に接する床」, キー (r1,r1)）
        #   へ合算する。これにより r1 の床面積（=外気接下面+内壁床の総和）が保存される。
        #   なお r1 自身が除外された場合は A_inner_floor_space[r1] を生成しない
        #   （上側の室が無く面積0のため、生存室側の内壁天井も発生しない）。
        # 按分の分母 ref_tot は参照住戸の全接続先（除外室を含む）の合計を用いるので、
        #   r2 が除外されても r1 へ配分される総量は変わらず、面積保存が成立する。
        self.inner_floor = {}
        for r1 in self.spaces:
            ref_tot = self.ref.inner_floor_total(r1)   # 参照住戸の全r2合計（除外室含む）
            if ref_tot <= 0:
                continue
            for r2 in ref.ALL_SPACES:                  # 参照住戸の全接続先を走査
                ref_pair = self.ref.inner_floor_area(r1, r2)
                if ref_pair <= 0:
                    continue
                area = self.A_inner_floor_space[r1] * (ref_pair / ref_tot)
                # 接続先が実在すれば (r1,r2)、除外されていれば r1 自身 (r1,r1) へ合算
                key = (r1, r2) if r2 in self.spaces else (r1, r1)
                self.inner_floor[key] = self.inner_floor.get(key, 0.0) + area

    # ---- 3.4.4 室容積 ------------------------------------------------------
    def _s3_4_4(self):
        self.A_UF = self.A_btm.get("UF", 0.0)
        self.V = {}
        for s in self.spaces:
            if s == "UF":
                self.V[s] = 0.4 * self.A_UF
            else:
                self.V[s] = 2.4 * self.A[s]

    # ---- 3.4.5 自然風換気量 ------------------------------------------------
    def _s3_4_5(self):
        self.Q_ntrl = {s: self.V[s] * self.n_r.get(s, 0.0) for s in self.spaces}

    # ---- 3.4.6 熱貫流率 ----------------------------------------------------
    def _s3_4_6(self):
        # 3.4.6.1 外気接面積の集計
        self.A_roof_ex = sum(self.A_wall_ex[s]["top"] for s in self.spaces)
        self.A_wall_neuf_ex = sum(self.A_wall_ex[s][d] for s in self.spaces if s != "UF"
                                  for d in VERT)
        self.A_wall_uf_ex = sum(self.A_wall_ex["UF"][d] for d in VERT) if self.has_uf else 0.0
        if self.has_uf:
            self.A_floor_ex = sum(self._A_env_rd("UF", "bottom") for _ in [0])  # UF下面
            self.A_floor_ex = self.A_btm.get("UF", 0.0)
        else:
            self.A_floor_ex = sum(self.A_wall_ex[s]["bottom"] for s in self.spaces if s != "UF")
        self.A_win_ex = sum(self.A_win[s][d] for s in self.spaces for d in VERT)
        self.A_door_ex = sum(self.A_door[s][d] for s in self.spaces for d in VERT)

        H_top = self._H("top")
        H_vert = self._H("vert")
        H_floor = self._H("floor_ne_uf_btm")  # 0.7（床は一旦床断熱想定）

        def su(part):
            return ref.spec_u(self.bt, part, self.region)

        # 仕様基準熱貫流率での熱損失量
        q_spec = {
            "roof":  su("roof") * self.A_roof_ex * H_top,
            "floor": su("floor") * self.A_floor_ex * H_floor,
            "wall":  su("wall") * self.A_wall_neuf_ex * H_vert,
            "win":   su("window") * self.A_win_ex * H_vert,
            "door":  su("door") * self.A_door_ex * H_vert,
        }
        q_spec_all = sum(q_spec.values())
        q_target_all = self.UA * self.A_env_input
        self.q_target = {k: (q_target_all * v / q_spec_all if q_spec_all > 0 else 0.0)
                         for k, v in q_spec.items()}

        self.U = {}
        self.U["roof"] = self.q_target["roof"] / (H_top * self.A_roof_ex) if self.A_roof_ex > 0 else 0.0
        self.U["wall"] = self.q_target["wall"] / (H_vert * self.A_wall_neuf_ex) if self.A_wall_neuf_ex > 0 else 0.0
        self.U["win"] = self.q_target["win"] / (H_vert * self.A_win_ex) if self.A_win_ex > 0 else 0.0
        self.U["door"] = self.q_target["door"] / (H_vert * self.A_door_ex) if self.A_door_ex > 0 else 0.0

        if self.has_uf:
            # 3.4.6.2.2 基礎断熱: 基礎壁 + 土間床外周ψ に按分
            q_spec_uf_wall = su("uf_wall") * self.A_wall_uf_ex * H_vert
            q_spec_uf_hb = su("uf_perimeter") * self.L_uf_ex * H_vert
            denom = q_spec_uf_wall + q_spec_uf_hb
            qt_uf_wall = self.q_target["floor"] * (q_spec_uf_wall / denom) if denom > 0 else 0.0
            qt_uf_hb = self.q_target["floor"] * (q_spec_uf_hb / denom) if denom > 0 else 0.0
            self.U["uf_wall"] = qt_uf_wall / (H_vert * self.A_wall_uf_ex) if self.A_wall_uf_ex > 0 else 0.0
            self.psi_uf = qt_uf_hb / (H_vert * self.L_uf_ex) if self.L_uf_ex > 0 else 0.0
            self.U["floor"] = 0.0  # 基礎断熱では床(対外気床)は無し
        else:
            # 3.4.6.2.1 床断熱
            self.U["floor"] = self.q_target["floor"] / (H_floor * self.A_floor_ex) if self.A_floor_ex > 0 else 0.0
            self.U["uf_wall"] = 0.0
            self.psi_uf = 0.0

    # ---- 3.4.7 断熱材厚さ --------------------------------------------------
    def _s3_4_7(self):
        self.d_ins = {}

        def thick(part, U):
            ins = ref.insulation_layer(self.bt, part)
            if ins is None or U <= 0:
                return 0.0
            r_no = ref.r_noins(self.bt, part)
            return max((1.0 / U - r_no) * ins["lambda"], 0.0)

        self.d_ins["roof"] = thick("roof", self.U["roof"])
        self.d_ins["wall"] = thick("wall", self.U["wall"])
        if self.has_uf:
            self.d_ins["uf_wall"] = thick("uf_wall", self.U["uf_wall"])
        elif self.bt == "detached":
            self.d_ins["floor"] = thick("floor", self.U["floor"])

    # ---- 3.4.8〜3.4.10 平均日射熱取得率・窓η・補正 -------------------------
    def _s3_4_8_9_10(self):
        # 3.4.8 目標ηA
        self.eta_A_target = ((self.eta_AC * self.n_c + self.eta_AH * self.n_h)
                             / (self.n_c + self.n_h))
        nu = self.nu

        # 3.4.9 不透明部位の日射取得量 m_model_wall
        m = 0.0
        m += self.U["roof"] * self.A_roof_ex * nu["top"]
        # 外壁（居室）
        for s in self.spaces:
            if s == "UF":
                continue
            for d in VERT:
                m += self.U["wall"] * self.A_wall_ex[s][d] * nu[d]
        # 基礎壁（UF）
        if self.has_uf:
            for d in VERT:
                m += self.U["uf_wall"] * self.A_wall_ex["UF"][d] * nu[d]
        # ドア
        for s in self.spaces:
            for d in VERT:
                m += self.U["door"] * self.A_door[s][d] * nu[d]
        # 床下面は nu_bottom=0 のため寄与なし（省略）
        self.m_model_wall = M_OPAQUE_COEF * m

        # 窓のη暫定値
        denom = sum(self.A_win[s][d] * nu[d] for s in self.spaces for d in VERT)
        target_gain = self.eta_A_target * self.A_env_input / 100.0
        self.eta_win_temp = ((target_gain - self.m_model_wall) / denom) if denom > 0 else float("nan")

        # 3.4.10 クリップと窓面積・U値補正
        et = self.eta_win_temp
        if math.isnan(et):
            eta_win = ETA_WIN_MIN
            clipped = True
        elif et < ETA_WIN_MIN:
            eta_win = ETA_WIN_MIN
            clipped = True
        elif et > ETA_WIN_MAX:
            eta_win = ETA_WIN_MAX
            clipped = True
        else:
            eta_win = et
            clipped = False
        self.eta_win = eta_win
        self.eta_win_clipped = clipped

        # 窓面積・U値の補正
        self.A_win_mod = {s: {d: 0.0 for d in VERT} for s in self.spaces}
        self.U_win_mod = {s: {d: 0.0 for d in VERT} for s in self.spaces}
        for s in self.spaces:
            for d in VERT:
                A0 = self.A_win[s][d]
                if clipped and not math.isnan(et) and eta_win > 0:
                    A1 = max(A0 * (et / eta_win), 0.0)  # et<0 の高断熱域では0
                else:
                    A1 = A0
                self.A_win_mod[s][d] = A1
                self.U_win_mod[s][d] = (self.U["win"] * (A0 / A1)) if A1 > 0 else self.U["win"]
        self.glass_area_ratio = GLASS_AREA_RATIO
