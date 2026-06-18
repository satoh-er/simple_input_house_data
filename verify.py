# -*- coding: utf-8 -*-
"""
verify.py
==============================================================================
構築した暖冷房負荷モデルから UA値・ηA値・各種面積を逆算し、設計住戸の入力値を
再現できているかを検証する。

検証方針（最終的に出力されるモデル量から逆算する end-to-end チェック）:
  - 床面積  : 空間床面積の合計 = 入力 A_A、MR/OR は入力値そのもの
  - 総外皮面積 A_env : モデルの全外皮（上面+下面+垂直, UF含む）合計 = 入力 A_env
       ※ 窓クリップ補正(3.4.10)が働いた場合は再現されない（仕様明記）
  - UA値    : Σ(U·A·H) + ψ·L·H を A_env で除す = 入力 UA
       ※ 窓補正は U·A を保存するため UA は再現される
  - ηA値    : (不透明部位日射取得 + 窓日射取得) / A_env ×100 = 目標ηA
       ※ 窓クリップ（特に暫定η<0）時は再現されない
==============================================================================
"""
import math
import reference_data as ref

VERT = ref.VERT_DIRS


def verify(m, tol=0.5):
    """m: HeatLoadModel。許容誤差 tol[%] 内かを各項目で判定し dict を返す。"""
    H_top = ref.temp_diff("top")
    H_vert = ref.temp_diff("vert")
    H_floor = ref.temp_diff("floor_ne_uf_btm")

    # --- 床面積 ---
    A_A_model = sum(m.A[s] for s in ref.SPACES)

    # --- 総外皮面積（最終, 窓補正後）---
    A_env_model = 0.0
    for s in m.spaces:
        A_env_model += m.A_top[s] + m.A_btm[s] + m._vert_total_model(s)
    # 窓補正により窓面積が変化した分を反映（外壁等は窓実面積で頭打ち済みだが、
    # A_env 自体は top/btm/vert の幾何面積合計なので窓補正では変わらない。
    # 窓補正で「窓面積>外皮」が起こり得る分のみ別途確認する）
    win_orig = sum(m.A_win[s][d] for s in m.spaces for d in VERT)
    win_mod = sum(m.A_win_mod[s][d] for s in m.spaces for d in VERT)
    A_env_model_eff = A_env_model + (win_mod - win_orig)

    # --- UA値（最終モデルの U·A·H を集計）---
    q = 0.0
    q += m.U["roof"] * m.A_roof_ex * H_top
    q += m.U["wall"] * m.A_wall_neuf_ex * H_vert
    # 窓: U·A は補正で保存されるので元の U_win×A_win_ex を用いる
    q += m.U["win"] * m.A_win_ex * H_vert
    q += m.U["door"] * m.A_door_ex * H_vert
    if m.has_uf:
        q += m.U["uf_wall"] * m.A_wall_uf_ex * H_vert
        q += m.psi_uf * m.L_uf_ex * H_vert      # 土間床外周（エンジン実装待ち項）
    else:
        q += m.U["floor"] * m.A_floor_ex * H_floor
    UA_model = q / m.A_env_input if m.A_env_input > 0 else 0.0

    # --- ηA値（不透明 + 窓, 最終）---
    nu = m.nu
    gain_win = sum(m.eta_win * m.A_win_mod[s][d] * nu[d] for s in m.spaces for d in VERT)
    gain_total = m.m_model_wall + gain_win
    etaA_model = gain_total / m.A_env_input * 100.0 if m.A_env_input > 0 else 0.0

    def pct_err(model, target):
        if target == 0:
            return 0.0 if abs(model) < 1e-9 else float("inf")
        return (model - target) / target * 100.0

    results = {
        "A_A":   _row("床面積合計 A_A [m2]", A_A_model, m.A_A, tol),
        "A_MR":  _row("主たる居室 A_MR [m2]", m.A["MR"], m.A["MR"], tol),
        "A_OR":  _row("その他居室 A_OR [m2]", m.A["OR"], m.A["OR"], tol),
        "A_env": _row("総外皮面積 A_env [m2]", A_env_model_eff, m.A_env_input, tol),
        "UA":    _row("外皮平均熱貫流率 UA [W/m2K]", UA_model, m.UA, tol),
        "etaA":  _row("平均日射熱取得率 ηA [-]", etaA_model, m.eta_A_target, tol),
    }
    results["_window_clipped"] = m.eta_win_clipped
    results["_eta_win_temp"] = m.eta_win_temp
    results["_eta_win"] = m.eta_win
    return results


def _row(label, model, target, tol):
    if target == 0:
        err = 0.0 if abs(model) < 1e-9 else float("inf")
    else:
        err = (model - target) / target * 100.0
    return {
        "label": label, "model": model, "target": target,
        "err_pct": err, "ok": abs(err) <= tol,
    }


def format_report(results):
    lines = []
    lines.append(f"{'項目':<28}{'目標(入力)':>14}{'モデル':>14}{'誤差%':>10}  判定")
    lines.append("-" * 80)
    for key in ("A_A", "A_MR", "A_OR", "A_env", "UA", "etaA"):
        r = results[key]
        mark = "OK" if r["ok"] else "NG"
        lines.append(f"{r['label']:<28}{r['target']:>14.3f}{r['model']:>14.3f}"
                     f"{r['err_pct']:>10.3f}  {mark}")
    if results["_window_clipped"]:
        lines.append("-" * 80)
        lines.append(f"※ 窓η暫定値={results['_eta_win_temp']:.4f} が上下限を外れ "
                     f"η={results['_eta_win']:.2f} にクリップ。"
                     f"この場合 A_env は再現されない（仕様3.4.10）。")
    return "\n".join(lines)
