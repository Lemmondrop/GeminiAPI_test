"""Stage 4 v3 단위테스트 (네트워크 불필요).

하드 컷(시총≥1000억 / PER<10 / PER≥100) · PER계산 · 적자/시총결측 ·
유사도 밴드 max_peers · 통계방식(iqr/trim, 하드컷 비활성 시) · 중앙값 헤드라인 · 피어부족.
실행: peer_extraction 루트에서  python tests\\test_stage4.py
"""
from __future__ import annotations
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import numpy as np
import pandas as pd

from config import Stage4Config
from stages.stage4_general import apply_stage4, final_comps, peer_per_summary

results: list[tuple[str, bool, str]] = []
NO_HARD = dict(mcap_max=None, per_min=None, per_max=None)   # 하드컷 끄기(통계방식 격리 테스트)


def check(name, cond, detail=""):
    results.append((name, bool(cond), detail))


def _peers(rows):
    return pd.DataFrame(rows)


def run():
    # ===== 하드 컷 (기본 cfg) =====
    # 억 단위를 원으로: 500억=5e10, 1000억=1e11, 1500억=1.5e11
    hard = _peers([
        {"회사명": "적정A", "tier": "높음", "sim_score": 90, "market_cap": 5e10, "net_income_ltm": 2.5e9},   # 시총500억 PER20 ✓
        {"회사명": "대형사", "tier": "높음", "sim_score": 88, "market_cap": 1.5e11, "net_income_ltm": 5e9},   # 시총1500억 ✗
        {"회사명": "저PER사", "tier": "높음", "sim_score": 86, "market_cap": 4e10, "net_income_ltm": 5e9},    # PER8 ✗
        {"회사명": "고PER사", "tier": "중간", "sim_score": 70, "market_cap": 8e10, "net_income_ltm": 5e8},    # PER160 ✗
        {"회사명": "적정B", "tier": "중간", "sim_score": 68, "market_cap": 6e10, "net_income_ltm": 4e9},      # 시총600억 PER15 ✓
        {"회사명": "적정C", "tier": "중간", "sim_score": 66, "market_cap": 9e10, "net_income_ltm": 3e9},      # 시총900억 PER30 ✓
    ])
    s4 = apply_stage4(hard, Stage4Config())
    fc = set(final_comps(s4)["회사명"])
    rmap = dict(zip(s4["회사명"], s4["drop_reason"]))

    check("하드컷 통과 = 적정 3사", fc == {"적정A", "적정B", "적정C"}, str(fc))
    check("시총≥1000억 제외", "적정" not in rmap.get("대형사", "") and "시총과대" in rmap["대형사"], rmap["대형사"])
    check("PER<10 제외", "PER과소" in rmap["저PER사"], rmap["저PER사"])
    check("PER≥100 제외", "PER과대" in rmap["고PER사"], rmap["고PER사"])

    summ = peer_per_summary(s4, Stage4Config())
    check("대표 멀티플 중앙값=20", summ["per_median"] == 20.0, str(summ))   # PER 20,15,30 → median 20
    check("피어 충분(thin=False)", summ["thin"] is False)

    # ===== 적자/시총결측 =====
    edge = _peers([
        {"회사명": "흑자", "tier": "높음", "sim_score": 90, "market_cap": 5e10, "net_income_ltm": 2.5e9},   # PER20 ✓
        {"회사명": "적자", "tier": "중간", "sim_score": 70, "market_cap": 5e10, "net_income_ltm": -1e9},    # PER<0 ✗
        {"회사명": "시총없음", "tier": "중간", "sim_score": 65, "market_cap": np.nan, "net_income_ltm": 2e9},  # ✗
        {"회사명": "적정D", "tier": "중간", "sim_score": 60, "market_cap": 6e10, "net_income_ltm": 4e9},     # PER15 ✓
        {"회사명": "적정E", "tier": "중간", "sim_score": 55, "market_cap": 7e10, "net_income_ltm": 3e9},     # PER23 ✓
    ])
    s4e = apply_stage4(edge, Stage4Config())
    fce = set(final_comps(s4e)["회사명"])
    re = dict(zip(s4e["회사명"], s4e["drop_reason"]))
    check("적자 제외", "적자" not in fce and "PER" in re["적자"], re["적자"])
    check("시총결측 제외", "시총없음" not in fce and "시총없음" in re["시총없음"], re["시총없음"])
    check("적정 3사 통과", fce == {"흑자", "적정D", "적정E"}, str(fce))

    # ===== 피어 부족 경고 =====
    thin = _peers([
        {"회사명": "유일", "tier": "높음", "sim_score": 90, "market_cap": 5e10, "net_income_ltm": 2.5e9},   # PER20 ✓
        {"회사명": "대형", "tier": "높음", "sim_score": 85, "market_cap": 2e11, "net_income_ltm": 5e9},      # 시총과대 ✗
    ])
    s4t = apply_stage4(thin, Stage4Config())
    st = peer_per_summary(s4t, Stage4Config())
    check("피어 부족 thin=True", st.get("thin") is True, str(st))

    # ===== 통계방식 격리 (하드컷 OFF) =====
    band = _peers([
        {"회사명": "T1", "tier": "높음", "sim_score": 90, "market_cap": 500, "net_income_ltm": 100},   # PER5
        {"회사명": "T2", "tier": "높음", "sim_score": 85, "market_cap": 1000, "net_income_ltm": 100},  # PER10
        {"회사명": "T3", "tier": "중간", "sim_score": 70, "market_cap": 1100, "net_income_ltm": 100},  # PER11
        {"회사명": "T4", "tier": "중간", "sim_score": 65, "market_cap": 1200, "net_income_ltm": 100},  # PER12
        {"회사명": "T5", "tier": "중간", "sim_score": 60, "market_cap": 3000, "net_income_ltm": 100},  # PER30
    ])
    s4tr = apply_stage4(band, Stage4Config(outlier_method="trim", trim_each_side=1, **NO_HARD))
    ftr = set(final_comps(s4tr)["회사명"])
    check("trim: 최저T1·최고T5 절사", ftr == {"T2", "T3", "T4"}, str(ftr))

    s4iq = apply_stage4(band, Stage4Config(outlier_method="iqr", **NO_HARD))
    check("iqr: PER30 이상치 제거", "T5" not in set(final_comps(s4iq)["회사명"]),
          str(set(final_comps(s4iq)["회사명"])))

    s4no = apply_stage4(band, Stage4Config(outlier_method="none", **NO_HARD))
    check("none: 5개 전부 유지", len(final_comps(s4no)) == 5)

    # max_peers 컷 (하드컷 OFF)
    s4m = apply_stage4(band, Stage4Config(max_peers=3, **NO_HARD))
    check("max_peers=3 → 상위 3(T1,T2,T3)", set(final_comps(s4m)["회사명"]) == {"T1", "T2", "T3"},
          str(set(final_comps(s4m)["회사명"])))

    # headline median/mean
    sm = peer_per_summary(s4no, Stage4Config(headline="median", **NO_HARD))
    check("headline=median", sm["headline_per"] == sm["per_median"])
    sm2 = peer_per_summary(s4no, Stage4Config(headline="mean", **NO_HARD))
    check("headline=mean", sm2["headline_per"] == sm2["per_mean"])

    # 시총 전부 결측 → 보류
    nomcap = band.copy()
    nomcap["market_cap"] = np.nan
    s4d = apply_stage4(nomcap, Stage4Config())
    check("시총 전부 결측 → 통과 0", len(final_comps(s4d)) == 0)
    check("시총 전부 결측 → 보류 사유", "KRX" in s4d["drop_reason"].iloc[0], s4d["drop_reason"].iloc[0])

    # ===== ② 고유사도 PER 컷 면제 (high_sim_exempt) =====
    # 높음 유사도지만 PER이 상·하한 밖인 피어. 면제 OFF면 컷, ON이면 생존.
    hs = _peers([
        {"회사명": "핵심高PER", "tier": "높음", "sim_score": 92, "market_cap": 8e10, "net_income_ltm": 1e9},   # PER80 (>50 가정)
        {"회사명": "핵심低PER", "tier": "높음", "sim_score": 90, "market_cap": 5e10, "net_income_ltm": 6.25e9}, # PER8 (<10)
        {"회사명": "보통",      "tier": "중간", "sim_score": 60, "market_cap": 6e10, "net_income_ltm": 3e9},     # PER20 ✓
    ])
    cfg_cut = Stage4Config(per_min=10, per_max=50, mcap_max=None, high_sim_exempt=False)
    cfg_keep = Stage4Config(per_min=10, per_max=50, mcap_max=None, high_sim_exempt=True)
    off = apply_stage4(hs, cfg_cut); off_re = dict(zip(off["회사명"], off["drop_reason"]))
    on = apply_stage4(hs, cfg_keep); on_fc = set(final_comps(on)["회사명"])

    check("면제OFF: 高PER 컷됨", "PER과대" in off_re["핵심高PER"], off_re["핵심高PER"])
    check("면제OFF: 低PER 컷됨", "PER과소" in off_re["핵심低PER"], off_re["핵심低PER"])
    check("면제ON: 高PER 생존", "핵심高PER" in on_fc, str(on_fc))
    check("면제ON: 低PER 생존", "핵심低PER" in on_fc, str(on_fc))
    check("면제ON: 3사 모두 생존", on_fc == {"핵심高PER", "핵심低PER", "보통"}, str(on_fc))

    # 면제는 PER 컷만 — PER≤0(적자)·시총과대는 여전히 제외
    hs2 = _peers([
        {"회사명": "高유사적자", "tier": "높음", "sim_score": 95, "market_cap": 5e10, "net_income_ltm": -1e9},  # 적자
        {"회사명": "高유사대형", "tier": "높음", "sim_score": 95, "market_cap": 2e11, "net_income_ltm": 5e9},   # 시총2000억
        {"회사명": "정상",      "tier": "중간", "sim_score": 60, "market_cap": 6e10, "net_income_ltm": 3e9},
    ])
    on2 = apply_stage4(hs2, Stage4Config(per_min=10, per_max=50, mcap_max=1000e8, high_sim_exempt=True))
    re2 = dict(zip(on2["회사명"], on2["drop_reason"]))
    check("면제ON이라도 적자는 제외", "PER≤0" in re2["高유사적자"], re2["高유사적자"])
    check("면제ON이라도 시총과대는 제외", "시총과대" in re2["高유사대형"], re2["高유사대형"])

    # ===== sim_weighted 대표값 =====
    sw = _peers([
        {"회사명": "A", "tier": "높음", "sim_score": 90, "market_cap": 8.4e10, "net_income_ltm": 1e9},   # PER84
        {"회사명": "B", "tier": "높음", "sim_score": 90, "market_cap": 1.16e10, "net_income_ltm": 1e9},  # PER11.6
        {"회사명": "C", "tier": "중간", "sim_score": 55, "market_cap": 1.29e10, "net_income_ltm": 1e9},  # PER12.9
    ])
    s4sw = apply_stage4(sw, Stage4Config(per_min=None, per_max=None, mcap_max=None))
    sm_med = peer_per_summary(s4sw, Stage4Config(headline="median"))
    sm_sw = peer_per_summary(s4sw, Stage4Config(headline="sim_weighted"))
    check("sim_weighted 키 존재", "per_sim_weighted" in sm_med, str(sm_med.keys()))
    check("headline=sim_weighted 적용", sm_sw["headline_per"] == sm_sw["per_sim_weighted"], str(sm_sw))
    # 가중평균은 median(12.9)보다 큼(高PER 피어가 高유사도라 비중↑)
    check("sim_weighted > median", sm_sw["per_sim_weighted"] > sm_med["per_median"],
          f"{sm_sw['per_sim_weighted']} vs {sm_med['per_median']}")


def main():
    run()
    passed = sum(1 for _, ok, _ in results if ok)
    print("=" * 60)
    for name, ok, detail in results:
        mark = "PASS" if ok else "FAIL"
        line = f"[{mark}] {name}"
        if not ok and detail:
            line += f"   <- {detail}"
        print(line)
    print("=" * 60)
    print(f"{passed}/{len(results)} passed")
    sys.exit(0 if passed == len(results) else 1)


if __name__ == "__main__":
    main()