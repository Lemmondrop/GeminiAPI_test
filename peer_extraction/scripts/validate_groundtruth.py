"""Ground-truth 검증 — 신고서 실제 피어 vs 파이프라인 출력.

빠른 모드(기본, LLM 없음): Stage 1(모집단) + Stage 2(재무)까지 돌려
  정답 피어가 각 단계를 '생존'하는 비율 측정 → 어느 단계에서 정답을 잃는지 진단.
정밀 모드(--full): Stage 3~4까지(LLM) 돌려 최종 precision/recall (비용↑, --limit 권장).

입력: data/labels/{peer_labels,issuer_ksic}.csv + filing_labels_cache.json(사업개요용)
실행:
  python scripts\\validate_groundtruth.py --limit 15           # 파일럿(빠른 모드)
  python scripts\\validate_groundtruth.py                       # 전체(빠른 모드)
  python scripts\\validate_groundtruth.py --full --limit 10     # 정밀(LLM)
"""
from __future__ import annotations
import argparse
import json
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import pandas as pd

import config
from core.cooccur import _norm_name, load_prior
from stages.stage1_population import StartupQuery, build_population
from stages.stage2_financial import apply_stage2, load_financials, survivors as s2_surv

_OVERVIEW_ANCHORS = ("사업의 개요", "회사의 개요", "사업의 내용", "주요 제품 및 서비스",
                     "주요 제품", "회사의 현황")


def business_overview(corp_code: str, rcept: str) -> str:
    """캐시된 신고서 본문에서 사업개요 ~2500자 추출(Lane B 임베딩 입력용)."""
    if not rcept:
        return ""
    fp = config.DART_DOC_CACHE_DIR / f"{rcept}.txt"
    if not fp.exists():
        return ""
    text = fp.read_text(encoding="utf-8")
    for kw in _OVERVIEW_ANCHORS:
        i = text.find(kw)
        if i >= 0:
            return re.sub(r"\s+", " ", text[i:i + 2500])
    return ""


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--labels-dir", default="data/labels")
    ap.add_argument("--limit", type=int, default=None)
    ap.add_argument("--full", action="store_true", help="Stage 3~4(LLM)까지 — 비용↑")
    ap.add_argument("--no-hardcut", action="store_true",
                    help="Stage4 하드컷(시총<1000억/PER 10~100) 끔 — 검증용(정답은 중·대형 피어)")
    ap.add_argument("--min-avail", type=int, default=2, help="검증 최소 정답 수")
    ap.add_argument("--exclude-loss-makers", action="store_true",
                    help="현재 적자(net_income_annual<=0) 정답 피어를 분모에서 제외(시점 불일치 보정)")
    ap.add_argument("--verbose", action="store_true", help="발행사별 상세")
    a = ap.parse_args()

    ldir = Path(a.labels_dir)
    peers = pd.read_csv(ldir / "peer_labels.csv", encoding="utf-8-sig", dtype=str)
    iksic = pd.read_csv(ldir / "issuer_ksic.csv", encoding="utf-8-sig", dtype=str)
    corp_seed = {r["corp_code"]: (str(r.get("induty_code") or "").strip()
                                  or str(r.get("issuer_sub") or "").strip())
                 for _, r in iksic.iterrows()}
    cache = (json.loads(config.FILING_LABELS_CACHE.read_text(encoding="utf-8"))
             if config.FILING_LABELS_CACHE.exists() else {})
    corp_rcept = {c: r.get("initial_rcept", "") for c, r in cache.items()}

    uni = pd.read_csv(config.LISTED_UNIVERSE_CSV, encoding=config.CSV_ENCODING, dtype=str)
    uni_norm = {_norm_name(n) for n in uni[config.NAME_COL].dropna()}

    # 현재 흑자 피어 집합(시점 불일치 보정용): net_income_annual>0
    profitable = set()
    if a.exclude_loss_makers:
        finu = pd.read_csv(config.UNIVERSE_FIN_CSV, dtype=str)
        code2na = {}
        for _, r in finu.iterrows():
            try:
                code2na[str(r.get("종목코드", "")).zfill(6)] = float(r.get("net_income_annual"))
            except (ValueError, TypeError):
                pass
        for _, r in uni.iterrows():
            code = str(r.get(config.STOCK_COL, "")).zfill(6)
            na = code2na.get(code)
            if na is not None and na > 0:
                profitable.add(_norm_name(r.get(config.NAME_COL)))
        print(f"현재 흑자 피어 집합: {len(profitable):,}개 (적자/결측 제외)\n")

    # 발행사별 정답(모집단 존재분; 옵션 시 현재 흑자분만)
    peers = peers[peers["peer"].astype(str).str.strip().ne("")]
    gt, excluded_lm = {}, 0
    for issuer, g in peers.groupby("issuer"):
        corp = g["corp_code"].iloc[0]
        avail = {_norm_name(p) for p in g["peer"] if _norm_name(p) in uni_norm}
        if a.exclude_loss_makers:
            before = len(avail)
            avail = {p for p in avail if p in profitable}
            excluded_lm += before - len(avail)
        if len(avail) >= a.min_avail and corp in corp_seed:
            gt[issuer] = {"corp": corp, "avail": avail}
    issuers = list(gt.items())
    if a.limit:
        issuers = issuers[:a.limit]
    note = " · 현재적자 제외" if a.exclude_loss_makers else ""
    print(f"검증 발행사: {len(issuers)}개 (정답 {a.min_avail}개+, 모집단 존재분{note})")
    if a.exclude_loss_makers:
        print(f"  시점불일치(현재적자) 제외 정답: {excluded_lm}개")
    print()

    # 의존성 1회 구성
    from core.embedder import GeminiEmbedder
    from core.vector_cache import VectorCache
    vcache = VectorCache(GeminiEmbedder())
    vcache.build_or_load()
    financials = load_financials()
    prior = load_prior() or None
    judge = None
    if a.full:
        from core.llm_judge import LlmJudge
        judge = LlmJudge()
        from stages.stage3_business import apply_stage3, final_peers as s3_final
        from stages.stage4_general import apply_stage4, final_comps

    rows = []
    for k, (issuer, info) in enumerate(issuers, 1):
        try:
            corp, avail = info["corp"], info["avail"]
            seed = corp_seed[corp]
            btext = business_overview(corp, corp_rcept.get(corp, ""))
            q = StartupQuery(seed_code=seed, business_text=btext)
            pool = build_population(q, vcache, cooccur_prior=prior, cfg=config.Stage1Config())
            pool = pool[pool[config.NAME_COL].map(_norm_name) != _norm_name(issuer)]
            pool_n = {_norm_name(n) for n in pool[config.NAME_COL]}
            in_pool = avail & pool_n

            s2 = apply_stage2(pool, financials, "general")
            s2_n = {_norm_name(n) for n in s2_surv(s2)[config.NAME_COL]}
            in_s2 = avail & s2_n

            rec = {"issuer": issuer, "n_avail": len(avail),
                   "in_pool": len(in_pool), "in_s2": len(in_s2),
                   "recall_pool": len(in_pool) / len(avail),
                   "recall_s2": len(in_s2) / len(avail)}

            if a.full:
                s4cfg = config.Stage4Config()
                if a.no_hardcut:                          # 검증: 스타트업 규모 필터 해제
                    s4cfg = config.Stage4Config(mcap_max=None, per_min=None, per_max=None)
                s3 = apply_stage3(s2_surv(s2), btext, judge, config.Stage3Config())
                s4 = apply_stage4(s3_final(s3), s4cfg)
                fin = final_comps(s4)
                fin_n = {_norm_name(n) for n in fin[config.NAME_COL]}
                tp = avail & fin_n
                rec["n_pred"] = len(fin_n)
                rec["precision"] = len(tp) / len(fin_n) if fin_n else 0.0
                rec["recall_final"] = len(tp) / len(avail)
        except Exception as e:
            print(f"  [{k}/{len(issuers)}] {issuer[:16]:<16} 실패: {type(e).__name__} — 건너뜀")
            pd.DataFrame(rows).to_csv(
                ldir / ("validation_full.csv" if a.full else "validation_fast.csv"),
                index=False, encoding="utf-8-sig")          # 부분 결과 보존
            continue
        rows.append(rec)
        if a.verbose or k % 10 == 0 or k == len(issuers):
            extra = (f" P {rec.get('precision',0):.2f} Rf {rec.get('recall_final',0):.2f}"
                     if a.full else "")
            print(f"  [{k}/{len(issuers)}] {issuer[:16]:<16} "
                  f"정답 {rec['n_avail']} → 모집단 {rec['in_pool']} → 재무 {rec['in_s2']}"
                  f"  (R@pool {rec['recall_pool']:.2f}/R@s2 {rec['recall_s2']:.2f}){extra}")

    df = pd.DataFrame(rows)
    if df.empty:
        print("\n[결과 없음] 모든 발행사 실패 — Gemini 503/네트워크 확인 후 재실행(캐시 이어감).")
        return
    out = ldir / ("validation_full.csv" if a.full else "validation_fast.csv")
    df.to_csv(out, index=False, encoding="utf-8-sig")

    print("\n" + "=" * 55)
    print(f"검증 발행사 {len(df)}개 / 정답 평균 {df['n_avail'].mean():.1f}개")
    print(f"Stage1 모집단 recall: {df['recall_pool'].mean():.3f}  "
          f"(정답의 {df['recall_pool'].mean()*100:.0f}%가 모집단 진입)")
    print(f"Stage2 재무 후 recall: {df['recall_s2'].mean():.3f}  "
          f"(모집단→재무에서 {(df['recall_pool'].mean()-df['recall_s2'].mean())*100:.0f}%p 손실)")
    if a.full:
        print(f"최종 precision: {df['precision'].mean():.3f} / "
              f"recall: {df['recall_final'].mean():.3f}")
    # 모집단에서 아예 못 잡은 발행사(Stage1 recall 0) 진단
    zero = df[df["recall_pool"] == 0]
    print(f"\n모집단 recall 0 발행사: {len(zero)}개 (Stage1이 정답을 전혀 못 잡음)")
    if len(zero):
        print("  " + ", ".join(zero["issuer"].head(10)))
    print(f"저장: {out}")


if __name__ == "__main__":
    main()