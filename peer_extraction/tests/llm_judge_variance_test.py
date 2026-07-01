"""
llm_judge_variance_test.py
────────────────────────────────────────────────────────────
LlmJudge의 (model, prompt_version, issuer, candidate) 캐시를
의도적으로 우회(cache_path=None)하여, 동일 입력 쌍에 대해
N회 반복 LLM 호출 → tier/score 변동성을 측정.

대상: 센스톤 run.py 실행 결과의 최종 8개 피어
      (사업유사 단계 통과 후보 — tier 경계에 가까운 케이스가
       재현성에 가장 민감하므로 우선 검증 대상)

실행 위치: peer_extraction/  (core/, config.py, stage3_business_text_final.json 와 같은 레벨)
실행:
  python llm_judge_variance_test.py
  python llm_judge_variance_test.py --n 5     # 반복 횟수 조정 (기본 10)

출력:
  llm_judge_variance_results.json  (전체 raw 기록)
  콘솔 요약 테이블 (tier 일관성 %, score 분산, 통과/탈락 경계 위험 플래그)
────────────────────────────────────────────────────────────
"""
import sys, json, time, statistics, argparse
from pathlib import Path
from collections import Counter

# ── 프로젝트 모듈 경로 ──────────────────────────────────────
SCRIPT_DIR = Path(__file__).parent          # peer_extraction/tests/
ROOT_DIR   = SCRIPT_DIR.parent                # peer_extraction/
sys.path.insert(0, str(ROOT_DIR))

# ── .env 로드 (GEMINI_API_KEY 보장) ─────────────────────────
def _parse_val(raw: str) -> str:
    v = raw.strip()
    if v.startswith('"'):
        e = v.find('"', 1); return v[1:e].strip() if e != -1 else v[1:].strip()
    if v.startswith("'"):
        e = v.find("'", 1); return v[1:e].strip() if e != -1 else v[1:].strip()
    for sep in (' #', '\t#'):
        if sep in v: v = v[:v.index(sep)]
    return v.strip().strip('"').strip("'").strip()

def ensure_gemini_key():
    import os
    if os.environ.get("GEMINI_API_KEY") or os.environ.get("GOOGLE_API_KEY"):
        return
    for env_path in [ROOT_DIR / ".." / ".env", ROOT_DIR / ".env"]:
        env_path = env_path.resolve()
        if not env_path.exists(): continue
        with open(env_path, encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if not line or line.startswith("#") or "=" not in line: continue
                k, _, v = line.partition("=")
                k = k.strip()
                if k in ("GEMINI_API_KEY", "GOOGLE_API_KEY"):
                    os.environ[k] = _parse_val(v)
        if os.environ.get("GEMINI_API_KEY") or os.environ.get("GOOGLE_API_KEY"):
            print(f"[.env] GEMINI_API_KEY 로드 완료: {env_path}")
            return
    print("[경고] GEMINI_API_KEY를 찾지 못했습니다. genai.Client() 호출이 실패할 수 있습니다.")

ensure_gemini_key()

from core.llm_judge import LlmJudge   # noqa: E402

# ── 발행사 텍스트 (센스톤, 검증된 원문) ─────────────────────
ISSUER_TEXT = (
    "OTAC(일회용 인증코드) 원천기술 기반 인증보안 소프트웨어 기업. "
    "단방향 다이내믹 인증 코드로 사용자·기기·명령을 식별·인증. "
    "주요 제품: 출입통제·로그인·가상카드번호·금융거래 인증(SWIDCH ACCESS), "
    "IoT 및 디지털키 인증(SWIDCH CONNECT), FIDO·mOTP·OTAC 통합 인증 SDK. "
    "OT/PLC 산업제어 인증보안 게이트웨이, 금융 인증, 디지털 ID 분야에 적용."
)

# ── 이번 run.py 결과의 최종 8개 피어 (이름 기준) ────────────
# tier 경계(중간/낮음 50~65)에 가까운 케이스를 포함해 재현성 위험을 우선 검증
CANDIDATE_NAMES = [
    "한국전자인증", "라온시큐어", "유니온바이오메트릭스",
    "케이사인", "파수", "모니터랩",
    "지란지교시큐리티", "이니텍",
]

# ── stage3_business_text_final.json에서 회사명으로 business_text 조회 ──
def load_candidate_texts() -> dict:
    cache_path = ROOT_DIR / "data" / "cache" / "stage3_business_text_final.json"
    if not cache_path.exists():
        sys.exit(f"[ERROR] 캐시 없음: {cache_path}")
    raw = json.loads(cache_path.read_text(encoding="utf-8"))

    name_to_text = {}
    for sc, v in raw.items():
        nm = v.get("corp_name", "")
        if nm:
            name_to_text[nm] = v.get("business_text", "")

    result = {}
    missing = []
    for name in CANDIDATE_NAMES:
        if name in name_to_text and name_to_text[name].strip():
            result[name] = name_to_text[name]
        else:
            missing.append(name)
    if missing:
        print(f"[경고] business_text 미발견: {missing} — 이 회사들은 검증에서 제외됩니다.")
    return result

# ══════════════════════════════════════════════════════════
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--n", type=int, default=10, help="반복 호출 횟수 (기본 10)")
    args = ap.parse_args()
    N = args.n

    print("="*60)
    print(f"LlmJudge 재현성(변동성) 검증 — {N}회 반복, 캐시 우회")
    print("="*60)

    candidates = load_candidate_texts()
    print(f"\n검증 대상: {len(candidates)}개 후보 × {N}회 = "
          f"{len(candidates)*N}회 LLM 호출 예정\n")

    # {후보명: [{"tier":, "score":, "rationale":}, ...]}  (길이 N)
    records: dict[str, list] = {name: [] for name in candidates}

    total_raw_calls = 0
    for rep in range(1, N + 1):
        # ★ 캐시 완전 우회 — 매 반복마다 새 인스턴스, cache_path=None
        judge = LlmJudge(cache_path=None)
        print(f"[반복 {rep}/{N}]", end="  ")
        for name, ctext in candidates.items():
            result = judge.judge(ISSUER_TEXT, ctext, save=False)
            records[name].append(result)
            # _call_with_consistency가 다수결을 발동했으면 _consistency_votes가
            # 채워짐 — 이번 판정에 실제로 몇 번의 Gemini 호출이 들어갔는지 표시.
            n_calls = len(result.get("_consistency_votes", [result["tier"]]))
            total_raw_calls += n_calls
            tag = f"×{n_calls}" if n_calls > 1 else ""
            print(f"{name}={result['tier']}({result['score']}){tag}", end="  ")
            time.sleep(0.2)   # API 과부하 방지
        print()

    expected_calls = len(candidates) * N
    print(f"\n[호출 통계] 논리적 판정 {expected_calls}건 → "
          f"실제 Gemini 호출 {total_raw_calls}건 "
          f"(다수결로 인한 증가분 {total_raw_calls - expected_calls}건, "
          f"{(total_raw_calls/expected_calls - 1)*100:.0f}% 증가)")

    # ── 저장 — data/test/ 폴더로 통일 (tests/는 스크립트 위치,
    #          결과물은 data/test/에 분리 저장) ─────────────────
    out_dir = ROOT_DIR / "data" / "test"
    out_dir.mkdir(parents=True, exist_ok=True)
    out_path = out_dir / "llm_judge_variance_results.json"
    out_path.write_text(
        json.dumps(records, ensure_ascii=False, indent=2), encoding="utf-8")

    # ── 분석 ───────────────────────────────────────────────
    PASS_TIERS = {"높음", "중간"}   # stage3.py Stage3Config.keep_tiers와 일치 가정

    print(f"\n{'='*70}")
    print("재현성 분석 결과")
    print(f"{'='*70}")
    print(f"{'회사명':14s} {'tier 일관성':10s} {'주tier':8s} "
          f"{'score 평균':10s} {'score 표준편차':12s} {'경계위험':8s}")
    print("-"*70)

    risk_list = []
    for name, recs in records.items():
        tiers  = [r["tier"] for r in recs]
        scores = [r["score"] for r in recs]
        tier_counts = Counter(tiers)
        mode_tier, mode_count = tier_counts.most_common(1)[0]
        consistency = mode_count / len(recs) * 100

        score_mean = statistics.mean(scores)
        score_std  = statistics.stdev(scores) if len(scores) > 1 else 0.0

        # 통과/탈락 경계를 넘나드는지 (낮음 vs 높음·중간 혼재)
        pass_set = {t in PASS_TIERS for t in tiers}
        boundary_risk = len(pass_set) > 1   # True/False 둘 다 있으면 경계 넘나듦
        if boundary_risk:
            risk_list.append(name)

        flag = "⚠ 위험" if boundary_risk else ("🔶 변동" if consistency < 100 else "✅ 안정")
        print(f"{name:14s} {consistency:6.0f}%     {mode_tier:6s}   "
              f"{score_mean:7.1f}     {score_std:8.2f}       {flag}")

    print(f"\n{'='*70}")
    if risk_list:
        print(f"⚠ 통과/탈락 경계를 넘나든 회사 ({len(risk_list)}개): {risk_list}")
        print("  → 이 회사들은 실행 시점에 따라 최종 피어 목록에 포함되거나")
        print("    빠질 수 있음. Stage3Config.top_k_judge 확대 또는 임계값 재검토 고려.")
    else:
        print("✅ 모든 후보가 통과/탈락 경계를 넘나들지 않음 (tier 분류 안정적)")

    print(f"\n  상세 결과 저장: {out_path}")

if __name__ == "__main__":
    main()