"""파이프라인 통합테스트 (네트워크 불필요).

HashingEmbedder + MockJudge + 합성 유니버스/재무로 run_pipeline 전체 배선을 태운다.
stage1→2→3→4 컬럼 연결, track_type 분기, 최종 비교군 산출을 검증.
캐시 경로는 임시 디렉토리로 돌려 실제 코퍼스 캐시를 건드리지 않는다.
실행: peer_extraction 루트에서  python tests/test_pipeline.py
"""
from __future__ import annotations
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import numpy as np
import pandas as pd

import core.vector_cache as vcmod
from core.embedder import HashingEmbedder
from core.llm_judge import LlmJudge
from core.vector_cache import VectorCache
from config import Stage1Config, Stage3Config, Stage4Config
from pipeline import run_pipeline

results: list[tuple[str, bool, str]] = []


def check(name, cond, detail=""):
    results.append((name, bool(cond), detail))


class MockJudge(LlmJudge):
    def __init__(self, table):
        super().__init__(client="dummy", cache_path=None)
        self.table = table

    def _call(self, issuer, candidate):
        return self.table.get(candidate, {"tier": "낮음", "score": 10, "rationale": "기본"})


UNIVERSE = pd.DataFrame({
    "회사명": ["알파다이캐스팅", "베타캐스팅", "감마부품", "델타소재", "적자사", "비12월사"],
    "종목코드": ["000001", "000002", "000003", "000004", "000005", "000006"],
    "시장구분": ["KOSDAQ"] * 6,
    "산업분류코드": ["C00303", "C00303", "C00303", "C00281", "C00303", "C00303"],
    "주요제품": ["알루미늄 다이캐스팅 자동차부품", "마그네슘 다이캐스팅 경량 부품",
                "자동차 엔진 부품", "알루미늄 압출 건축자재", "다이캐스팅 부품", "다이캐스팅"],
    "업종": ["자동차부품", "자동차부품", "자동차부품", "1차금속", "자동차부품", "자동차부품"],
    "결산월": ["12월", "12월", "12월", "12월", "12월", "3월"],
})

FIN = pd.DataFrame({
    "종목코드": ["000001", "000002", "000003", "000004", "000005", "000006"],
    "결산월": ["12월", "12월", "12월", "12월", "12월", "3월"],
    "op_income_annual": [100, 110, 120, 60, -10, 100],
    "op_income_ltm": [110, 120, 130, 65, -5, 110],
    "net_income_annual": [70, 80, 90, 50, -20, 70],
    "net_income_ltm": [80, 90, 100, 50, -15, 80],
    "audit_opinion": ["적정의견"] * 6,
    "has_financials": [True] * 6,
    "market_cap": [1000, 1200, 1100, 900, 800, 700],
})

JUDGE_TABLE = {
    "알루미늄 다이캐스팅 자동차부품 | 자동차부품": {"tier": "높음", "score": 90, "rationale": "직접경쟁"},
    "마그네슘 다이캐스팅 경량 부품 | 자동차부품": {"tier": "높음", "score": 88, "rationale": "동일사업"},
    "자동차 엔진 부품 | 자동차부품": {"tier": "중간", "score": 65, "rationale": "부분중첩"},
    "알루미늄 압출 건축자재 | 1차금속": {"tier": "낮음", "score": 20, "rationale": "소재/건설"},
}

ISSUER_TEXT = "알루미늄 마그네슘 다이캐스팅 자동차 경량화 부품"


def run():
    tmp = Path(tempfile.mkdtemp())
    # 캐시 경로를 임시로 (실제 코퍼스 캐시 보호)
    vcmod.CORPUS_EMB_NPY = tmp / "emb.npy"
    vcmod.CORPUS_META_JSON = tmp / "meta.json"
    csv = tmp / "universe.csv"
    UNIVERSE.to_csv(csv, index=False, encoding="utf-8-sig")

    vc = VectorCache(HashingEmbedder(), csv_path=csv)
    judge = MockJudge(JUDGE_TABLE)

    res = run_pipeline(
        "테스트발행사", "C00303", ISSUER_TEXT, "general",
        vcache=vc, judge=judge, financials=FIN,
        s1cfg=Stage1Config(top_k_embed=6),
        s3cfg=Stage3Config(top_k_judge=10),
        s4cfg=Stage4Config(),
    )

    c = res.counts()
    check("모집단 ≥ 4", c["모집단"] >= 4, str(c))
    check("재무통과: 적자사·비12월사 제외", c["재무통과"] >= 3 and c["재무통과"] <= 4, str(c))

    final_names = set(res.final["회사명"])
    check("최종에 다이캐스팅 2사 포함", {"알파다이캐스팅", "베타캐스팅"} <= final_names, str(final_names))
    check("델타소재(낮음) 최종 제외", "델타소재" not in final_names, str(final_names))
    check("적자사 최종 제외", "적자사" not in final_names)
    check("비12월사 최종 제외", "비12월사" not in final_names)

    # PER 부착·요약
    check("per 컬럼 존재·양수", (res.final["per"] > 0).all() if len(res.final) else False)
    check("PER 요약 n>0", res.per_summary.get("n", 0) > 0, str(res.per_summary))

    # provenance 체인 존재
    check("stage2 provenance", "drop_reason" in res.stage2)
    check("stage3 tier", "tier" in res.stage3)
    check("stage4 per", "per" in res.stage4)

    import shutil
    shutil.rmtree(tmp, ignore_errors=True)


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
