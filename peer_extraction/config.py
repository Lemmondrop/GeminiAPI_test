"""peer_extraction 전역 설정.

경로 · 임베딩 모델/차원 · Stage1 파라미터 · track_type별 조건세트를 한곳에 모은다.
분기 로직을 코드 곳곳에 흩뿌리지 않기 위해, 트랙별 규칙은 TRACK_RULES 하나로 통합한다.

※ 실행은 반드시 peer_extraction 루트에서. (예: `python run.py` / cwd = peer_extraction)
"""
from __future__ import annotations
from dataclasses import dataclass
from pathlib import Path

# ----------------------------- 경로 -----------------------------
ROOT = Path(__file__).resolve().parent


# ----------------------------- .env 로딩 -----------------------------
# GEMINI_API_KEY 등은 상위 'OCR Sample' 디렉토리의 .env에 선언됨 (peer_extraction의 부모).
# import 시 1회 os.environ에 주입 → genai.Client()가 자동으로 키를 읽는다.
def _load_env() -> None:
    try:
        from dotenv import find_dotenv, load_dotenv
    except ImportError:
        return  # python-dotenv 미설치 시 조용히 패스 (오프라인 테스트는 키 불필요)
    parent_env = ROOT.parent / ".env"          # .../OCR Sample/.env
    if parent_env.exists():
        load_dotenv(parent_env)
        return
    found = find_dotenv(usecwd=True)           # 폴백: 상위 디렉토리로 탐색
    if found:
        load_dotenv(found)


_load_env()

# 소스 CSV(상장사+KSIC)는 peer_extraction 외부의 공유 디렉토리에 위치 (다른 모듈과 공유).
#   .../OCR Sample/data/metadata/company_cord_prototype.csv
SHARED_DATA_DIR = ROOT.parent / "data"
LISTED_UNIVERSE_CSV = SHARED_DATA_DIR / "metadata" / "company_cord_prototype.csv"

# 빌드 산출물(임베딩 캐시)은 peer_extraction 내부에 보관.
DATA_DIR = ROOT / "data"
CACHE_DIR = DATA_DIR / "cache"
CORPUS_EMB_NPY = CACHE_DIR / "corpus_embeddings.npy"      # 코퍼스 1회 임베딩 캐시
CORPUS_META_JSON = CACHE_DIR / "corpus_meta.json"         # 캐시 무결성 메타
CACHE_DIR.mkdir(parents=True, exist_ok=True)

# CSV 인코딩 (한글 CSV가 cp949인 경우가 많음 → 실패 시 utf-8-sig 폴백은 vector_cache에서 처리)
CSV_ENCODING = "cp949"

# ----------------------------- 임베딩 -----------------------------
EMBED_MODEL = "gemini-embedding-001"   # 버전 핀. floating alias 금지 (재현성 잠금의 핵심)
OUTPUT_DIM = 1024                       # MRL 축소 차원. 캐시/FAISS 전부 이 값으로 고정.
EMBED_BATCH = 100                       # API 1회 요청당 텍스트 수 (대용량 코퍼스 청크 처리)
EMBED_MAX_PER_MIN = 2500                # 분당 임베딩 항목 상한 (paid 3000/min 쿼터에 안전마진)
EMBED_MAX_RETRIES = 6                    # 429(RESOURCE_EXHAUSTED) 자동 재시도 횟수

# 코퍼스 임베딩 텍스트 컬럼.
#   '업종'은 KSIC 분류 라벨을 그대로 담은 값이라, 포함하면 같은 코드끼리 인위적으로 뭉쳐
#   Lane B가 Lane A(코드)의 복제로 붕괴된다(게이트 1-B 1차 실험에서 실증). → '주요제품' 단독 사용.
#   (이상적으로는 DART '사업의 내용' 매출구성으로 보강 — 더 풍부한 사업 설명)
CORPUS_TEXT_COLS = ("주요제품",)
# 위 주 텍스트가 빈 행(주요제품 결측 등)만 아래 컬럼으로 코얼레스 폴백 (빈 Part API 오류 방지).
# 주의: 전체 행에 업종을 더하는 게 아니라, 빈 행에 한해서만 사용한다.
CORPUS_TEXT_FALLBACK_COLS = ("업종", "회사명")
NAME_COL = "회사명"
CODE_COL = "산업분류코드"
MARKET_COL = "시장구분"
STOCK_COL = "종목코드"            # DART corp_code 매핑·pykrx 시총 조인 키
FISCAL_MONTH_COL = "결산월"       # Stage 2 12월 결산 필터 (CSV 보유)

# ----------------------------- DART 재무 수집 -----------------------------
# 인증키는 상위 .env에서 자동 로드(_load_env가 os.environ에 주입). 키 이름은 아래 상수로 중앙화.
DART_API_KEY_ENV = "DART_API_BUSINESS_KEY"   # .env의 DART 인증키 이름
KRX_API_KEY_ENV = "KRX_API_BUSINESS_KEY"     # .env의 KRX 인증키 이름
DART_BASE = "https://opendart.fss.or.kr/api"
DART_MAX_PER_MIN = 900            # 분당 호출 상한 (안전 페이싱; 상태코드 020=한도초과 시 재시도)
DART_MAX_RETRIES = 5

# KRX Data Marketplace OpenAPI — 일별매매정보에서 시가총액(MKTCAP) 수집. 헤더 AUTH_KEY 인증.
KRX_BASE = "https://data-dbg.krx.co.kr/svc/apis/sto"
KRX_DAILY_TRADE = {                # 시장별 일별매매정보 엔드포인트
    "KOSPI": "/stk_bydd_trd",
    "KOSDAQ": "/ksq_bydd_trd",
    "KONEX": "/knx_bydd_trd",      # 코넥스(있으면). 실패 시 건너뜀.
}
# 공식 KRX 결과가 이 수 미만이면 일부 시장 미인가로 보고 폴백(KOSPI만≈946 < 임계 < 양시장≈2600)
KRX_MIN_COVERAGE = 1500

# 기간 설정 (2026-06 기준; 12월 결산 가정. 비-12월 결산사는 Stage2에서 필터됨)
ANNUAL_YEAR = "2025"             # 최신 사업보고서 연도 (온기)
ANNUAL_FALLBACK_YEAR = "2024"    # 2025 미제출 시 폴백
LTM_QUARTER_YEAR = "2026"        # 최신 분기 연도
LTM_QUARTER_CODE = "11013"       # 1분기보고서 (45일 내 제출 → 6월엔 가용)
REPRT_ANNUAL = "11011"           # 사업보고서

# 재무제표 기준 우선순위 (연결 우선, 없으면 별도)
FS_DIV_PREFERENCE = ("CFS", "OFS")

# 수집 산출물
UNIVERSE_FIN_CSV = CACHE_DIR / "universe_financials.csv"
DART_CORPCODE_CACHE = CACHE_DIR / "dart_corpcode_map.csv"   # 종목코드→corp_code 1회 캐시
FIN_CKPT_JSON = CACHE_DIR / "financials_progress.json"      # 수집 중단 재개 체크포인트
COOCCUR_PRIOR_JSON = CACHE_DIR / "cooccur_prior.json"       # Stage1 Lane C 동시출현 prior(데이터 산출)
DART_DOC_CACHE_DIR = CACHE_DIR / "dart_docs"                # 신고서 원본문서 본문 텍스트 캐시(rcept_no별)
FILING_LABELS_CACHE = CACHE_DIR / "filing_labels_cache.json"  # 신고서별 추출 라벨 캐시(Gemini 재호출 방지)


# ----------------------------- 생성 LLM (Stage 3 사업유사성 판정) -----------------------------
# 모델 버전 핀. 접근 가능한 모델로 조정(2026-06 기준 gemini-2.0-flash 종료, 2.5/3.x 계열 사용).
GEMINI_LLM_MODEL = "gemini-2.5-flash"
LLM_JUDGE_CACHE = CACHE_DIR / "llm_judge_cache.json"   # (발행사,후보) 판정 영구 캐시 → 재현성
LLM_PROMPT_VERSION = "v1"     # 프롬프트 개정 시 올리면 캐시 자동 무효화(키에 포함)
LLM_MAX_RETRIES = 5

# 후보 사업설명 텍스트 컬럼 (LLM 판정 입력). 주요제품 중심 + 업종 보강.
BUSINESS_TEXT_COLS = ("주요제품", "업종")


# ----------------------------- Stage 1 -----------------------------
@dataclass
class Stage1Config:
    top_k_embed: int = 200       # Lane B 임베딩 top-K
    use_mid_fallback: bool = True  # Lane A 중분류 fallback 사용 여부
    recall_cap: int = 250        # 모집단 상한 (결정론 히트 우선 유지, 임베딩-only는 점수순 cap)


# ----------------------------- Stage 3 -----------------------------
@dataclass
class Stage3Config:
    top_k_judge: int = 35                      # 임베딩 상위 K개만 LLM 판정(recall↑, 캐시라 1회 비용)
    keep_tiers: tuple = ("높음", "중간")        # 최종 유지 tier (낮음 탈락)
    temperature: float = 0.0                   # 첫 실행 분산 최소화 (재현성은 캐시가 보장)


# ----------------------------- Stage 4 -----------------------------
@dataclass
class Stage4Config:
    per_basis: str = "net_income_ltm"          # PER 분모 (general=LTM 순이익)
    # 하드 컷 — 스타트업 비교 부적합 상장사 제외 (None이면 해당 컷 비활성)
    mcap_max: float | None = 1000e8            # 시총 ≥ 1000억 제외 (대형사)
    per_min: float | None = 10.0               # PER < 10 제외 (저PER 가치주)
    per_max: float | None = 100.0              # PER ≥ 100 제외 (과대평가)
    drop_nonpositive_per: bool = True          # 적자/PER≤0 제거
    # 통계적 이상치 — 하드컷 이후 보조 (기본 none: 하드컷이 주 필터)
    outlier_method: str = "none"               # "none" | "iqr" | "trim"
    iqr_k: float = 1.5
    trim_each_side: int = 1
    min_peers: int = 3                          # 이하면 '피어 부족' 경고 (하드컷은 완화 안 함)
    max_peers: int = 8                          # 유사도 상위 후보 밴드 크기
    headline: str = "median"                    # 대표 멀티플: "median" | "mean" | "sim_weighted"
    # 고유사도 예외 — 사업 매우 유사한 피어는 PER 상·하한 컷에서 면제(핵심 피어 보존)
    high_sim_exempt: bool = False               # True면 아래 조건의 피어를 per_min/per_max 컷 면제
    high_sim_tier: str = "높음"                  # 면제 기준 tier
    high_sim_min_score: float = 85.0            # 또는 sim_score 이 값 이상이면 면제

# -------------------- track_type별 조건세트 (Stage 2 / Stage 4) --------------------
# 2단계(재무)·4단계(일반/outlier)의 트랙별 엄격도 차이를 데이터로 분리.
#  - general      : 순이익 시현 단일 조건 + PER 양방향 컷 + 시총 하한
#  - tech_special : 영업+순이익을 온기·LTM 양쪽 요구 + PER 상방만 컷, 시총 기준 없음
#  - konex_transfer: 별도 필링 구조 → 우선 general 준용, 케이스별 수동 검토 플래그
TRACK_RULES = {
    "general": {
        "financial": {
            "require": ["net_income_annual>0"],   # 신고서 2차 기준=연간 순이익 흑자(LTM 아님)
            "fiscal_month": 12,
            "audit": "적정",
        },
        "general": {"per_min": 10, "per_max": 50, "mktcap_min_bn": 100},
    },
    "tech_special": {
        "financial": {
            "require": [
                "op_income_annual>0", "net_income_annual>0",
                "op_income_ltm>0", "net_income_ltm>0",
            ],
            "fiscal_month": 12,
            "audit": "적정",
        },
        "general": {"per_min": None, "per_max": 50, "mktcap_min_bn": None},
    },
    "konex_transfer": {
        "financial": {
            "require": ["net_income_ltm>0"],
            "fiscal_month": 12,
            "audit": "적정",
        },
        "general": {"per_min": 10, "per_max": 50, "mktcap_min_bn": 100},
        "_note": "KONEX 이전상장 — 필링 구조 상이, 케이스별 수동 검토 필요",
    },
}

VALID_TRACKS = tuple(TRACK_RULES.keys())