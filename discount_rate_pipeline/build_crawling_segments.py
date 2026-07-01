"""
Add one-off crawled segment labels to training_features.csv.

Column naming intentionally uses crawling_* rather than a vendor/source name:
  - crawling_sector
  - crawling_industry
  - crawling_summary
  - crawling_source_url
  - crawling_collected_at

The fetched text is converted into simple segment_* flags for modeling.
"""

import argparse
import json
import re
import sys
import time
from datetime import datetime
from pathlib import Path
from typing import Dict, Optional

import pandas as pd
import requests
from bs4 import BeautifulSoup


for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass


COMPANY_INFO_URL = "https://navercomp.wisereport.co.kr/v2/company/c1010001.aspx?cmp_cd={code}"
HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0 Safari/537.36"
    ),
    "Referer": "https://finance.naver.com/",
}

SEGMENT_KEYWORDS = {
    "segment_bio": [
        "바이오",
        "신약",
        "제약",
        "의약",
        "임상",
        "항암",
        "유전자",
        "세포",
        "치료제",
    ],
    "segment_healthcare": [
        "헬스케어",
        "의료",
        "진단",
        "의료기기",
        "병원",
        "마이크로니들",
        "디지털헬스",
    ],
    "segment_software": [
        "소프트웨어",
        "솔루션",
        "플랫폼",
        "SaaS",
        "보안",
        "클라우드",
        "데이터",
    ],
    "segment_ai": [
        "AI",
        "인공지능",
        "머신러닝",
        "딥러닝",
        "영상인식",
    ],
    "segment_robotics": [
        "로봇",
        "로보틱스",
        "자동화",
        "제어",
    ],
    "segment_semiconductor": [
        "반도체",
        "팹리스",
        "웨이퍼",
        "PCB",
        "장비",
    ],
    "segment_materials": [
        "소재",
        "화학",
        "배터리",
        "2차전지",
        "금속",
        "알루미늄",
    ],
    "segment_manufacturing": [
        "제조",
        "부품",
        "생산",
        "가공",
        "전자부품",
    ],
    "segment_mobility": [
        "자동차",
        "모빌리티",
        "전장",
        "자율주행",
        "차량",
    ],
    "segment_fintech": [
        "핀테크",
        "금융",
        "결제",
        "뱅크",
        "보험",
    ],
    "segment_deeptech": [
        "우주",
        "항공",
        "위성",
        "양자",
        "딥테크",
        "센서",
        "레이더",
    ],
    "segment_platform": [
        "플랫폼",
        "커머스",
        "콘텐츠",
        "미디어",
        "광고",
    ],
    "segment_consumer": [
        "화장품",
        "식품",
        "소비재",
        "브랜드",
        "뷰티",
    ],
}


def normalize_stock_code(value) -> str:
    if value is None:
        return ""
    text = str(value).strip().upper()
    if text == "" or text.lower() == "nan":
        return ""
    if re.fullmatch(r"\d+\.0", text):
        text = text[:-2]
    text = re.sub(r"[^0-9A-Z]", "", text)
    if not re.fullmatch(r"[0-9A-Z]{1,6}", text):
        return ""
    return text.zfill(6) if text else ""


def clean_text(text: str) -> str:
    return re.sub(r"\s+", " ", str(text or "")).strip()


def parse_company_info(html: str) -> Dict[str, str]:
    soup = BeautifulSoup(html, "lxml")
    all_text = clean_text(soup.get_text(" ", strip=True))

    sector = ""
    industry = ""

    m = re.search(r"KOSDAQ\s*:\s*([^W]+?)(?=\s*WICS\s*:)", all_text)
    if not m:
        m = re.search(r"KOSPI\s*:\s*([^W]+?)(?=\s*WICS\s*:)", all_text)
    if m:
        sector = clean_text(m.group(1))

    m = re.search(r"WICS\s*:\s*([^E]+?)(?=\s*EPS\s+|$)", all_text)
    if m:
        industry = clean_text(m.group(1))

    summary_el = soup.select_one(".cmp_comment")
    summary = clean_text(summary_el.get_text(" ", strip=True)) if summary_el else ""

    return {
        "crawling_sector": sector,
        "crawling_industry": industry,
        "crawling_summary": summary,
    }


def fetch_company_info(code: str, cache_dir: Path, rate_limit_sec: float) -> Dict[str, str]:
    cache_dir.mkdir(parents=True, exist_ok=True)
    cache_file = cache_dir / f"{code}.json"
    if cache_file.exists():
        return json.loads(cache_file.read_text(encoding="utf-8"))

    url = COMPANY_INFO_URL.format(code=code)
    data = {
        "crawling_sector": "",
        "crawling_industry": "",
        "crawling_summary": "",
        "crawling_source_url": url,
        "crawling_collected_at": datetime.now().isoformat(timespec="seconds"),
        "crawling_status": "not_started",
    }

    try:
        resp = requests.get(url, headers=HEADERS, timeout=20)
        resp.raise_for_status()
        resp.encoding = "utf-8"
        data.update(parse_company_info(resp.text))
        data["crawling_status"] = "ok" if (
            data["crawling_sector"] or data["crawling_industry"] or data["crawling_summary"]
        ) else "empty"
    except Exception as exc:
        data["crawling_status"] = f"error:{type(exc).__name__}"

    cache_file.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    time.sleep(rate_limit_sec)
    return data


def add_segment_flags(df: pd.DataFrame) -> pd.DataFrame:
    text = (
        df.get("crawling_sector", "").fillna("").astype(str)
        + " "
        + df.get("crawling_industry", "").fillna("").astype(str)
        + " "
        + df.get("crawling_summary", "").fillna("").astype(str)
    )

    for col, keywords in SEGMENT_KEYWORDS.items():
        pattern = "|".join(re.escape(k) for k in keywords)
        df[col] = text.str.contains(pattern, case=False, regex=True).astype(int)

    segment_cols = list(SEGMENT_KEYWORDS.keys())
    df["segment_unknown"] = (df[segment_cols].sum(axis=1) == 0).astype(int)
    return df


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--input", default="training_features.csv")
    parser.add_argument("--output", default="training_features_segmented.csv")
    parser.add_argument("--cache-dir", default="crawling_cache")
    parser.add_argument("--rate-limit-sec", type=float, default=0.2)
    parser.add_argument("--max-rows", type=int, default=None)
    args = parser.parse_args()

    df = pd.read_csv(args.input, encoding="utf-8-sig", dtype={"stock_code": str})
    if "stock_code" not in df.columns:
        raise SystemExit("stock_code column is required")

    target_idx = df.index if args.max_rows is None else df.index[: args.max_rows]
    cache_dir = Path(args.cache_dir)

    for i, idx in enumerate(target_idx, start=1):
        code = normalize_stock_code(df.at[idx, "stock_code"])
        company = df.at[idx, "company_name"] if "company_name" in df.columns else ""
        if not code:
            info = {
                "crawling_sector": "",
                "crawling_industry": "",
                "crawling_summary": "",
                "crawling_source_url": "",
                "crawling_collected_at": datetime.now().isoformat(timespec="seconds"),
                "crawling_status": "no_stock_code",
            }
        else:
            info = fetch_company_info(code, cache_dir, args.rate_limit_sec)

        for key, value in info.items():
            df.at[idx, key] = value

        if i % 20 == 0 or i == len(target_idx):
            print(f"{i}/{len(target_idx)} processed")

    df = add_segment_flags(df)
    df.to_csv(args.output, index=False, encoding="utf-8-sig")

    print(f"Saved: {args.output}")
    print(f"Rows: {len(df)}")
    print("\nCrawling status:")
    print(df["crawling_status"].value_counts(dropna=False).to_string())
    print("\nSegments:")
    for col in list(SEGMENT_KEYWORDS.keys()) + ["segment_unknown"]:
        print(f"  {col:24s}: {int(df[col].sum()):3d}")


if __name__ == "__main__":
    main()
