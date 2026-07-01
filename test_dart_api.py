"""
OPEN DART API 키 일괄 테스트 스크립트
======================================

발급받은 API 키가 정상 작동하는지 주요 엔드포인트를 순차적으로 호출해 확인합니다.

사용법:
    1. 프로젝트 루트에 .env 파일 생성
        DART_API_BUSINESS_KEY=발급받은_40자리_키
    2. 실행
        python dart_api_test.py

요구사항:
    pip install requests python-dotenv
"""

import os
import sys
import json
import time
from datetime import datetime, timedelta

import requests
from dotenv import load_dotenv

load_dotenv()

BASE_URL = "https://opendart.fss.or.kr/api"
API_KEY = (os.getenv("DART_API_BUSINESS_KEY") or "").strip()

# 테스트용 기본값 (삼성전자)
SAMSUNG_CORP_CODE = "00126380"   # 삼성전자 고유번호
SAMSUNG_STOCK = "005930"

# 색상 출력 (터미널)
class C:
    OK = "\033[92m"
    FAIL = "\033[91m"
    WARN = "\033[93m"
    INFO = "\033[96m"
    END = "\033[0m"
    BOLD = "\033[1m"


def banner(text: str) -> None:
    line = "=" * 60
    print(f"\n{C.BOLD}{line}\n{text}\n{line}{C.END}")


def call(name: str, endpoint: str, params: dict) -> dict | None:
    """엔드포인트 호출 후 결과 요약 출력."""
    url = f"{BASE_URL}/{endpoint}"
    masked = {**params, "crtfc_key": params["crtfc_key"][:6] + "..." + params["crtfc_key"][-4:]}
    print(f"\n{C.INFO}[{name}]{C.END} GET {endpoint}")
    print(f"  params: {masked}")

    try:
        r = requests.get(url, params=params, timeout=15)
    except requests.RequestException as e:
        print(f"  {C.FAIL}✗ 네트워크 오류: {e}{C.END}")
        return None

    print(f"  HTTP {r.status_code}")

    # 일부 엔드포인트는 zip 바이너리(고유번호 corpCode.xml) 반환
    ctype = r.headers.get("Content-Type", "")
    if "json" not in ctype.lower():
        if r.status_code == 200 and len(r.content) > 0:
            print(f"  {C.OK}✓ 바이너리 응답 수신 ({len(r.content):,} bytes, {ctype}){C.END}")
            return {"status": "000", "binary": True, "size": len(r.content)}
        print(f"  {C.FAIL}✗ 예상치 못한 응답 (Content-Type: {ctype}){C.END}")
        return None

    try:
        data = r.json()
    except json.JSONDecodeError:
        print(f"  {C.FAIL}✗ JSON 파싱 실패{C.END}")
        return None

    status = data.get("status")
    message = data.get("message", "")
    interpret(status, message, data)
    return data


def interpret(status: str, message: str, data: dict) -> None:
    """OPEN DART status 코드 해석."""
    table = {
        "000": (C.OK, "정상"),
        "010": (C.FAIL, "등록되지 않은 키"),
        "011": (C.FAIL, "사용할 수 없는 키 (오픈API 이용등록 필요)"),
        "012": (C.FAIL, "접근할 수 없는 IP"),
        "013": (C.WARN, "조회된 데이터 없음"),
        "014": (C.FAIL, "파일이 존재하지 않음"),
        "020": (C.FAIL, "요청 제한 초과 (분당 호출한도)"),
        "021": (C.FAIL, "조회 가능 한도 초과"),
        "100": (C.FAIL, "필드 오류"),
        "101": (C.FAIL, "부적절한 접근"),
        "800": (C.WARN, "시스템 점검 중"),
        "900": (C.FAIL, "정의되지 않은 오류"),
        "901": (C.FAIL, "사용자 계정 보안 조치"),
    }
    color, label = table.get(status, (C.WARN, "알 수 없는 status"))
    print(f"  {color}status={status} ({label}) message={message!r}{C.END}")

    if status == "000":
        # 데이터 미리보기
        if "list" in data and isinstance(data["list"], list):
            print(f"  {C.OK}✓ list {len(data['list'])}건 수신{C.END}")
            if data["list"]:
                sample = data["list"][0]
                preview_keys = list(sample.keys())[:5]
                preview = {k: sample[k] for k in preview_keys}
                print(f"  샘플: {preview}")
        elif data.get("corp_name"):
            print(f"  {C.OK}✓ 회사명: {data.get('corp_name')} / 종목: {data.get('stock_code')}{C.END}")


def main() -> int:
    banner("OPEN DART API 키 일괄 테스트")

    if not API_KEY:
        print(f"{C.FAIL}DART_API_BUSINESS_KEY 를 찾을 수 없습니다.{C.END}")
        print("  프로젝트 루트의 .env 파일에 다음 줄을 추가하세요:")
        print("    DART_API_BUSINESS_KEY=발급받은_40자리_키")
        return 1

    if len(API_KEY) != 40:
        print(f"{C.WARN}경고: 일반적인 DART API 키 길이는 40자입니다. 현재 {len(API_KEY)}자.{C.END}")

    # 최근 7일 범위 (공시검색용)
    today = datetime.now()
    bgn_de = (today - timedelta(days=7)).strftime("%Y%m%d")
    end_de = today.strftime("%Y%m%d")

    results: list[tuple[str, str | None]] = []

    # 1) 공시검색 - 가장 기본
    r1 = call(
        "1. 공시검색 list.json",
        "list.json",
        {
            "crtfc_key": API_KEY,
            "corp_code": SAMSUNG_CORP_CODE,
            "bgn_de": bgn_de,
            "end_de": end_de,
            "page_count": 10,
        },
    )
    results.append(("공시검색", r1.get("status") if r1 else None))

    time.sleep(0.3)

    # 2) 기업개황
    r2 = call(
        "2. 기업개황 company.json",
        "company.json",
        {"crtfc_key": API_KEY, "corp_code": SAMSUNG_CORP_CODE},
    )
    results.append(("기업개황", r2.get("status") if r2 else None))

    time.sleep(0.3)

    # 3) 단일회사 주요계정 (재무제표)
    r3 = call(
        "3. 단일회사 주요계정 fnlttSinglAcnt.json",
        "fnlttSinglAcnt.json",
        {
            "crtfc_key": API_KEY,
            "corp_code": SAMSUNG_CORP_CODE,
            "bsns_year": str(today.year - 1),
            "reprt_code": "11011",  # 사업보고서
        },
    )
    results.append(("재무제표", r3.get("status") if r3 else None))

    time.sleep(0.3)

    # 4) 고유번호 전체 다운로드 (zip 바이너리)
    r4 = call(
        "4. 고유번호 corpCode.xml",
        "corpCode.xml",
        {"crtfc_key": API_KEY},
    )
    results.append(("고유번호", r4.get("status") if r4 else None))

    # 요약
    banner("테스트 결과 요약")
    for name, status in results:
        if status == "000":
            mark = f"{C.OK}✓ PASS{C.END}"
        elif status == "013":
            mark = f"{C.WARN}△ 데이터 없음 (키는 정상){C.END}"
        elif status is None:
            mark = f"{C.FAIL}✗ 호출 실패{C.END}"
        else:
            mark = f"{C.FAIL}✗ FAIL (status={status}){C.END}"
        print(f"  {name:12s} → {mark}")

    # 키 유효성 종합 판정
    ok_or_no_data = {"000", "013"}
    statuses = {s for _, s in results if s is not None}
    if statuses and statuses.issubset(ok_or_no_data):
        print(f"\n{C.OK}{C.BOLD}✓ API 키 정상 작동{C.END}")
        return 0
    else:
        print(f"\n{C.FAIL}{C.BOLD}✗ 일부 호출 실패 — 위 status 코드 확인 필요{C.END}")
        return 2


if __name__ == "__main__":
    sys.exit(main())