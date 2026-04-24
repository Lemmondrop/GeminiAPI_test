import os
import time
import re
import requests
import random
from bs4 import BeautifulSoup
import pandas as pd
import io
import concurrent.futures

# =========================================================
# 0. 종목코드 조회 함수
# =========================================================
def get_random_ua():
    UA = [
        'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36',
        'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/92.0.4515.107 Safari/537.36'
    ]
    return random.choice(UA)

def get_stock_code_from_csv(company_name: str, csv_path: str) -> str:
    """
    회사명으로 CSV에서 종목코드 찾기 + 6자리 패딩
    
    Args:
        company_name: 회사명
        csv_path: company_cord_prototype.csv 경로
    
    Returns:
        6자리 종목코드 (예: "064400") or None
    """
    if not os.path.exists(csv_path):
        return None
    
    encodings = ['cp949', 'euc-kr', 'utf-8-sig', 'utf-8']
    
    for enc in encodings:
        try:
            df = pd.read_csv(csv_path, encoding=enc)
            
            # 회사명 컬럼 찾기
            name_col = None
            for col in df.columns:
                if '회사명' in col:
                    name_col = col
                    break
            
            if not name_col:
                continue
            
            # 종목코드 컬럼 찾기  
            code_col = None
            for col in df.columns:
                if '종목코드' in col:
                    code_col = col
                    break
            
            if not code_col:
                continue
            
            # 정확히 일치하는 회사 찾기
            matched = df[df[name_col] == company_name.strip()]
            
            if len(matched) == 0:
                # 부분 일치 검색
                matched = df[df[name_col].str.contains(company_name.strip(), na=False, case=False)]
            
            if len(matched) > 0:
                code = str(matched.iloc[0][code_col]).strip()
                # 6자리 패딩
                code = code.zfill(6)
                return code
            
        except Exception as e:
            continue
    
    return None

# =========================================================
# 3단계: 사업 유사성 필터링
# =========================================================

def get_business_description(company_code: str) -> dict:
    """
    WiseReport(네이버 기업개요 원본)에서 기업 개요 + 주요제품 매출구성(cTB203) 크롤링
    Returns: {"business": str, "main_products": str}
    """
    time.sleep(random.uniform(0.4, 0.9))

    url = f"https://navercomp.wisereport.co.kr/v2/company/c1020001.aspx?cmp_cd={company_code}&cn="

    headers = {
        "User-Agent": get_random_ua(),
        # referer는 있어도 되고 없어도 되는 경우가 많지만, 안정성 위해 유지
        "Referer": f"https://finance.naver.com/item/main.naver?code={company_code}",
        "Accept-Language": "ko-KR,ko;q=0.9,en-US;q=0.8,en;q=0.7",
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        "Connection": "keep-alive",
    }

    try:
        res = requests.get(url, headers=headers, timeout=10, allow_redirects=True)
        if res.status_code != 200:
            return {"business": "", "main_products": ""}

        # WiseReport는 보통 utf-8
        if not res.encoding or res.encoding.lower() in ("iso-8859-1", "latin-1"):
            res.encoding = "utf-8"

        soup = BeautifulSoup(res.text, "html.parser")

        # -------------------------
        # 1) 기업 개요 텍스트(최소한이라도 확보)
        #    - selector가 확정되면 그걸로 교체 권장
        # -------------------------
        business_summary = ""
        candidates = []
        for node in soup.find_all(["div", "p", "td"]):
            txt = node.get_text(" ", strip=True)
            if len(txt) >= 200:
                candidates.append((len(txt), txt))
        if candidates:
            candidates.sort(reverse=True, key=lambda x: x[0])
            business_summary = candidates[0][1][:700]

        # -------------------------
        # 2) 주요제품 매출구성(cTB203)
        # -------------------------
        main_products = ""
        table = soup.find("table", id="cTB203")
        if table:
            products = []
            for tr in table.select("tbody tr"):
                th = tr.select_one("th span.cut") or tr.select_one("th")
                td = tr.select_one("td.c2.num") or tr.select_one("td")
                if not th or not td:
                    continue

                name = th.get_text(" ", strip=True).replace("\xa0", "").strip()
                ratio = td.get_text(" ", strip=True).replace("\xa0", "").strip()

                # 빈 행 제외
                if not name or not ratio or name == "&nbsp;" or ratio == "&nbsp;":
                    continue

                products.append(f"{name} ({ratio}%)")

            if products:
                main_products = ", ".join(products)

        return {"business": business_summary, "main_products": main_products}

    except Exception:
        return {"business": "", "main_products": ""}

def check_business_similarity(target_business: str, peer_info: dict, threshold: float = 0.3) -> tuple:
    """
    Gemini를 사용하여 타겟 기업과 Peer의 사업 유사도 판정
    
    Args:
        target_business: 타겟 기업의 사업 설명 (PDF에서 추출)
        peer_info: {"name": str, "code": str, "business": str, "main_products": str}
        threshold: 유사도 임계치 (0.3 = 30% 이상 유사)
    
    Returns:
        (company_name, passed: bool, similarity_score: float, reason: str)
    """
    from utils import call_gemini, safe_json_loads
    
    name = peer_info['name']
    peer_business = peer_info.get('business', '')
    peer_products = peer_info.get('main_products', '')
    
    # 주요제품 정보를 우선적으로 사용
    peer_description = f"{peer_business}\n주요제품 구성: {peer_products}" if peer_products else peer_business
    
    if not peer_description.strip():
        return (name, False, 0.0, "사업 정보 없음 (네이버 증권 크롤링 실패)")
    
    prompt = f"""
당신은 사업 유사성 분석 전문가입니다.

[타겟 기업 사업 내용]
{target_business}

[비교 기업: {name}]
{peer_description}

위 두 기업의 사업이 얼마나 유사한지 0.0~1.0 점수로 평가하십시오.

[평가 기준]
- 1.0: 거의 동일한 사업 (같은 제품/서비스, 같은 시장)
- 0.7~0.9: 상당히 유사 (관련 산업, 유사 제품/고객)
- 0.5~0.7: 부분 유사 (일부 사업 영역 겹침)
- 0.3~0.5: 약간 유사 (간접적 관련성)
- 0.0~0.3: 유사성 낮음 (다른 산업/제품)

[중요]
- 주요제품 구성비를 중요하게 고려하십시오
- 구체적인 제품명/서비스명이 유사한지 판단하십시오
- 산업 분류만 같다고 높은 점수를 주지 마십시오

[Output JSON]
{{
  "similarity_score": 0.0~1.0,
  "reason": "유사성 판단 근거 (1~2문장, 주요제품 언급)"
}}
"""
    
    res = call_gemini(prompt, max_tokens=500)
    if not res.get("ok"):
        return (name, False, 0.0, "AI 분석 실패")
    
    data = safe_json_loads(res.get("text", ""))
    score = float(data.get("similarity_score", 0.0))
    reason = data.get("reason", "근거 없음")
    
    passed = score >= threshold
    
    return (name, passed, score, reason)


def filter_peers_stage3(
    target_pdf_path: str,
    peer_companies: list,
    company_name: str,
    threshold: float = 0.3,
    max_workers: int = 5
) -> dict:
    from utils import call_gemini, safe_json_loads
    import concurrent.futures
    
    print(f"   📋 [Step 3] 사업 유사성 필터링 시작 (Input: {len(peer_companies)}개 사)")
    
    # 1) 타겟 기업의 사업 설명 추출
    prompt_target = f"""
당신은 사업 분석 전문가입니다.
제공된 IR 자료에서 '{company_name}'의 핵심 사업 내용을 200자 이내로 요약하십시오.
- 주요 제품/서비스
- 타겟 고객
- 핵심 기술/사업 모델

[Output JSON]
{{ "business_summary": "사업 요약" }}
"""
    
    res = call_gemini(prompt_target, pdf_path=target_pdf_path, max_tokens=1000)
    if not res.get("ok"):
        print("      ⚠️ 타겟 기업 사업 추출 실패 - 필터링 스킵")
        return {"business_passed": [p['name'] for p in peer_companies], "similarity_details": []}
    
    target_business = safe_json_loads(res.get("text", "")).get("business_summary", "")
    if not target_business:
        print("      ⚠️ 타겟 사업 정보 없음 - 필터링 스킵")
        return {"business_passed": [p['name'] for p in peer_companies], "similarity_details": []}
    
    print(f"      👉 타겟 사업: {target_business[:100]}...")
    
    # 2) 각 Peer 기업의 사업 정보 크롤링
    print(f"      👉 Peer 기업 사업 정보 크롤링 중...")
    
    peer_business_info = []
    for peer in peer_companies:
        biz_info = get_business_description(peer['code'])
        peer_business_info.append({
            "name": peer['name'],
            "code": peer['code'],
            "business": biz_info['business'],
            "main_products": biz_info['main_products']
        })
    
    # 3) 🚨 오리지널 복구: ThreadPoolExecutor를 이용한 전체 기업 동시 병렬 분석 (컷오프 없음)
    print(f"      👉 사업 유사도 분석 중...")
    
    results = []
    with concurrent.futures.ThreadPoolExecutor(max_workers=max_workers) as executor:
        futures = [
            executor.submit(check_business_similarity, target_business, peer_info, threshold)
            for peer_info in peer_business_info
        ]
        for future in concurrent.futures.as_completed(futures):
            results.append(future.result())
    
    # 4) 통과 기업 필터링
    business_passed = []
    similarity_details = []
    
    # peer_business_info와 results를 매칭하여 주요제품 정보 포함
    peer_info_map = {p['name']: p for p in peer_business_info}
    
    for name, passed, score, reason in results:
        peer_info = peer_info_map.get(name, {})
        main_products = peer_info.get('main_products', '정보 없음')
        
        similarity_details.append({
            "company": name,
            "score": score,
            "reason": reason,
            "main_products": main_products
        })
        
        if passed:
            business_passed.append(name)
    
    # 정렬 (유사도 높은 순)
    similarity_details.sort(key=lambda x: x['score'], reverse=True)
    
    print(f"      👉 사업 유사성 통과: {len(business_passed)}개 사")
    
    # 모든 기업 상세 정보 출력 (통과/미통과 구분)
    print(f"\n      📊 사업 유사도 분석 결과 (주요제품 구성 포함):")
    print(f"      " + "="*70)
    
    for item in similarity_details[:15]:  # 콘솔 출력만 상위 15개
        name = item['company']
        score = item['score']
        reason = item['reason']
        main_products = item['main_products']
        
        status = "✅ 통과" if name in business_passed else "❌ 미통과"
        print(f"\n      {status} [{name}] 유사도 {score:.2f}")
        print(f"         주요제품: {main_products}")
        print(f"         판단근거: {reason[:80]}...")
    
    return {
        "business_passed": business_passed,
        "similarity_details": similarity_details
    }

# =========================================================
# 4단계: 일반 요건 유사성 필터링
# =========================================================

def get_listing_info(company_code: str, debug: bool = False) -> dict:
    """
    네이버 증권(item/main.naver)에서 관리종목 여부, 시가총액, PER, PBR, EV/EBITDA 추출
    - 시가총액: '시가총액' 테이블 (div.first) / em#_market_sum
    - PER/PBR: table.per_table / em#_per, em#_pbr
    - EV/EBITDA: '투자정보' 테이블 (table.gHead03) 내부 th 파싱
    """
    from datetime import datetime
    import time, random, re
    import requests
    from bs4 import BeautifulSoup

    def clean_text(s: str) -> str:
        if not s:
            return ""
        return re.sub(r"\s+", " ", s).replace("\xa0", " ").strip()

    def parse_market_cap_eok(text: str) -> float:
        """
        예: '40조1,935' / '1조6,222' / '6,222' 등을 '억원' 단위 숫자로 변환
        - 조(兆) = 10,000억원
        """
        if not text:
            return 0.0
        s = re.sub(r"\s+", "", text)
        s = s.replace("억원", "")

        m = re.search(r"(?:(\d+(?:,\d+)*)조)?(?:(\d+(?:,\d+)*))?$", s)
        if not m:
            m2 = re.search(r"([\d,]+)", s)
            return float(m2.group(1).replace(",", "")) if m2 else 0.0

        jo_part = m.group(1)
        eok_part = m.group(2)

        jo = float(jo_part.replace(",", "")) if jo_part else 0.0
        eok = float(eok_part.replace(",", "")) if eok_part else 0.0
        return jo * 10000.0 + eok

    def parse_float_first(text: str):
        """문자열에서 첫 번째 숫자(float)만 뽑아 반환. 없으면 None."""
        if not text:
            return None
        t = text.replace(",", "")
        m = re.search(r"([-+]?\d+(?:\.\d+)?)", t)
        if not m:
            return None
        try:
            return float(m.group(1))
        except:
            return None

    time.sleep(random.uniform(0.3, 0.8))
    url = f"https://finance.naver.com/item/main.naver?code={company_code}"

    result = {
        "listing_date": None,
        "is_warning": False,
        "market_cap": 0.0,
        "per": None,
        "pbr": None,
        "ev_ebitda": None, # 🚨 [추가] EV/EBITDA 키 추가
        "fetch_date": datetime.now().strftime("%Y-%m-%d"),
    }

    try:
        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
            "Referer": "https://finance.naver.com/",
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8",
            "Accept-Language": "ko-KR,ko;q=0.9,en-US;q=0.8,en;q=0.7",
            "Connection": "keep-alive",
        }

        res = requests.get(url, headers=headers, timeout=15)
        if res.status_code != 200:
            if debug:
                print(f"      [DEBUG] HTTP {res.status_code} - {company_code}")
            return result

        # 네이버 금융은 euc-kr인 경우가 많음 (utf-8로 강제하면 깨지는 케이스 방지)
        if not res.encoding or res.encoding.lower() in ("iso-8859-1", "latin-1"):
            res.encoding = "euc-kr"

        html = res.text
        if len(html) < 1000:
            if debug:
                print(f"      [DEBUG] HTML too short ({len(html)} bytes) - {company_code}")
            return result

        soup = BeautifulSoup(html, "html.parser")

        # =========================================================
        # 1) 관리종목/투자위험 여부
        # =========================================================
        wrap = soup.select_one("div.wrap_company")
        if wrap:
            wrap_text = clean_text(wrap.get_text(" ", strip=True))
            if any(k in wrap_text for k in ["관리종목", "투자위험", "투자주의", "투자경고", "거래정지"]):
                result["is_warning"] = True

        # =========================================================
        # 2) 시가총액: '시가총액 정보' 테이블(별도 블록)에서만 추출
        # =========================================================
        market_sum = None

        # 1) 가장 확실한 id
        market_node = soup.select_one("div.first em#_market_sum") or soup.select_one("em#_market_sum")
        if market_node:
            market_sum = clean_text(market_node.get_text(strip=True))

        # 2) fallback: caption 기반으로 테이블 찾아 row 파싱
        if not market_sum:
            cap_table = None
            for t in soup.select("div.first table"):
                cap = t.select_one("caption")
                if cap and "시가총액" in clean_text(cap.get_text()):
                    cap_table = t
                    break
            if cap_table:
                for tr in cap_table.select("tbody tr"):
                    th = tr.select_one("th")
                    td = tr.select_one("td")
                    if not th or not td:
                        continue
                    if "시가총액" in clean_text(th.get_text(" ", strip=True)):
                        em = td.select_one("em")
                        market_sum = clean_text(em.get_text(strip=True)) if em else clean_text(td.get_text(" ", strip=True))
                        break

        if market_sum and market_sum not in ("N/A", "-", "NA"):
            result["market_cap"] = parse_market_cap_eok(market_sum)

        # =========================================================
        # 3) PER / 4) PBR: per_table에서만 추출
        # =========================================================
        per_table = soup.select_one("table.per_table")

        per_raw = None
        pbr_raw = None

        per_node = (per_table.select_one("em#_per") if per_table else None) or soup.select_one("em#_per")
        if per_node:
            per_raw = clean_text(per_node.get_text(strip=True))

        pbr_node = (per_table.select_one("em#_pbr") if per_table else None) or soup.select_one("em#_pbr")
        if pbr_node:
            pbr_raw = clean_text(pbr_node.get_text(strip=True))

        if per_table and (not per_raw or per_raw in ("N/A", "-", "NA")):
            for tr in per_table.select("tr"):
                th = tr.select_one("th")
                td = tr.select_one("td")
                if not th or not td:
                    continue
                if "PER" in clean_text(th.get_text(" ", strip=True)):
                    per_raw = clean_text(td.get_text(" ", strip=True))
                    break

        if per_table and (not pbr_raw or pbr_raw in ("N/A", "-", "NA")):
            for tr in per_table.select("tr"):
                th = tr.select_one("th")
                td = tr.select_one("td")
                if not th or not td:
                    continue
                if "PBR" in clean_text(th.get_text(" ", strip=True)):
                    pbr_raw = clean_text(td.get_text(" ", strip=True))
                    break

        if per_raw and "N/A" not in per_raw and per_raw not in ("-", "NA"):
            result["per"] = parse_float_first(per_raw)

        if pbr_raw and "N/A" not in pbr_raw and pbr_raw not in ("-", "NA"):
            result["pbr"] = parse_float_first(pbr_raw)

        # =========================================================
        # 🚨 [추가] 5) EV/EBITDA 추출 로직 (fnGuide 펀더멘털 실적 테이블)
        # =========================================================
        ev_raw = None
        try:
            # 1. 공유해주신 정확한 wisereport 원본 URL 사용
            fnguide_url = f"https://navercomp.wisereport.co.kr/v2/company/c1010001.aspx?cmp_cd={company_code}"
            
            # 2. 크롤링 차단(403 Forbidden 등)을 막기 위한 전용 헤더
            fn_headers = {
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
                "Referer": f"https://finance.naver.com/item/main.naver?code={company_code}",
                "Accept-Language": "ko-KR,ko;q=0.9,en-US;q=0.8,en;q=0.7"
            }
            
            res_fn = requests.get(fnguide_url, headers=fn_headers, timeout=10)
            
            if res_fn.status_code == 200:
                soup_fn = BeautifulSoup(res_fn.text, "html.parser")
                
                # 3. 제공해주신 HTML 구조 타겟팅 (div.fund.fl_le 안의 table.gHead03)
                fund_table = soup_fn.select_one("div.fund.fl_le table.gHead03")
                
                if fund_table:
                    # 4. table 안의 모든 행(tr)을 순회하며 가장 확실하게 찾기
                    for tr in fund_table.select("tbody tr"):
                        th = tr.select_one("th")
                        if th and "EV/EBITDA" in th.get_text(strip=True).upper():
                            # 해당 행의 td들을 가져옴
                            tds = tr.select("td")
                            if len(tds) > 0:
                                # 첫 번째 td 값이 2024/12(A) 실적값 (예: "5.67")
                                ev_raw = clean_text(tds[0].get_text(strip=True))
                                break
                
                # 5. 숫자(Float) 변환 및 저장
                if ev_raw and ev_raw not in ("N/A", "-", "NA", ""):
                    result["ev_ebitda"] = parse_float_first(ev_raw)
                    
        except Exception as e:
            if debug:
                print(f"      [DEBUG] EV/EBITDA 크롤링 실패: {e}")

        # =========================================================
        # DEBUG
        # =========================================================
        if debug:
            print(
                f"      [DEBUG] {company_code}: 시가총액={result['market_cap']}억, "
                f"PER={result['per']}, PBR={result['pbr']}, EV/EBITDA={result['ev_ebitda']}, warning={result['is_warning']}"
            )
            print(f"      [DEBUG] raw: market_sum='{market_sum}', per_raw='{per_raw}', pbr_raw='{pbr_raw}', ev_raw='{ev_raw}'")

        return result

    except Exception as e:
        if debug:
            print(f"      [DEBUG] 크롤링 실패: {str(e)[:120]} - {company_code}")
        return result
    
def check_general_requirements(company_info: dict, strict_per_bottom: bool = True, debug: bool = False) -> tuple:
    """
    일반 요건 체크 (Python 1차 필터링: 관리종목, PER 조건부 하한 10~100, 시총 1000억 이상)
    """
    code = company_info.get('code', '')
    name = company_info.get('name', '')
    info = get_listing_info(code, debug=False)  

    listing_date = info.get("listing_date")
    is_warning = bool(info.get("is_warning", False))
    market_cap = info.get("market_cap", 0) or 0
    per = info.get("per", None)
    pbr = info.get("pbr", None)
    ev_ebitda = info.get("ev_ebitda", None) # 🚨 EV/EBITDA 추가

    reasons = []

    # 1) 관리종목/투자 위험 경고
    if is_warning:
        reasons.append("관리/투자 위험 종목")

    # 2) 시가총액 조건 (1천억 이상 제외)
    if market_cap and market_cap >= 1000:
        reasons.append(f"시가총액 {int(market_cap)}억 (1천억 이상)")

    # 3) 🚨 PER 조건 (N/A, 적자, 100 초과는 완벽 차단 / 10 미만은 조건부 차단)
    if per is None or per <= 0:
        reasons.append("PER N/A (데이터 없음 또는 적자)")
    else:
        if strict_per_bottom and per < 10: # 🚨 여기에 스위치 적용!
            reasons.append(f"PER {per:.1f} (10 미만)")
        elif per > 100:
            reasons.append(f"PER {per:.1f} (100 초과)")

    passed = (len(reasons) == 0)
    reason_str = ", ".join(reasons) if reasons else "OK"

    # 🚨 info 반환 딕셔너리에 ev_ebitda 추가
    return name, passed, reason_str, {
        "listing_date": listing_date,
        "is_warning": is_warning,
        "market_cap": market_cap,
        "per": per,
        "pbr": pbr,
        "ev_ebitda": ev_ebitda, 
        "fetch_date": info.get("fetch_date"),
    }

def filter_peers_stage4(
    peer_companies: list,  
    max_workers: int = 5,
    debug: bool = False 
) -> dict:
    """
    [4단계] 일반 요건 필터링 및 🚨 아웃라이어(MAX/MIN) 제거 (조건부 하한선 완화 적용)
    """
    import concurrent.futures
    print(f"   🔍 [Step 4] 일반 요건 및 Outlier 필터링 시작 (Input: {len(peer_companies)}개 사)")

    # 🚨 [핵심] 필터링과 아웃라이어 제거 로직을 하나의 내부 함수로 묶어 1차/2차 재실행을 쉽게 만듭니다.
    def run_filtering(strict: bool):
        with concurrent.futures.ThreadPoolExecutor(max_workers=max_workers) as executor:
            # strict 인자(True/False)를 check_general_requirements에 함께 전달합니다.
            futures = [executor.submit(check_general_requirements, company, strict, debug) for company in peer_companies]
            results = [f.result() for f in concurrent.futures.as_completed(futures)]
        
        general_passed_info = []
        details = []
        
        # 1차 절대 기준(PER, 시총 등) 통과 기업 분류
        for name, passed, reason, info in results:
            if passed:
                general_passed_info.append({"name": name, "info": info})
            else:
                details.append((name, False, reason, info))
                
        final_passed_names = []
        
        # -------------------------------------------------------------
        # Python 2차 필터링: MAX/MIN 아웃라이어 제거 로직
        # -------------------------------------------------------------
        if len(general_passed_info) > 4:
            metrics = ['per', 'pbr', 'market_cap', 'ev_ebitda']
            companies_to_drop = set()
            drop_reasons = {}
            
            for metric in metrics:
                valid_peers = [p for p in general_passed_info if p['info'].get(metric) is not None]
                if len(valid_peers) > 2:
                    max_val = max(p['info'][metric] for p in valid_peers)
                    min_val = min(p['info'][metric] for p in valid_peers)
                    
                    for p in valid_peers:
                        if p['info'][metric] == max_val:
                            companies_to_drop.add(p['name'])
                            drop_reasons[p['name']] = f"MAX {metric.upper()} ({max_val})"
                        elif p['info'][metric] == min_val:
                            companies_to_drop.add(p['name'])
                            drop_reasons[p['name']] = f"MIN {metric.upper()} ({min_val})"
            
            # [방어로직] 제외 기업이 너무 많아 최종 피어그룹이 3개 미만이 되는 경우
            if len(general_passed_info) - len(companies_to_drop) < 3:
                # 경고 문구는 1차 시도에서만 출력되도록 처리하거나 생략할 수 있습니다.
                companies_to_drop = set()
                drop_reasons = {}
                per_valid = [p for p in general_passed_info if p['info'].get('per') is not None]
                if len(per_valid) > 2:
                    max_per = max(p['info']['per'] for p in per_valid)
                    min_per = min(p['info']['per'] for p in per_valid)
                    for p in per_valid:
                        if p['info']['per'] in (max_per, min_per):
                            companies_to_drop.add(p['name'])
                            drop_reasons[p['name']] = f"MAX/MIN PER ({p['info']['per']})"
                            
                # 그래도 3개 미만이면 아예 아웃라이어 제거 취소
                if len(general_passed_info) - len(companies_to_drop) < 3:
                    companies_to_drop = set()

            # 최종 탈락/합격 분류 적용
            for p in general_passed_info:
                name = p['name']
                info = p['info']
                if name in companies_to_drop:
                    reason = drop_reasons[name]
                    details.append((name, False, reason, info)) # Outlier로 탈락
                else:
                    details.append((name, True, "OK", info))
                    final_passed_names.append(name)
        else:
            # 기업이 4개 이하면 모두 합격 처리
            for p in general_passed_info:
                details.append((p['name'], True, "OK", p['info']))
                final_passed_names.append(p['name'])
                
        return final_passed_names, details

    # =========================================================
    # 🚨 1차 시도: 엄격한 필터링 (PER 10 미만 차단)
    # =========================================================
    final_passed_names, details = run_filtering(strict=True)

    # =========================================================
    # 🚨 2차 시도 (Fallback): 전멸 시 저평가(PER 10 미만) 허용
    # =========================================================
    if len(final_passed_names) == 0 and len(peer_companies) > 0:
        print("      ⚠️ [Fallback] 1차 필터링 통과 기업이 0개입니다. PER 하한선(10 미만) 조건을 해제하고 재검색합니다.")
        final_passed_names, details = run_filtering(strict=False)

    print(f"      👉 일반 요건 및 Outlier 필터링 최종 통과: {len(final_passed_names)}개 사")
    
    # 탈락 사유(아웃라이어 포함) 출력
    failed = [(n, r, i) for n, p, r, i in details if not p]
    if failed:
        print(f"      ⚠️ 주요 탈락 사유:")
        for name, reason, info in failed[:8]:
            per_str = f"PER {info['per']:.1f}" if info.get('per') is not None else "PER N/A"
            cap_str = f"시총 {info['market_cap']:.0f}억" if info.get('market_cap', 0) > 0 else "시총 N/A"
            
            ebitda_val = info.get('ev_ebitda')
            ebitda_str = f"EV/EBITDA {ebitda_val}" if ebitda_val is not None else "EV/EBITDA N/A"
            
            print(f"         - {name}: {reason} ({per_str}, {cap_str}, {ebitda_str})")
            
    return {
        "general_passed": final_passed_names,
        "details": details
    }

# =========================================================
# 통합 필터링 파이프라인
# =========================================================

def full_peer_filtering_pipeline(
    target_pdf_path: str,
    company_name: str,
    raw_peer_names: list,
    company_csv_path: str,
    similarity_threshold: float = 0.3
) -> dict:
    """
    1~4단계 전체 필터링 파이프라인 (통과 기업 수에 상관없이 끝까지 진행)
    """
    from utils import filter_peers_stage2
    
    print(f"\n{'='*60}")
    print(f"  🎯 Peer Group 정밀 필터링 파이프라인 시작")
    print(f"{'='*60}\n")
    
    # Stage 1: Raw list
    stage1_raw = raw_peer_names
    print(f"✅ [Stage 1] 산업분류 매칭: {len(stage1_raw)}개 사")
    
    # Stage 2: 재무 필터링 (12월 결산 + 흑자)
    stage2_result = filter_peers_stage2(stage1_raw, company_csv_path)
    stage2_profit = stage2_result.get("profit_passed", [])
    
    # [수정] 2단계 통과 기업이 3개 미만이더라도 강제로 3, 4단계를 진행하도록 스킵 로직 제거
    if len(stage2_profit) == 0:
        print("      ⚠️ 2단계 통과 기업이 0개입니다. (이후 단계는 빈 리스트로 진행됩니다)")
    
    # Stage 2 결과를 {"name": str, "code": str} 형태로 변환
    peer_objs = []
    failed_companies = []
    
    for company_name in stage2_profit:
        code = get_stock_code_from_csv(company_name, company_csv_path)
        if code:
            peer_objs.append({"name": company_name, "code": code})
        else:
            failed_companies.append(company_name)
    
    if failed_companies:
        print(f"      ⚠️ 종목코드 찾기 실패: {len(failed_companies)}개사")
        for name in failed_companies[:3]:
            print(f"         - {name}")
    
    # Stage 3: 사업 유사성 필터링
    stage3_result = filter_peers_stage3(
        target_pdf_path=target_pdf_path,
        peer_companies=peer_objs,
        company_name=company_name,
        threshold=similarity_threshold
    )
    stage3_business = stage3_result.get("business_passed", [])
    
    # [수정] 3단계 통과 기업이 부족해도 2단계 결과로 되돌리는(롤백) 로직 제거
    if len(stage3_business) == 0 and len(peer_objs) > 0:
        print("      ⚠️ 3단계 사업 유사성 통과 기업이 0개입니다.")
    
    # Stage 3 통과 기업의 객체만 추출
    stage3_objs = [p for p in peer_objs if p['name'] in stage3_business]
    
    # Stage 4: 일반 요건 필터링
    stage4_result = filter_peers_stage4(stage3_objs, debug=False)
    stage4_final = stage4_result.get("general_passed", [])
    
    # [수정] 4단계 통과 기업이 부족해도 3단계 결과로 되돌리는(롤백) 로직 제거
    if len(stage4_final) == 0 and len(stage3_objs) > 0:
        print("      ⚠️ 4단계 일반 요건 통과 기업이 0개입니다.")
    
    print(f"\n{'='*60}")
    print(f"  ✅ 필터링 완료")
    print(f"     - Stage 1 (산업): {len(stage1_raw)}개")
    print(f"     - Stage 2 (재무): {len(stage2_profit)}개")
    print(f"     - Stage 3 (사업): {len(stage3_business)}개")
    print(f"     - Stage 4 (요건): {len(stage4_final)}개")
    print(f"{'='*60}\n")
    
    return {
        "stage1_raw": stage1_raw,
        "stage2_dec_passed": stage2_result.get("dec_passed", []),
        "stage2_profit_passed": stage2_profit,
        "stage3_business_passed": stage3_business,
        "stage4_final_peers": stage4_final[:10],  # 최대 10개로 제한
        "details": {
            "stage3_similarity": stage3_result.get("similarity_details", []),
            "stage4_requirements": stage4_result.get("details", [])
        }
    }