import requests
from bs4 import BeautifulSoup

UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36"

def test_ctb203_wisereport(code: str):
    url = f"https://navercomp.wisereport.co.kr/v2/company/c1020001.aspx?cmp_cd={code}&cn="

    headers = {
        "User-Agent": UA,
        # 보통 referer 없어도 되는데, 막히면 아래 둘 중 하나로 바꿔가며 테스트
        "Referer": f"https://finance.naver.com/item/coinfo.naver?code={code}",
        # "Referer": "https://finance.naver.com/",
        "Accept-Language": "ko-KR,ko;q=0.9,en-US;q=0.8,en;q=0.7",
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    }

    r = requests.get(url, headers=headers, timeout=10, allow_redirects=True)

    # 와이즈리포트는 대체로 utf-8이지만, 혹시 모르면 자동/강제 보정
    if not r.encoding or r.encoding.lower() in ("iso-8859-1", "latin-1"):
        r.encoding = "utf-8"

    html = r.text
    soup = BeautifulSoup(html, "html.parser")

    print("URL       :", url)
    print("Status    :", r.status_code)
    print("Final URL :", r.url)
    print("len(html) :", len(html))
    print("has cTB203(str) :", 'id="cTB203"' in html)
    print("has cTB203(bs4) :", soup.find("table", id="cTB203") is not None)

    table = soup.find("table", id="cTB203")
    if table:
        rows = table.select("tbody tr")
        print("rows:", len(rows))
        # 샘플 5개 출력
        for tr in rows[:5]:
            th = tr.select_one("th span.cut") or tr.select_one("th")
            td = tr.select_one("td")
            if th and td:
                name = th.get_text(" ", strip=True).replace("\xa0", "")
                ratio = td.get_text(" ", strip=True).replace("\xa0", "")
                if name and ratio:
                    print("-", name, ratio)

if __name__ == "__main__":
    test_ctb203_wisereport("006800")
