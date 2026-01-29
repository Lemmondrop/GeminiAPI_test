import os
import zipfile
import requests
import xml.etree.ElementTree as ET
import csv
from dotenv import load_dotenv

load_dotenv()
API_KEY = os.getenv("DART_API_KEY")
URL = "https://opendart.fss.or.kr/api/corpCode.xml"

CSV_ALL = "corp_code_all.csv"
CSV_LISTED = "corp_code_listed.csv"

# 1) ZIP 다운로드
resp = requests.get(URL, params={"crtfc_key": API_KEY}, timeout = 60)
resp.raise_for_status()

zip_path = "CorpCode.zip"
with open(zip_path, "wb") as f:
    f.write(resp.content)

# 2) ZPI 해제
extract_dir = "corpCode_extracted"
os.makedirs(extract_dir, exist_ok=True)

with zipfile.ZipFile(zip_path, "r") as zf:
    zf.extractall(extract_dir)

# 3) XML 파일 찾기 (zip 안에 1개인 경우가 일반적)
xml_files = [p for p in os.listdir(extract_dir) if p.lower().endswith("xml")]
if not xml_files:
    raise RuntimeError("압축 해제 후 XML 없음, 다운로드 파일 점검 필요.")

xml_path = os.path.join(extract_dir, xml_files[0])

# 4) XML 파싱 (corp_code / corp_name / stock_code / modify_date)
tree = ET.parse(xml_path)
root = tree.getroot()

rows = []
for item in root.findall("list"):
    rows.append({
        "corp_code": (item.findtext("corp_code") or "").strip(),
        "corp_name": (item.findtext("corp_name") or "").strip(),
        "stock_code": (item.findtext("stock_code") or "").strip(),
        "modify_date": (item.findtext("modify_date") or "").strip(),
    })

print("rows", len(rows))
print("sample:", rows[:5])


# -------------------------------------------------
# 4. corp_code → corp_name dict
# -------------------------------------------------
corp_code_to_name = {
    r["corp_code"]: r["corp_name"]
    for r in rows
    if r["corp_code"]
}

print("corp_code dict size:", len(corp_code_to_name))

# -------------------------------------------------
# 5. 상장사 필터링
# -------------------------------------------------
listed_rows = [r for r in rows if r["stock_code"]]

print("Listed companies:", len(listed_rows))

# -------------------------------------------------
# 6. CSV 저장
# -------------------------------------------------
def save_csv(path, data):
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(
            f,
            fieldnames=["corp_code", "corp_name", "stock_code", "modify_date"]
        )
        writer.writeheader()
        writer.writerows(data)

save_csv(CSV_ALL, rows)
save_csv(CSV_LISTED, listed_rows)

print("Saved:")
print("-", CSV_ALL)
print("-", CSV_LISTED)