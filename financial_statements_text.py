import os
import json
import requests
from dotenv import load_dotenv

load_dotenv()
DART_API_KEY = (os.getenv("DART_API_KEY") or "").strip()
BASE_JSON = "https://opendart.fss.or.kr/api"

def call_opendart(endpoint: str, params: dict, timeout=30):
    url = f"{BASE_JSON}/{endpoint}"
    r = requests.get(url, params=params, timeout=timeout)
    r.raise_for_status()
    return r

def save_raw_json(obj, path: str):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(obj, f, ensure_ascii=False, indent=2)

if __name__ == "__main__":
    corp_code = "01448544"
    bsns_year = "2024"
    reprt_code = "11011"

    params = {
        "crtfc_key": DART_API_KEY,
        "corp_code": corp_code,
        "bsns_year": bsns_year,
        "reprt_code": reprt_code,
    }

    r = call_opendart("fnlttSinglAcnt.json", params)

    print("HTTP:", r.status_code)
    print("Final URL:", r.url)
    print("Content-Type:", r.headers.get("Content-Type"))

    # --- JSON 파싱 (실패해도 raw text 보존) ---
    try:
        data = r.json()
        print("JSON_KEYS:", list(data.keys()))
        print("status/message:", data.get("status"), data.get("message"))
    except Exception as e:
        print("JSON parse failed:", repr(e))
        data = {
            "_raw_text": r.text,
            "_parse_error": repr(e),
        }

    # --- raw json 저장 ---
    out_path = (
        f"./debug_dump/"
        f"fnlttSinglAcnt_{corp_code}_{bsns_year}_{reprt_code}.json"
    )
    save_raw_json(data, out_path)
    print("saved:", out_path)
