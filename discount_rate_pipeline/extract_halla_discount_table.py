"""
Extract IPO discount-rate history tables from peer-group PDF files.

The original script was written for a single Halla Cast source.  This version
keeps the old function names for compatibility, but supports both:

  1. table extraction with pdfplumber, when it is installed
  2. text-line extraction with pypdf, which works well for the newer
     *_peer_group_extraction.pdf files

Output columns:
  listing_date, company_name, discount_rate_high_pct, discount_rate_low_pct,
  discount_rate_mid_pct, rate_col1, rate_col2, source_pdf, source_page,
  extraction_method
"""

import argparse
import csv
import re
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Sequence, Tuple


OUTPUT_FIELDS = [
    "listing_date",
    "company_name",
    "discount_rate_high_pct",
    "discount_rate_low_pct",
    "discount_rate_mid_pct",
    "rate_col1",
    "rate_col2",
    "source_pdf",
    "source_page",
    "extraction_method",
]

SUMMARY_NAMES = {
    "평균",
    "합계",
    "소계",
    "계",
    "구분",
    "회사명",
    "상장일",
}

DATE_RE = r"(?P<date>20\d{2}[-./]\d{2}[-./]\d{2})"
RATE_RE = r"(?P<r1>\d{1,2}(?:\.\d{1,2})?)\s*%\s+(?P<r2>\d{1,2}(?:\.\d{1,2})?)\s*%"

# Common layout in recent peer-group PDFs:
#   애드바이오텍 2022-01-24 43.90% 35.90%
# Company names may contain digits, e.g. M83.
COMPANY_FIRST_RE = re.compile(
    rf"(?P<company>[^\n\r%]{{1,50}}?)\s+{DATE_RE}\s+{RATE_RE}"
)

# Older/table layout:
#   2025-05-20 바이오비쥬 39.67% 31.42%
DATE_FIRST_RE = re.compile(
    rf"{DATE_RE}\s+(?P<company>[^\n\r%]{{1,50}}?)\s+{RATE_RE}"
)

SPLIT_DATE_PREFIX_RE = re.compile(r"^(?P<year>20\d{2})-(?P<month>\d{2})-\s*(?P<rest>.*)$")
SPLIT_DATE_DAY_RE = re.compile(r"^(?P<day>\d{2})$")
SPLIT_DATE_ROW_RE = re.compile(
    r"^(?P<company>[^\n\r%]{1,50}?)\s+"
    r"(?P<r1>\d{1,2}(?:\.\d{1,2})?)\s*%\s+"
    r"(?P<r2>\d{1,2}(?:\.\d{1,2})?)\s*%$"
)


def _clean_company_name(value: str) -> str:
    value = re.sub(r"\s+", " ", str(value or "")).strip()
    value = re.sub(r"^주\d+\)\s*", "", value)
    value = re.sub(r"^[|:ㆍ·\-\s]+|[|:ㆍ·\-\s]+$", "", value)
    value = re.sub(r"^(?:KOSPI|KOSDAQ|KONEX)\s+", "", value, flags=re.IGNORECASE)
    value = re.sub(r"\s+(?:KOSPI|KOSDAQ|KONEX)$", "", value, flags=re.IGNORECASE)
    return value.strip()


def _normalize_date(value: str) -> Optional[str]:
    value = str(value or "").replace(".", "-").replace("/", "-")
    m = re.fullmatch(r"(20\d{2})-(\d{2})-(\d{2})", value)
    if not m:
        return None

    year, month, day = map(int, m.groups())
    if not (2020 <= year <= 2030 and 1 <= month <= 12 and 1 <= day <= 31):
        return None
    return f"{year:04d}-{month:02d}-{day:02d}"


def _is_valid_company_name(value: str) -> bool:
    if not value or value in SUMMARY_NAMES:
        return False
    if len(value) < 2 or len(value) > 40:
        return False
    if re.fullmatch(r"[\d,\s.%-]+", value):
        return False
    if not re.search(r"[가-힣A-Za-z]", value):
        return False
    invalid_tokens = (
        "시장지수",
        "공모주식",
        "공모가",
        "희망공모",
        "확정공모",
        "청약",
        "수익률",
        "변동률",
        "제3자배정",
        "유상증자",
        "전환권",
        "구분",
    )
    if any(token in value for token in invalid_tokens):
        return False
    if re.search(r"\d{1,3},\d{3}", value):
        return False
    # Filter accidental prose fragments that can precede dates in non-table text.
    if any(token in value for token in ("출처", "주석", "기준주가", "분석기준일")):
        return False
    return True


def _valid_rates(rate1: float, rate2: float) -> bool:
    return 0.0 < rate1 < 90.0 and 0.0 < rate2 < 90.0


def _record(
    *,
    listing_date: str,
    company_name: str,
    rate1: float,
    rate2: float,
    source_pdf: str,
    source_page: int,
    extraction_method: str,
) -> Optional[Dict]:
    listing_date = _normalize_date(listing_date)
    company_name = _clean_company_name(company_name)

    if not listing_date or not _is_valid_company_name(company_name):
        return None
    if not _valid_rates(rate1, rate2):
        return None

    return {
        "listing_date": listing_date,
        "company_name": company_name,
        "rate_col1": round(float(rate1), 4),
        "rate_col2": round(float(rate2), 4),
        "source_pdf": source_pdf,
        "source_page": source_page,
        "extraction_method": extraction_method,
    }


def _records_from_text(text: str, source_pdf: str, source_page: int) -> List[Dict]:
    records: List[Dict] = []
    for pattern in (COMPANY_FIRST_RE, DATE_FIRST_RE):
        for match in pattern.finditer(text):
            record = _record(
                listing_date=match.group("date"),
                company_name=match.group("company"),
                rate1=float(match.group("r1")),
                rate2=float(match.group("r2")),
                source_pdf=source_pdf,
                source_page=source_page,
                extraction_method="pypdf_text",
            )
            if record:
                records.append(record)
    records.extend(_records_from_split_date_lines(text, source_pdf, source_page))
    return records


def _records_from_split_date_lines(text: str, source_pdf: str, source_page: int) -> List[Dict]:
    """Extract rows where PDFs split dates as '2025-04-' / row / '04'."""
    records: List[Dict] = []
    pending_prefix: Optional[Tuple[str, str]] = None
    pending_row: Optional[Tuple[str, str, str, str, str]] = None

    for raw_line in text.splitlines():
        line = re.sub(r"\s+", " ", raw_line or "").strip()
        if not line:
            continue

        prefix_match = SPLIT_DATE_PREFIX_RE.match(line)
        if prefix_match:
            pending_prefix = (prefix_match.group("year"), prefix_match.group("month"))
            pending_row = None
            rest = prefix_match.group("rest").strip()
            if rest:
                row_match = SPLIT_DATE_ROW_RE.match(rest)
                if row_match:
                    pending_row = (
                        pending_prefix[0],
                        pending_prefix[1],
                        row_match.group("company"),
                        row_match.group("r1"),
                        row_match.group("r2"),
                    )
            continue

        if pending_prefix and not pending_row:
            row_match = SPLIT_DATE_ROW_RE.match(line)
            if row_match:
                pending_row = (
                    pending_prefix[0],
                    pending_prefix[1],
                    row_match.group("company"),
                    row_match.group("r1"),
                    row_match.group("r2"),
                )
                continue

        if pending_row:
            day_match = SPLIT_DATE_DAY_RE.match(line)
            if day_match:
                year, month, company, r1, r2 = pending_row
                record = _record(
                    listing_date=f"{year}-{month}-{day_match.group('day')}",
                    company_name=company,
                    rate1=float(r1),
                    rate2=float(r2),
                    source_pdf=source_pdf,
                    source_page=source_page,
                    extraction_method="split_date_text",
                )
                if record:
                    records.append(record)
                pending_prefix = None
                pending_row = None

    return records


def extract_discount_table_pypdf(pdf_path: str) -> List[Dict]:
    """Extract discount rows from page text with pypdf."""
    try:
        from pypdf import PdfReader
    except ImportError:
        print("pypdf is not installed; skipping text extraction")
        return []

    path = Path(pdf_path)
    results: List[Dict] = []
    reader = PdfReader(str(path))
    for page_idx, page in enumerate(reader.pages, start=1):
        text = page.extract_text() or ""
        if "%" not in text or not re.search(r"20\d{2}[-./]\d{2}[-./](?:\d{2})?", text):
            continue
        results.extend(_records_from_text(text, path.name, page_idx))
    return results


def _parse_rates_from_cells(cells: Sequence[str]) -> List[float]:
    rates: List[float] = []
    for cell in cells:
        for match in re.finditer(r"([+-]?\d{1,2}(?:\.\d{1,2})?)\s*%?", str(cell or "")):
            try:
                value = float(match.group(1))
            except ValueError:
                continue
            if 0.0 < value < 90.0:
                rates.append(value)
    return rates


def _parse_table_row(row: Sequence[str], source_pdf: str, source_page: int) -> Optional[Dict]:
    cells = [str(cell or "").strip() for cell in row]
    if len(cells) < 3:
        return None

    # Layout A: date, company, rate, rate
    date_match = re.match(r"(20\d{2})[-./](\d{2})[-./](\d{2})", cells[0])
    if date_match:
        rates = _parse_rates_from_cells(cells[2:])
        if len(rates) >= 2:
            return _record(
                listing_date="-".join(date_match.groups()),
                company_name=cells[1],
                rate1=rates[0],
                rate2=rates[1],
                source_pdf=source_pdf,
                source_page=source_page,
                extraction_method="pdfplumber_table",
            )

    # Layout B: company, date, rate, rate
    row_text = " ".join(cells)
    for pattern in (COMPANY_FIRST_RE, DATE_FIRST_RE):
        match = pattern.search(row_text)
        if match:
            return _record(
                listing_date=match.group("date"),
                company_name=match.group("company"),
                rate1=float(match.group("r1")),
                rate2=float(match.group("r2")),
                source_pdf=source_pdf,
                source_page=source_page,
                extraction_method="pdfplumber_table",
            )

    return None


def extract_discount_table_pdfplumber(pdf_path: str) -> List[Dict]:
    """Extract discount rows with pdfplumber tables when available."""
    try:
        import pdfplumber
    except ImportError:
        return []

    path = Path(pdf_path)
    results: List[Dict] = []
    with pdfplumber.open(str(path)) as pdf:
        for page_idx, page in enumerate(pdf.pages, start=1):
            text = page.extract_text() or ""
            has_date = bool(re.search(r"20\d{2}[-./]\d{2}[-./](?:\d{2})?", text))
            has_rate = "%" in text
            if not (has_date and has_rate):
                continue

            for table in page.extract_tables() or []:
                for row in table or []:
                    record = _parse_table_row(row, path.name, page_idx)
                    if record:
                        results.append(record)

            # Some PDFs expose the table better as text than as table cells.
            results.extend(_records_from_text(text, path.name, page_idx))

    return results


def extract_discount_table(pdf_path: str, method: str = "auto") -> List[Dict]:
    """Extract discount-rate rows from one PDF."""
    if method not in {"auto", "pdfplumber", "pypdf"}:
        raise ValueError("method must be one of: auto, pdfplumber, pypdf")

    records: List[Dict] = []
    if method in {"auto", "pdfplumber"}:
        records.extend(extract_discount_table_pdfplumber(pdf_path))
    if method in {"auto", "pypdf"}:
        records.extend(extract_discount_table_pypdf(pdf_path))
    return records


def normalize_high_low(records: List[Dict]) -> List[Dict]:
    """Normalize two extracted rate columns into high/low/mid fields."""
    normalized: List[Dict] = []
    for record in records:
        out = dict(record)
        c1 = float(out["rate_col1"])
        c2 = float(out["rate_col2"])
        out["discount_rate_high_pct"] = round(max(c1, c2), 4)
        out["discount_rate_low_pct"] = round(min(c1, c2), 4)
        out["discount_rate_mid_pct"] = round((c1 + c2) / 2, 4)
        normalized.append(out)
    return normalized


def remove_summary_rows(records: List[Dict]) -> List[Dict]:
    """Remove summary/header rows and obvious non-company rows."""
    return [r for r in records if _is_valid_company_name(str(r.get("company_name", "")))]


def _dedupe_key(record: Dict, include_rates: bool = False) -> Tuple:
    key = (
        record.get("listing_date", ""),
        _clean_company_name(record.get("company_name", "")),
    )
    if include_rates:
        key += (
            round(float(record.get("discount_rate_high_pct", record.get("rate_col1", 0)) or 0), 2),
            round(float(record.get("discount_rate_low_pct", record.get("rate_col2", 0)) or 0), 2),
        )
    return key


def deduplicate(records: List[Dict], include_rates: bool = False) -> List[Dict]:
    """Deduplicate by listing date and company name, preserving first source."""
    seen = set()
    result: List[Dict] = []
    for record in records:
        key = _dedupe_key(record, include_rates=include_rates)
        if key in seen:
            continue
        seen.add(key)
        result.append(record)
    return result


KNOWN_REFERENCE_ROWS = [
    {"listing_date": "2025-05-20", "company_name": "바이오비쥬", "high": 39.67, "low": 31.42},
    {"listing_date": "2024-08-22", "company_name": "M83", "high": 43.84, "low": 33.63},
    {"listing_date": "2023-11-24", "company_name": "한선엔지니어링", "high": 32.99, "low": 22.68},
    {"listing_date": "2023-07-24", "company_name": "뷰티스킨", "high": 18.36, "low": 6.69},
    {"listing_date": "2023-01-19", "company_name": "한주라이트메탈", "high": 40.95, "low": 32.20},
]


def validate_extraction(records: List[Dict]) -> Dict:
    """Validate against a small known sample when the source includes it."""
    result = {"match": [], "mismatch": [], "missing": []}
    by_key = {(r["listing_date"], r["company_name"]): r for r in records}

    for ref in KNOWN_REFERENCE_ROWS:
        extracted = by_key.get((ref["listing_date"], ref["company_name"]))
        if extracted is None:
            result["missing"].append(ref["company_name"])
            continue

        high_match = abs(float(extracted["discount_rate_high_pct"]) - ref["high"]) < 0.1
        low_match = abs(float(extracted["discount_rate_low_pct"]) - ref["low"]) < 0.1
        if high_match and low_match:
            result["match"].append(ref["company_name"])
        else:
            result["mismatch"].append(
                {
                    "company": ref["company_name"],
                    "expected": (ref["high"], ref["low"]),
                    "extracted": (
                        extracted["discount_rate_high_pct"],
                        extracted["discount_rate_low_pct"],
                    ),
                }
            )

    result["accuracy_pct"] = (
        len(result["match"]) / len(KNOWN_REFERENCE_ROWS) * 100
        if KNOWN_REFERENCE_ROWS
        else 0
    )
    return result


def _iter_input_pdfs(pdf: Optional[str], input_dir: Optional[str], glob_pattern: str) -> List[Path]:
    paths: List[Path] = []
    if pdf:
        paths.append(Path(pdf))
    if input_dir:
        paths.extend(sorted(Path(input_dir).glob(glob_pattern)))

    unique: List[Path] = []
    seen = set()
    for path in paths:
        resolved = path.resolve()
        if resolved in seen:
            continue
        seen.add(resolved)
        unique.append(path)
    return unique


def write_csv(records: List[Dict], output: str) -> None:
    with open(output, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=OUTPUT_FIELDS, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(records)


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Extract IPO discount-rate tables from peer-group PDFs."
    )
    parser.add_argument("--pdf", help="single PDF path")
    parser.add_argument("--input-dir", help="directory containing peer-group PDFs")
    parser.add_argument("--glob", default="*_peer_group_extraction.pdf")
    parser.add_argument("--output", default="ipo_discount_rates.csv")
    parser.add_argument("--method", choices=["auto", "pdfplumber", "pypdf"], default="auto")
    parser.add_argument(
        "--dedupe-with-rates",
        action="store_true",
        help="dedupe by company/date/rates instead of company/date only",
    )
    parser.add_argument("--validate", action="store_true")
    args = parser.parse_args()

    pdfs = _iter_input_pdfs(args.pdf, args.input_dir, args.glob)
    if not pdfs:
        raise SystemExit("Provide --pdf or --input-dir")

    print("\n" + "=" * 60)
    print("Peer-group discount-rate extraction")
    print("=" * 60)

    raw_records: List[Dict] = []
    for pdf_path in pdfs:
        records = extract_discount_table(str(pdf_path), method=args.method)
        raw_records.extend(records)
        print(f"{pdf_path.name}: raw {len(records)} rows")

    cleaned = remove_summary_rows(raw_records)
    normalized = normalize_high_low(cleaned)
    deduped = deduplicate(normalized, include_rates=args.dedupe_with_rates)

    write_csv(deduped, args.output)

    print("\nSummary")
    print(f"  PDFs:        {len(pdfs)}")
    print(f"  Raw rows:    {len(raw_records)}")
    print(f"  Clean rows:  {len(cleaned)}")
    print(f"  Output rows: {len(deduped)}")
    print(f"  Saved:       {args.output}")

    if deduped:
        print("\nFirst 5 rows")
        for row in deduped[:5]:
            print(
                f"  {row['listing_date']} | {row['company_name']:20s} | "
                f"high {row['discount_rate_high_pct']:6.2f}% / "
                f"low {row['discount_rate_low_pct']:6.2f}% | "
                f"{row['source_pdf']} p.{row['source_page']}"
            )

    if args.validate:
        validation = validate_extraction(deduped)
        print("\nValidation sample")
        print(
            f"  match: {len(validation['match'])}/{len(KNOWN_REFERENCE_ROWS)} "
            f"({validation['accuracy_pct']:.0f}%)"
        )
        if validation["mismatch"]:
            print(f"  mismatch: {validation['mismatch']}")
        if validation["missing"]:
            print(f"  missing: {validation['missing']}")


if __name__ == "__main__":
    main()
