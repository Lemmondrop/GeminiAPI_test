"""KSIC(한국표준산업분류) 코드 파싱·정규화.

두 가지 코드 포맷을 모두 지원한다:
  - 유니버스 CSV 패딩형 'C00259' = [대분류][00][소분류3] → 소분류 = 뒤 3자리(259)
  - 신고서 풀 KSIC 'C26429'    = [대분류][중분류2][소분류이하] → 소분류 = 앞 3자리(264)
구분: 대분류 뒤 숫자가 '00'으로 시작하면 패딩형. 둘 다 소분류 3자리로 정규화돼 상호 매칭된다.
스타트업 자가입력 시드 'C00303' / '30320' / '303' 등 혼재 형식도 정규화한다.
"""
from __future__ import annotations


def parse_ksic(code: str) -> tuple[str, str, str]:
    """'C00259'->('C','25','259'), 'C26429'->('C','26','264')  (대분류, 중분류, 소분류)."""
    s = str(code).strip().upper()
    maj = s[0] if s[:1].isalpha() else ""
    digits = "".join(ch for ch in (s[1:] if maj else s) if ch.isdigit())
    sub3 = digits[2:5] if digits[:2] == "00" else digits[:3]   # 패딩형 뒤3 / 풀형 앞3
    return maj, sub3[:2], sub3


def normalize_seed(seed: str) -> tuple[str | None, str, str]:
    """스타트업 자가입력 코드 정규화. 'C00303' / '30320' / '303' 모두 허용.

    반환: (대분류 or None, 중분류 2자리, 소분류 3자리)
    숫자만 입력된 경우 대분류는 미상(None)이 될 수 있다(임베딩 레인이 강건하게 보완).
    """
    s = str(seed).strip()
    if s and s[0].isalpha():                 # 'C00303' / 'C26429'
        return parse_ksic(s)
    digits = "".join(ch for ch in s if ch.isdigit())
    sub3 = digits[:3] if len(digits) >= 3 else digits.zfill(3)  # '30320'->'303', '303'->'303'
    return None, sub3[:2], sub3              # 대분류 미상 가능