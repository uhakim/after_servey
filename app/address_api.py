import json
import re
import urllib.parse
import urllib.request
from functools import lru_cache


API_URL = "https://business.juso.go.kr/addrlink/addrLinkApi.do"


@lru_cache(maxsize=2048)
def _lookup(keyword, confm_key, timeout_sec):
    params = {
        "currentPage": 1,
        "countPerPage": 5,
        "keyword": keyword,
        "confmKey": confm_key,
        "resultType": "json",
    }
    url = f"{API_URL}?{urllib.parse.urlencode(params)}"
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    with urllib.request.urlopen(req, timeout=timeout_sec) as resp:
        data = json.loads(resp.read().decode("utf-8", errors="replace"))

    results = data.get("results", {})
    common = results.get("common", {})
    if common.get("errorCode") != "0":
        return None, common.get("errorMessage") or "주소 API 오류"

    juso_list = results.get("juso", []) or []
    if not juso_list:
        return None, "검색 결과 없음"

    best = juso_list[0]
    road_addr = best.get("roadAddr", "").strip()
    if not road_addr:
        return None, "표준 도로명주소 없음"
    return road_addr, None


def _extract_detail_tokens(raw):
    s = str(raw or "")
    patterns = [
        r"\d{1,4}동\s*\d{1,5}호",   # 101동2604호
        r"\d{1,4}동",                # 101동
        r"\d{1,5}호",                # 2604호
        r"\d{1,2}층",                # 3층
        r"\bB\d{1,2}\b",           # B1
        r"\b[A-Za-z]-\d{3,5}\b",   # A-3501
        r"\b\d{2,4}-\d{3,5}\b",   # 102-5909 (동-호 축약)
    ]

    tokens = []
    for p in patterns:
        for m in re.finditer(p, s, flags=re.IGNORECASE):
            t = re.sub(r"\s+", "", m.group(0)).upper()
            if t not in tokens:
                tokens.append(t)

    # Remove short tokens already covered by longer combined token.
    filtered = []
    for t in tokens:
        covered = False
        for other in tokens:
            if t != other and t in other and (t.endswith("동") or t.endswith("호")):
                covered = True
                break
        if not covered:
            filtered.append(t)
    return filtered


def _merge_with_detail(road_addr, raw_addr):
    details = _extract_detail_tokens(raw_addr)
    if not details:
        return road_addr

    merged_details = []
    road_norm = re.sub(r"\s+", "", str(road_addr or "")).upper()
    for d in details:
        if d not in road_norm and d not in merged_details:
            merged_details.append(d)

    if not merged_details:
        return road_addr
    return f"{road_addr}, {' '.join(merged_details)}"


def resolve_addresses(records, confm_key, timeout_sec=3):
    if not confm_key:
        return 0, 0, []

    success = 0
    failed = 0
    issues = []

    for rec in records:
        raw = str(rec.get("address_raw", rec.get("address", "")) or "").strip()
        if not raw:
            continue
        try:
            road_addr, err = _lookup(raw, confm_key, int(timeout_sec))
            if road_addr:
                rec["address_api"] = _merge_with_detail(road_addr, raw)
                success += 1
            else:
                failed += 1
                issues.append({"name": rec.get("name", ""), "address": raw, "issue": err or "조회 실패"})
        except Exception as e:
            failed += 1
            issues.append({"name": rec.get("name", ""), "address": raw, "issue": f"API 예외: {e}"})

    return success, failed, issues