import re

from .normalizer import clean_choice_prefix, normalize_date, normalize_phone

WEEKDAYS = ["월요일", "화요일", "수요일", "목요일", "금요일"]
DAY_SHORT = {"월요일": "월", "화요일": "화", "수요일": "수", "목요일": "목", "금요일": "금"}


def _contains(text, token):
    return token in str(text or "")


def _first_idx(headers, token):
    for i, h in enumerate(headers):
        if _contains(h, token):
            return i
    return None


def _all_idx(headers, token):
    return [i for i, h in enumerate(headers) if _contains(h, token)]


def _first_value(row, indices):
    for idx in indices:
        v = row[idx]
        if v not in (None, ""):
            return v
    return None


def _norm_header(h):
    s = str(h or "").lower()
    s = re.sub(r"[\s_()\-.,]", "", s)
    return s


def _h_has(h, *tokens):
    hs = _norm_header(h)
    return all(t in hs for t in tokens)


def _parse_number(value):
    if value in (None, ""):
        return None
    m = re.search(r"\d+", str(value))
    if not m:
        return None
    try:
        return int(m.group(0))
    except Exception:
        return None


def _parse_day_segments(headers):
    day_method = {}
    day_time = {}
    day_vehicle2 = {}
    day_vehicle3 = {}
    day_loc2 = {}
    day_loc3 = {}

    day_anchors = []
    for day in WEEKDAYS:
        idx_method = None
        idx_time = None
        for i, h in enumerate(headers):
            hs = _norm_header(h)
            if day in str(h or "") and ("하교방법" in hs or ("하교" in hs and "방법" in hs)):
                idx_method = i
                break
        for i, h in enumerate(headers):
            hs = _norm_header(h)
            if day in str(h or "") and ("하교시간" in hs or ("하교" in hs and "시간" in hs)):
                idx_time = i
                break

        if idx_method is not None:
            day_method[day] = idx_method
        if idx_time is not None:
            day_time[day] = idx_time

        anchor = idx_method if idx_method is not None else idx_time
        if anchor is not None:
            day_anchors.append((day, anchor))

    day_anchors.sort(key=lambda x: x[1])
    for i, (day, start_idx) in enumerate(day_anchors):
        next_start = day_anchors[i + 1][1] if i + 1 < len(day_anchors) else len(headers)
        rng = list(range(start_idx, next_start))

        veh_candidates = []
        loc_candidates = []
        for c in rng:
            hs = _norm_header(headers[c])
            if "차량" in hs:
                veh_candidates.append(c)
            if ("장소" in hs) or ("하차" in hs):
                loc_candidates.append(c)

        v2 = None
        v3 = None
        for c in veh_candidates:
            hs = _norm_header(headers[c])
            if "2하교" in hs and v2 is None:
                v2 = c
            if "3하교" in hs and v3 is None:
                v3 = c

        # Fallback: if explicit 2/3 markers are missing, pick in-order vehicle columns.
        if v2 is None and veh_candidates:
            v2 = veh_candidates[0]
        if v3 is None and len(veh_candidates) >= 2:
            v3 = veh_candidates[1]
        elif v3 is None:
            v3 = v2

        day_vehicle2[day] = v2
        day_vehicle3[day] = v3

        if v2 is not None and v3 is not None and v2 <= v3:
            loc2 = [c for c in range(v2 + 1, v3) if c in loc_candidates]
            loc3 = [c for c in range(v3 + 1, next_start) if c in loc_candidates]
        else:
            loc2 = []
            loc3 = []

        # Fallback: use all location candidates in day range when segmented locations are not found.
        if not loc2:
            loc2 = loc_candidates[:]
        if not loc3:
            loc3 = loc_candidates[:]

        day_loc2[day] = loc2
        day_loc3[day] = loc3

    return day_method, day_time, day_vehicle2, day_vehicle3, day_loc2, day_loc3


def build_student_records(ws):
    headers = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]

    idx = {
        "name": _first_idx(headers, "학생이름"),
        "grade": _first_idx(headers, "학년"),
        "class": _first_idx(headers, "반"),
        "number": _first_idx(headers, "번호"),
        "birth": _first_idx(headers, "생년월일"),
        "address": _first_idx(headers, "주소(도로명주소)"),
        "mother_name": _first_idx(headers, "어머니 성명"),
        "mother_phone": _first_idx(headers, "어머니의 전화번호"),
        "father_name": _first_idx(headers, "아버지 성명"),
        "father_phone": _first_idx(headers, "아버지의 전화번호"),
        "siblings": _first_idx(headers, "형제가 있다면"),
        "boarding_method": _first_idx(headers, "(등교)_등교 방법"),
        "boarding_vehicle": _first_idx(headers, "(등교)_등교 탑승 차량"),
        "main_parent_phone": _first_idx(headers, "주 학부모전화번호"),
    }

    boarding_loc_indices = _all_idx(headers, "(등교)승차 장소를 선택해주세요")
    day_method, day_time, day_vehicle2, day_vehicle3, day_loc2, day_loc3 = _parse_day_segments(headers)

    records = []
    errors = []

    for r in range(2, ws.max_row + 1):
        row = [ws.cell(r, c).value for c in range(1, ws.max_column + 1)]
        name = row[idx["name"]] if idx["name"] is not None else None
        if name in (None, ""):
            continue

        number_raw = row[idx["number"]] if idx["number"] is not None else None
        number = _parse_number(number_raw)
        if number is None:
            errors.append({"row": r, "name": str(name), "field": "번호", "value": number_raw, "issue": "번호 파싱 실패"})

        birth_dt, birth_err = normalize_date(row[idx["birth"]] if idx["birth"] is not None else None)
        if birth_err:
            errors.append({"row": r, "name": str(name), "field": "생년월일", "value": row[idx["birth"]], "issue": birth_err})

        mother_phone, e1 = normalize_phone(row[idx["mother_phone"]] if idx["mother_phone"] is not None else None)
        if e1:
            errors.append({"row": r, "name": str(name), "field": "어머니전화", "value": row[idx["mother_phone"]], "issue": e1})

        father_phone, e2 = normalize_phone(row[idx["father_phone"]] if idx["father_phone"] is not None else None)
        if e2:
            errors.append({"row": r, "name": str(name), "field": "아버지전화", "value": row[idx["father_phone"]], "issue": e2})

        main_parent_phone, _ = normalize_phone(row[idx["main_parent_phone"]] if idx["main_parent_phone"] is not None else None)

        boarding_method = clean_choice_prefix(row[idx["boarding_method"]] if idx["boarding_method"] is not None else None)
        boarding_vehicle = clean_choice_prefix(row[idx["boarding_vehicle"]] if idx["boarding_vehicle"] is not None else None)
        boarding_loc = clean_choice_prefix(_first_value(row, boarding_loc_indices))

        dropoff = {}
        for day in WEEKDAYS:
            method = clean_choice_prefix(row[day_method[day]]) if day_method.get(day) is not None else ""
            time = clean_choice_prefix(row[day_time[day]]) if day_time.get(day) is not None else ""
            vehicle = ""
            location = ""

            if method == "학교차량이용":
                if "2하교" in str(time):
                    v_idx = day_vehicle2.get(day)
                    vehicle = clean_choice_prefix(row[v_idx]) if v_idx is not None else ""
                    location = clean_choice_prefix(_first_value(row, day_loc2.get(day, [])))
                elif "3하교" in str(time):
                    v_idx = day_vehicle3.get(day)
                    vehicle = clean_choice_prefix(row[v_idx]) if v_idx is not None else ""
                    location = clean_choice_prefix(_first_value(row, day_loc3.get(day, [])))

            dropoff[day] = {"method": method, "time": time, "vehicle": vehicle, "location": location}

        grade = str(row[idx["grade"]] if idx["grade"] is not None else "")
        class_ = str(row[idx["class"]] if idx["class"] is not None else "")

        records.append(
            {
                "row": r,
                "name": str(name).strip(),
                "grade_text": grade,
                "class_text": class_,
                "grade_num": int("".join(ch for ch in grade if ch.isdigit()) or 0),
                "class_num": int("".join(ch for ch in class_ if ch.isdigit()) or 0),
                "number": number,
                "birth_date": birth_dt,
                "birth_raw": row[idx["birth"]] if idx["birth"] is not None else "",
                "address": str(row[idx["address"]] or "").strip() if idx["address"] is not None else "",
                "address_raw": row[idx["address"]] if idx["address"] is not None else "",
                "mother_name": str(row[idx["mother_name"]] or "").strip() if idx["mother_name"] is not None else "",
                "mother_phone": mother_phone,
                "father_name": str(row[idx["father_name"]] or "").strip() if idx["father_name"] is not None else "",
                "father_phone": father_phone,
                "siblings": str(row[idx["siblings"]] or "").strip() if idx["siblings"] is not None else "",
                "boarding_method": boarding_method,
                "boarding_vehicle": boarding_vehicle,
                "boarding_location": boarding_loc,
                "main_parent_phone": main_parent_phone,
                "dropoff": dropoff,
            }
        )

        addr = str(row[idx["address"]] or "").strip() if idx["address"] is not None else ""
        if addr and ("구" not in addr):
            errors.append(
                {
                    "row": r,
                    "name": str(name),
                    "field": "주소",
                    "value": addr,
                    "issue": "구(區) 정보 누락 의심",
                }
            )

    return records, errors


def make_student_id(grade_num, class_num, number):
    if not number:
        return ""
    if grade_num and class_num:
        return f"{grade_num}{class_num}{int(number):02d}"
    return f"{int(number):02d}"
