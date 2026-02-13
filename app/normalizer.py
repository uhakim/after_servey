import datetime as dt
import re


def clean_choice_prefix(value):
    if value in (None, ""):
        return ""
    s = str(value).strip()
    s = re.sub(r"^\d+\s*[.)]\s*", "", s)
    s = s.replace("베내시티", "베네시티")
    return s.strip()


def normalize_date(value):
    if value in (None, ""):
        return None, "생년월일 미입력"

    s = str(value).strip()
    digits = re.sub(r"\D", "", s)

    try:
        # Korean-style date strings: 15년1월19일, 2015년 1월 19일
        m = re.match(r"^\s*(\d{2,4})\D+(\d{1,2})\D+(\d{1,2})\D*$", s)
        if m:
            y = int(m.group(1))
            if y < 100:
                y += 2000
            mo = int(m.group(2))
            d = int(m.group(3))
            return dt.date(y, mo, d), None

        if len(digits) == 6:  # YYMMDD
            y = int(digits[:2]) + 2000
            m = int(digits[2:4])
            d = int(digits[4:6])
            return dt.date(y, m, d), None
        if len(digits) == 8:  # YYYYMMDD
            y = int(digits[:4])
            m = int(digits[4:6])
            d = int(digits[6:8])
            return dt.date(y, m, d), None

        # formatted date strings
        for fmt in ("%Y-%m-%d", "%Y.%m.%d", "%Y/%m/%d"):
            try:
                return dt.datetime.strptime(s, fmt).date(), None
            except ValueError:
                pass
    except Exception:
        return None, "생년월일 파싱 실패"

    return None, "생년월일 형식 불일치"


def normalize_phone(value):
    if value in (None, ""):
        return "", "전화번호 미입력"

    s = str(value).strip()
    nums = re.sub(r"\D", "", s)

    if nums.startswith("010"):
        if len(nums) == 11:
            return f"{nums[:3]}-{nums[3:7]}-{nums[7:]}", None
        if len(nums) == 10:
            return f"{nums[:3]}-{nums[3:6]}-{nums[6:]}", None
        return s, "휴대전화 길이 이상"

    if nums.startswith("02"):
        if len(nums) == 10:
            return f"{nums[:2]}-{nums[2:6]}-{nums[6:]}", None
        if len(nums) == 9:
            return f"{nums[:2]}-{nums[2:5]}-{nums[5:]}", None

    if len(nums) in (10, 11):
        if len(nums) == 11:
            return f"{nums[:3]}-{nums[3:7]}-{nums[7:]}", None
        return f"{nums[:3]}-{nums[3:6]}-{nums[6:]}", None

    return s, "전화번호 형식 불일치"
