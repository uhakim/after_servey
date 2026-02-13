from io import BytesIO
from datetime import date
import os

import openpyxl
import streamlit as st

from app.address_api import resolve_addresses
from app.mapper import build_student_records
from app.validator import build_validation_frame
from app.builders import (
    build_boarding_report,
    build_dropoff_result,
    build_student_roster,
)

DEFAULT_ROSTER_TEMPLATE = "template_roster.xlsx"
DEFAULT_VEHICLE_TEMPLATE = "template_dropoff.xlsx"


def _safe_secret(key, default=""):
    try:
        return st.secrets.get(key, default)
    except Exception:
        return default


st.set_page_config(page_title="등하교 설문 변환기", layout="wide")
st.title("등하교 설문 변환기")
st.caption("설문 엑셀 업로드 -> 산출물 엑셀 다운로드")

if "result_bundle" not in st.session_state:
    st.session_state.result_bundle = None

col1, col2 = st.columns(2)
with col1:
    survey_file = st.file_uploader("설문 결과 파일(.xlsx)", type=["xlsx"], key="survey")
with col2:
    school_year = st.number_input("학년도", min_value=2020, max_value=2100, value=date.today().year)

run = st.button("변환 실행", type="primary", use_container_width=True)

if run:
    if not survey_file:
        st.error("설문 파일을 먼저 업로드하세요.")
        st.stop()

    try:
        wb = openpyxl.load_workbook(BytesIO(survey_file.read()), data_only=True)
        if "학생" not in wb.sheetnames:
            st.error("설문 파일에 '학생' 시트가 없습니다.")
            st.stop()

        records, errors = build_student_records(wb["학생"])
        if not records:
            st.error("학생 데이터를 읽지 못했습니다.")
            st.stop()

        juso_key = os.getenv("JUSO_API_KEY") or _safe_secret("JUSO_API_KEY", "")
        api_ok = 0
        api_fail = 0
        if juso_key:
            api_ok, api_fail, _ = resolve_addresses(records, juso_key, timeout_sec=3)

        df_err = build_validation_frame(errors)

        out1 = build_student_roster(
            records,
            school_year=school_year,
            template_bytes=None,
            default_template_path=DEFAULT_ROSTER_TEMPLATE,
        )

        out2 = build_dropoff_result(
            records,
            template_bytes=None,
            default_template_path=DEFAULT_VEHICLE_TEMPLATE,
        )

        out3 = build_boarding_report(records)

        st.session_state.result_bundle = {
            "school_year": school_year,
            "student_count": len(records),
            "df_err": df_err,
            "out1": out1,
            "out2": out2,
            "out3": out3,
            "api_ok": api_ok,
            "api_fail": api_fail,
            "api_on": bool(juso_key),
        }

    except Exception as e:
        st.exception(e)

bundle = st.session_state.result_bundle
if bundle:
    st.subheader("검증 결과")
    st.write(f"학생 수: {bundle['student_count']}명")
    st.write(f"경고 수: {len(bundle['df_err'])}건")
    if bundle.get("api_on"):
        st.write(f"주소 API 적용: 성공 {bundle.get('api_ok', 0)}건 / 실패 {bundle.get('api_fail', 0)}건")
    else:
        st.info("주소 API 키가 없어 규칙 기반 주소 정규화만 적용했습니다.")
    if not bundle["df_err"].empty:
        st.dataframe(bundle["df_err"], use_container_width=True, height=240)

    st.subheader("다운로드")
    st.download_button(
        "1) 학생일람표 다운로드",
        data=bundle["out1"],
        file_name=f"{bundle['school_year']}학년도_학생일람표_자동생성.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="dl1",
    )
    st.download_button(
        "2) 하교차량조사결과 다운로드",
        data=bundle["out2"],
        file_name=f"{bundle['school_year']}학년도_하교차량조사결과_자동생성.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="dl2",
    )
    st.download_button(
        "3) 등교차량 조사(개선형) 다운로드",
        data=bundle["out3"],
        file_name=f"{bundle['school_year']}학년도_등교차량조사_개선형.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="dl3",
    )

    if not bundle["df_err"].empty:
        st.download_button(
            "검증 로그(csv) 다운로드",
            data=bundle["df_err"].to_csv(index=False).encode("utf-8-sig"),
            file_name="검증로그.csv",
            mime="text/csv",
            key="dl_log",
        )

    st.success("변환 결과가 준비되었습니다. 다운로드 버튼을 순서대로 눌러도 화면이 유지됩니다.")
