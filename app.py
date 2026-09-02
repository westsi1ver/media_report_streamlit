from pathlib import Path
import re
import streamlit as st
from report_engine import (
    build_monthly_docx,
    build_quarterly_docx,
    build_quarterly_hwpx,
    parse_weekly_text,
    quarter_for_month,
)

BASE = Path(__file__).parent
MONTH_TEMPLATE = BASE / 'templates' / 'monthly_template.docx'
QUARTER_HWPX_TEMPLATE = BASE / 'templates' / 'quarter_template.hwpx'

st.set_page_config(page_title='미디어팀 보고서 자동화', page_icon='🎨', layout='wide')
st.title('🎨 미디어팀 보고서 자동화')
st.caption('주간 텍스트 → 월 보고서 DOCX → 분기 보고서 HWPX')

monthly_tab, quarter_tab = st.tabs(['주간 → 월 보고서', '월 보고서 → 분기 보고서'])

with monthly_tab:
    c1, c2 = st.columns(2)
    with c1:
        year = st.number_input('연도', min_value=2020, max_value=2100, value=2026, step=1)
    with c2:
        month = st.selectbox('보고 월', list(range(1, 13)), index=7, format_func=lambda x: f'{x}월')
    text = st.text_area('주간 사역 내용을 그대로 붙여넣으세요', height=420, placeholder='🎨미디어팀 08/09\n\n● 온라인 주보 ...')
    st.info('없는 내용은 생성하지 않습니다. 프로그램이 새 분류명을 만든 경우(예: [기타])에만 파란색으로 표시합니다.')
    if text:
        parsed = [x for x in parse_weekly_text(text) if x.month == month]
        if parsed:
            st.write('인식된 주차:', ', '.join(f'{x.month:02d}/{x.day:02d} → {x.week_no}주' for x in parsed))
        else:
            st.warning('선택한 월과 일치하는 “미디어팀 MM/DD” 블록을 찾지 못했습니다.')
    if st.button('월 보고서 만들기', type='primary', disabled=not text):
        try:
            data = build_monthly_docx(MONTH_TEMPLATE, text, int(year), int(month))
            st.success('월 보고서를 만들었습니다.')
            st.download_button('📥 Word 월 보고서 다운로드', data=data,
                               file_name=f'[미디어팀] {month}월 보고서.docx',
                               mime='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
        except Exception as e:
            st.exception(e)

with quarter_tab:
    st.write('분기 규칙: **1분기=12·1·2월 / 2분기=3·4·5월 / 3분기=6·7·8월 / 4분기=9·10·11월**')
    st.caption('분기 HWPX는 업로드하신 원본 HWPX의 표·셀·글꼴·여백 스타일을 유지하고 내용만 교체합니다.')
    qyear = st.number_input('분기 보고서 기준 연도', min_value=2020, max_value=2100, value=2026, step=1, key='qyear')
    uploads = st.file_uploader('월 보고서 DOCX 3개 업로드', type=['docx'], accept_multiple_files=True)
    month_hints = []
    if uploads:
        for up in uploads:
            m = re.search(r'(\d{1,2})월', up.name)
            month_hints.append(int(m.group(1)) if m else 0)
        st.caption('파일명에서 인식된 월: ' + ', '.join(str(x) if x else '?' for x in month_hints))
    if st.button('분기 보고서 만들기', type='primary', disabled=len(uploads or []) == 0):
        if len(uploads) != 3:
            st.error('분기 보고서는 월 보고서 3개를 업로드해 주세요.')
        elif any(m == 0 for m in month_hints):
            st.error('파일명에 “6월”, “7월”처럼 월 정보가 있어야 합니다.')
        else:
            qs = {quarter_for_month(m)[0] for m in month_hints}
            if len(qs) != 1:
                st.error('서로 같은 분기에 속한 3개 월 보고서를 업로드해 주세요.')
            else:
                q = next(iter(qs))
                docs = [(up.getvalue(), int(qyear) - 1 if q == 1 and month == 12 else int(qyear), month) for up, month in zip(uploads, month_hints)]
                try:
                    out = build_quarterly_hwpx(QUARTER_HWPX_TEMPLATE, docs)
                    st.success('원본 HWPX 포맷을 기반으로 분기 보고서를 만들었습니다.')
                    st.download_button('📥 한글 HWPX 분기 보고서 다운로드', data=out,
                                       file_name=f'[미디어팀] {q}분기 보고서.hwpx',
                                       mime='application/vnd.hancom.hwpx')
                    with st.expander('DOCX 분기본도 함께 필요할 때'):
                        docx_out = build_quarterly_docx(docs)
                        st.download_button('📥 Word 분기 보고서 다운로드', data=docx_out,
                                           file_name=f'[미디어팀] {q}분기 보고서.docx',
                                           mime='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
                except Exception as e:
                    st.exception(e)
