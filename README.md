# 미디어팀 보고서 자동화 (Streamlit)

주간 텍스트를 월 보고서 Word 양식에 넣고, 월 보고서 3개를 사용자가 제공한 **HWPX 분기 템플릿의 포맷을 유지한 채** 분기 보고서로 만드는 앱입니다.

## 핵심 원칙
- 입력에 없는 사역 내용은 생성하지 않습니다.
- 원문의 의미를 바꾸는 요약/재작성은 하지 않습니다.
- 프로그램이 새로 만든 분류명(현재 `[기타]`)은 파란색으로 표시합니다.
- 월 보고서는 기존 Word 양식을 템플릿으로 사용합니다.
- 분기 HWPX는 기존 표/병합/글꼴/여백 스타일을 유지하고 셀 내용과 필요한 날짜 행 수만 바꿉니다.
- 분기 규칙: 1분기 12·1·2월 / 2분기 3·4·5월 / 3분기 6·7·8월 / 4분기 9·10·11월.

## 로컬 실행
```bash
python -m venv .venv
# Windows: .venv\\Scripts\\activate
# macOS/Linux: source .venv/bin/activate
pip install -r requirements.txt
streamlit run app.py
```

## Streamlit Community Cloud
1. 이 폴더를 GitHub 저장소에 업로드합니다.
2. Streamlit Community Cloud에서 저장소를 선택합니다.
3. Main file path를 `app.py`로 지정합니다.

## Render 배포
- Build Command: `pip install -r requirements.txt`
- Start Command: `streamlit run app.py --server.port $PORT --server.address 0.0.0.0`

## 폴더 구조
- `app.py`: Streamlit 화면
- `report_engine.py`: 텍스트 파싱, DOCX/HWPX 생성 로직
- `templates/monthly_template.docx`: 월 보고서 Word 양식
- `templates/quarter_template.hwpx`: 사용자 제공 분기 보고서 HWPX 양식
- `examples/sample_weekly.txt`: 주간 입력 예시

## HWPX 포맷 보존 방식
HWPX는 ZIP/XML 형식입니다. 이 앱은 새 문서를 그리지 않고 제공된 `quarter_template.hwpx`를 복제한 뒤 `Contents/section0.xml` 안의 기존 월별 표 셀 텍스트를 교체합니다. 월별 일요일 수가 4회/5회로 달라지면 기존 데이터 행 스타일을 복제하거나 제거합니다.

내용이 기존 셀보다 매우 길 경우 한글 프로그램의 자동 줄바꿈 때문에 행 높이/페이지 나눔은 달라질 수 있습니다.


## 외부사역 중복 처리
분기 보고서 생성 시 월 보고서의 주차별 외부사역에서 **완전히 동일한 문구는 최초 1회만 유지**합니다. 의미가 비슷하지만 문구가 다른 항목은 자동 병합하지 않습니다.
