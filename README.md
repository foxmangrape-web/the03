# 파일 가계부 정리기

Python Flask로 만든 학생 프로젝트입니다. DB 없이 Excel을 분석해 같은 화면에 합계, 월별 PNG 그래프, 카테고리 지출 표를 보여주고 요약 Excel을 만듭니다. JavaScript는 사용하지 않습니다.

## 실행

Python 3.12 환경에서 실행.bat를 더블클릭하세요. 처음에는 패키지 설치에 인터넷이 필요합니다. 직접 실행하려면:

```sh
python -m venv .venv
.venv\Scripts\python -m pip install -r requirements.txt
.venv\Scripts\python app.py
```

브라우저에서 http://127.0.0.1:5000 으로 접속합니다. HTML을 직접 열지 말고 Python 서버를 실행합니다.

## 파일 구성

- app.py: 파일 읽기, 계산, 그래프 및 Excel 생성, Flask 서버
- templates/index.html: 업로드와 결과를 함께 보여주는 유일한 화면
- requirements.txt: Python 패키지 목록
- render.yaml: Render Web Service 배포 설정
- 실행.bat: Windows 실행 도우미

## Render 배포

### Python 버전 설정 필수

이 프로젝트의 Matplotlib 3.9.2는 Render 기본 Python 3.14.10에서 그래프 생성 시 RecursionError가 발생할 수 있습니다. **CSV로 바꿀 문제가 아니라 실행 버전 문제입니다.**

1. Render 서비스 → Environment에서 `PYTHON_VERSION`을 `3.12.8`로 추가하거나 수정합니다.
2. 저장 후 Manual Deploy → Clear build cache & deploy를 실행합니다.
3. 빌드 로그에서 Python 3.12.8을 사용하는지 확인합니다.

저장소 루트에 `.python-version`도 포함했습니다. Environment의 PYTHON_VERSION이 파일보다 우선합니다. 기존 서비스를 수동으로 만든 경우 render.yaml을 올리는 것만으로 그 서비스의 환경변수가 변경되지는 않습니다.

이 파일들을 GitHub 저장소 루트에 올리고 **New → Web Service**로 저장소를 연결합니다.

| 설정 | 값 |
| --- | --- |
| Language | Python 3 |
| Root Directory | 비워두기 |
| Build Command | pip install -r requirements.txt |
| Start Command | gunicorn app:app --bind 0.0.0.0:$PORT --timeout 120 |
| Environment Variable | PYTHON_VERSION = 3.12.8 |

**Publish Directory가 보이면 Static Site를 선택한 것입니다. Web Service로 다시 생성하세요.** Flask에는 Publish Directory를 입력하지 않습니다. New → Blueprint에서 render.yaml로 배포할 수도 있습니다.

공식 안내: https://render.com/docs/deploy-flask

## 입력 양식

| 날짜 | 구분 | 카테고리 | 금액 | 메모 |
| --- | --- | --- | ---: | --- |
| 2026-11-01 | 수입 | 용돈 | 180000 | 월초 용돈 |
| 2026-11-02 | 지출 | 식비 | 12000 | 점심 |

화면에서 양식을 내려받을 수 있습니다. XLSX와 XLS의 첫 번째 시트를 읽으며 최대 8MB, 20,000행을 지원합니다. 날짜는 엑셀 날짜 또는 YYYY-MM-DD, 구분은 수입/지출, 금액은 0 이상의 정수입니다. 메모는 선택이고 빈 카테고리는 기타로 표시합니다. 수식은 계산 후 저장된 값을 읽습니다. 암호 파일은 지원하지 않습니다.

원본은 저장하지 않습니다. 결과 Excel만 서버 임시 폴더에 보관하고 1시간 후 다음 분석/다운로드 요청 때 정리합니다. 서버 재시작 시 결과가 사라질 수 있습니다. 단일 서버용이며 결과 URL을 아는 사람은 파일을 내려받을 수 있습니다.

## 검증

```sh
python -m unittest discover -s tests
```
