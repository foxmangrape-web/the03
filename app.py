"""Excel 가계부를 분석하는 한 화면 Flask 앱."""
import base64
import io
import os
import time
import uuid
from collections import defaultdict
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path
from tempfile import gettempdir

os.environ.setdefault('MPLCONFIGDIR', str(Path(gettempdir()) / 'ledger-matplotlib'))

import xlrd
from flask import Flask, abort, render_template, request, send_file
import matplotlib

matplotlib.use('Agg')

from matplotlib.backends.backend_agg import FigureCanvasAgg
from matplotlib.figure import Figure
from openpyxl import Workbook, load_workbook

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 8 * 1024 * 1024
RESULT_DIR = Path(gettempdir()) / 'flask-ledger-results'
RESULT_DIR.mkdir(exist_ok=True)


def read_rows(file):
    """xlsx와 xls의 첫 번째 시트를 같은 행 목록으로 변환합니다."""
    # Flask의 업로드 임시 스트림 대신, Python 버전에 관계없이 탐색 가능한
    # 메모리 버퍼를 Excel 라이브러리에 전달합니다.
    content = file.read()
    if Path(file.filename).suffix.lower() == '.xls':
        book = xlrd.open_workbook(file_contents=content, on_demand=True)
        try:
            sheet = book.sheet_by_index(0)
            rows = []
            for index in range(sheet.nrows):
                values = []
                for cell in sheet.row(index):
                    value = cell.value
                    if cell.ctype == xlrd.XL_CELL_DATE:
                        value = xlrd.xldate_as_datetime(value, book.datemode)
                    values.append(value)
                rows.append(values)
                if len(rows) > 20001:
                    raise ValueError('한 번에 20,000행까지 분석할 수 있습니다.')
            return rows
        finally:
            book.release_resources()
    book = load_workbook(io.BytesIO(content), read_only=True, data_only=True)
    try:
        rows = []
        for row in book.worksheets[0].iter_rows(values_only=True):
            rows.append(list(row))
            if len(rows) > 20001:
                raise ValueError('한 번에 20,000행까지 분석할 수 있습니다.')
        return rows
    finally:
        book.close()


def analyze(rows):
    if not rows:
        raise ValueError('비어 있는 파일입니다.')
    headers = [str(value or '').strip() for value in rows[0]]
    required = ['날짜', '구분', '카테고리', '금액']
    if any(name not in headers for name in required):
        raise ValueError('첫 번째 행에 날짜, 구분, 카테고리, 금액 열을 넣어주세요.')
    positions = {name: headers.index(name) for name in required}
    months = defaultdict(lambda: {'income': 0, 'expense': 0})
    categories = defaultdict(int)
    count = 0
    for number, row in enumerate(rows[1:], 2):
        if not any(value is not None and str(value).strip() for value in row):
            continue
        values = {name: row[index] if index < len(row) else None for name, index in positions.items()}
        day = values['날짜']
        try:
            if not isinstance(day, (date, datetime)):
                day = date.fromisoformat(str(day).strip())
        except ValueError:
            raise ValueError(f'{number}행: 날짜를 2026-11-01 형식으로 입력해주세요.') from None
        kind = str(values['구분'] or '').strip()
        if kind not in ('수입', '지출'):
            raise ValueError(f'{number}행: 구분은 수입 또는 지출로 입력해주세요.')
        try:
            amount = Decimal(str(values['금액']).replace(',', '').strip())
            if not amount.is_finite() or amount < 0 or amount != amount.to_integral_value() or amount > 10**15:
                raise ValueError
            amount = int(amount)
        except (InvalidOperation, ValueError):
            raise ValueError(f'{number}행: 금액은 0 이상의 정수로 입력해주세요.') from None
        month = day.strftime('%Y-%m')
        months[month]['income' if kind == '수입' else 'expense'] += amount
        if kind == '지출':
            categories[str(values['카테고리'] or '').strip() or '기타'] += amount
        count += 1
    if not count:
        raise ValueError('분석할 거래 내역이 없습니다.')
    monthly = [dict(month=month, **values) for month, values in sorted(months.items())]
    income = sum(row['income'] for row in monthly)
    expense = sum(row['expense'] for row in monthly)
    return {'income': income, 'expense': expense, 'balance': income-expense, 'count': count,
            'monthly': monthly, 'categories': sorted(categories.items(), key=lambda item: -item[1])}


def make_chart(summary):
    """웹 서버에서도 작동하도록 창을 띄우지 않고 PNG를 만듭니다."""
    rows = summary['monthly']
    figure = Figure(figsize=(8, 4), layout='constrained')
    axes = figure.subplots()
    x = list(range(len(rows)))
    axes.bar([i - .2 for i in x], [row['income'] for row in rows], width=.4, label='Income', color='#318265')
    axes.bar([i + .2 for i in x], [row['expense'] for row in rows], width=.4, label='Expense', color='#d19c85')
    axes.set_xticks(x, [row['month'] for row in rows], rotation=45 if len(rows) > 6 else 0)
    axes.set_ylabel('KRW')
    axes.ticklabel_format(axis='y', style='plain')
    axes.spines[['top', 'right']].set_visible(False)
    axes.legend(frameon=False)
    image = io.BytesIO()
    figure.set_dpi(120)
    FigureCanvasAgg(figure).print_png(image)
    return image.getvalue()


def make_excel(summary):
    book = Workbook()
    overview = book.active
    overview.title = '요약'
    for row in [['항목', '값'], ['총수입', summary['income']], ['총지출', summary['expense']],
                ['잔액', summary['balance']], ['거래 수', summary['count']]]:
        overview.append(row)
    monthly = book.create_sheet('월별')
    monthly.append(['월', '수입', '지출', '잔액'])
    for row in summary['monthly']:
        monthly.append([row['month'], row['income'], row['expense'], row['income']-row['expense']])
    categories = book.create_sheet('카테고리')
    categories.append(['카테고리', '지출'])
    for name, amount in summary['categories']:
        categories.append([name, amount])
        categories.cell(categories.max_row, 1).data_type = 's'
    for sheet in book:
        sheet.freeze_panes = 'A2'
        for column in sheet.columns:
            sheet.column_dimensions[column[0].column_letter].width = 20
    output = io.BytesIO()
    book.save(output)
    output.seek(0)
    return output


def clean_results():
    for path in RESULT_DIR.glob('*.xlsx'):
        try:
            if time.time() - path.stat().st_mtime > 3600:
                path.unlink(missing_ok=True)
        except FileNotFoundError:
            pass


@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'GET':
        return render_template('index.html')
    file = request.files.get('file')
    if not file or not file.filename:
        return render_template('index.html', error='엑셀 파일을 선택해주세요.'), 400
    if Path(file.filename).suffix.lower() not in ('.xlsx', '.xls'):
        return render_template('index.html', error='XLSX 또는 XLS 파일만 업로드할 수 있습니다.'), 400
    stage = 'READ'
    try:
        summary = analyze(read_rows(file))
        stage = 'CHART'
        chart = base64.b64encode(make_chart(summary)).decode('ascii')
        stage = 'EXPORT'
        excel = make_excel(summary).getvalue()
        stage = 'SAVE'
        clean_results()
        key = uuid.uuid4().hex
        (RESULT_DIR / f'{key}.xlsx').write_bytes(excel)
        stage = 'PAGE'
        return render_template('index.html', summary=summary, chart=chart, key=key)
    except ValueError as error:
        if stage == 'READ':
            return render_template('index.html', error=str(error)), 400
        return server_error(stage)
    except Exception:
        return server_error(stage)


def server_error(stage):
    error_id = f'{stage}-{uuid.uuid4().hex[:8]}'
    app.logger.exception('Excel analysis failed [%s]', error_id)
    labels = {'READ': '엑셀 읽기', 'CHART': '그래프 생성', 'EXPORT': '결과 엑셀 생성',
              'SAVE': '결과 저장', 'PAGE': '결과 표시'}
    message = f'{labels[stage]} 중 오류가 발생했습니다. 오류 번호: {error_id}'
    return render_template('index.html', error=message), 500


@app.get('/download/<key>')
def download(key):
    if len(key) != 32 or any(char not in '0123456789abcdef' for char in key):
        abort(404)
    clean_results()
    path = RESULT_DIR / f'{key}.xlsx'
    if not path.is_file():
        abort(404)
    return send_file(path, as_attachment=True, download_name='ledger-summary.xlsx')


@app.get('/template')
def template():
    book = Workbook()
    sheet = book.active
    sheet.title = '가계부'
    sheet.append(['날짜', '구분', '카테고리', '금액', '메모'])
    sheet.append(['2026-11-01', '수입', '용돈', 180000, '월초 용돈'])
    sheet.append(['2026-11-02', '지출', '식비', 12000, '점심'])
    output = io.BytesIO()
    book.save(output)
    output.seek(0)
    return send_file(output, as_attachment=True, download_name='ledger-template.xlsx')


@app.errorhandler(413)
def too_large(error):
    return render_template('index.html', error='파일 크기는 8MB 이하여야 합니다.'), 413


@app.errorhandler(404)
def not_found(error):
    return render_template('index.html', error='결과 파일이 만료되었거나 주소가 올바르지 않습니다. 다시 분석해주세요.'), 404


@app.template_filter('won')
def won(value):
    return f'{value:,}원'


if __name__ == '__main__':
    app.run(host='127.0.0.1', port=int(os.environ.get('PORT', 5000)))
