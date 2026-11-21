import base64
import io
import re
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from openpyxl import Workbook, load_workbook
import app as ledger


class LedgerTests(unittest.TestCase):
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory()
        self.patch = patch.object(ledger, 'RESULT_DIR', Path(self.directory.name))
        self.patch.start()
        self.client = ledger.app.test_client()

    def tearDown(self):
        self.patch.stop()
        self.directory.cleanup()

    def workbook(self, rows):
        book = Workbook()
        for row in rows:
            book.active.append(row)
        output = io.BytesIO()
        book.save(output)
        output.seek(0)
        return output

    def check_result(self, response):
        self.assertEqual(response.status_code, 200)
        self.assertIn('168,000원', response.text)
        self.assertIn('2건', response.text)
        png = re.search(r'data:image/png;base64,([^" ]+)', response.text).group(1)
        self.assertTrue(base64.b64decode(png).startswith(b'\x89PNG'))
        url = re.search(r'href="(/download/[a-f0-9]+)"', response.text).group(1)
        download = self.client.get(url)
        self.assertEqual(download.status_code, 200)
        book = load_workbook(io.BytesIO(download.data))
        self.assertEqual(book['요약']['B4'].value, 168000)
        self.assertEqual(book.sheetnames, ['요약', '월별', '카테고리'])
        book.close()
        download.close()

    def test_xlsx_upload_graph_and_download(self):
        template = self.client.get('/template')
        self.check_result(self.client.post('/', data={'file': (io.BytesIO(template.data), 'budget.xlsx')}))

    def test_read_without_upload_stream_seeking(self):
        data = self.client.get('/template').data
        class ReadOnlyUpload:
            filename = 'budget.xlsx'
            def read(self):
                return data
        summary = ledger.analyze(ledger.read_rows(ReadOnlyUpload()))
        self.assertEqual(summary['balance'], 168000)

    def test_graph_failure_reports_stage(self):
        data = self.client.get('/template').data
        with patch.object(ledger, 'make_chart', side_effect=RuntimeError('test graph failure')):
            with self.assertLogs(ledger.app.logger, level='ERROR') as logs:
                response = self.client.post('/', data={'file': (io.BytesIO(data), 'budget.xlsx')})
        self.assertEqual(response.status_code, 500)
        self.assertIn('CHART-', response.text)
        self.assertIn('test graph failure', '\n'.join(logs.output))
        self.assertNotIn('test graph failure', response.text)

    def test_xls_upload_graph_and_download(self):
        fixture = Path(__file__).with_name('sample.xls').read_bytes()
        self.check_result(self.client.post('/', data={'file': (io.BytesIO(fixture), 'budget.xls')}))

    def test_invalid_rows_and_files(self):
        headers = ['날짜', '구분', '카테고리', '금액']
        for rows in [[headers], [['日付', '金額']],
                     [headers, ['bad date', '수입', '용돈', 10]],
                     [headers, ['2026-11-01', '기타', '용돈', 10]],
                     [headers, ['2026-11-01', '지출', '식비', -10]]]:
            with self.subTest(rows=rows):
                response = self.client.post('/', data={'file': (self.workbook(rows), 'bad.xlsx')})
                self.assertEqual(response.status_code, 400)
                self.assertIn('role="alert"', response.text)
        self.assertEqual(self.client.post('/').status_code, 400)
        self.assertEqual(self.client.post('/', data={'file': (io.BytesIO(b'bad'), 'bad.csv')}).status_code, 400)

    def test_home_zero_income_and_unknown_download(self):
        self.assertEqual(self.client.get('/').status_code, 200)
        rows = [['날짜', '구분', '카테고리', '금액'], ['2026-11-01', '지출', '식비', 100]]
        response = self.client.post('/', data={'file': (self.workbook(rows), 'expense.xlsx')})
        self.assertEqual(response.status_code, 200)
        self.assertIn('-100원', response.text)
        self.assertEqual(self.client.get('/download/' + 'a' * 32).status_code, 404)


if __name__ == '__main__':
    unittest.main()
