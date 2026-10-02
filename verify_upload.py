import io
import sqlite3
from openpyxl import Workbook
from app import app

conn = sqlite3.connect('student_records.db')
conn.execute("DELETE FROM students WHERE register_number IN ('2024001', '2024002')")
conn.commit()
conn.close()

wb = Workbook()
ws = wb.active
ws.append(['Student Name', 'Register Number', 'Student Phone', 'Course', 'Year', 'Semester', 'Section', 'Parent Name', 'Parent Phone'])
ws.append(['Alice', '2024001', '9876543210', 'BCA', 1, 'I Semester', 'A', 'Bob', '9988776655'])
ws.append(['Carol', '2024002', '9123456780', 'BBA', '2nd Year', 'II Semester', 'B', 'Dan', '9000000000'])

buf = io.BytesIO()
wb.save(buf)
buf.seek(0)

with app.test_client() as client:
    with client.session_transaction() as sess:
        sess['faculty_username'] = 'alvas'
    upload_response = client.post(
        '/faculty/students/upload',
        data={'file': (buf, 'Students.xlsx')},
        content_type='multipart/form-data',
        follow_redirects=False,
    )
    print('UPLOAD_STATUS', upload_response.status_code)
    print('UPLOAD_BODY_SNIPPET', upload_response.get_data(as_text=True)[:250])

    records_response = client.get('/faculty/students-records')
    body = records_response.get_data(as_text=True)
    print('RECORDS_STATUS', records_response.status_code)
    print('HAS_SORT_ERROR', "'<' not supported between instances of 'str' and 'int'" in body or "'<' not supported between instances of 'str' and 'int'" in body)
    print('HAS_ALICE', 'Alice' in body)
    print('HAS_CAROL', 'Carol' in body)
