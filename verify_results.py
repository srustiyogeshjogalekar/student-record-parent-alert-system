from app import app

with app.test_client() as client:
    with client.session_transaction() as sess:
        sess['faculty_username'] = 'alvas'
    resp = client.get('/faculty/semester-results?course=BCA&year=3rd%20Year&semester=I%20Semester')
    print('STATUS', resp.status_code)
    body = resp.get_data(as_text=True)
    print('HAS_RESULTS_PAGE', 'Semester Results' in body)
    print('HAS_ERROR', 'Traceback' in body or 'Internal Server Error' in body)
    print('BODY_SNIPPET', body[:400])
