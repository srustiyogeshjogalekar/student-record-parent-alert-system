import sqlite3

conn = sqlite3.connect('student_records.db')
print('TABLES', conn.execute("SELECT name FROM sqlite_master WHERE type='table' ORDER BY name").fetchall())
for tbl in ['students', 'attendance', 'internal_marks']:
    try:
        print('\nTABLE', tbl)
        print(conn.execute(f'PRAGMA table_info({tbl})').fetchall())
        print(conn.execute(f'SELECT * FROM {tbl} LIMIT 5').fetchall())
    except Exception as exc:
        print('ERR', tbl, exc)
conn.close()
