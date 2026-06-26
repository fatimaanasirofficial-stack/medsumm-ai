import sqlite3

DB_NAME = "reports.db"

def init_db():
    conn = sqlite3.connect(DB_NAME)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS reports (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            filename TEXT,
            upload_date TEXT,
            raw_text TEXT,
            summary TEXT,
            model_used TEXT
        )
    """)
    conn.commit()
    conn.close()

def save_report(filename, raw_text, summary, model_used):
    conn = sqlite3.connect(DB_NAME)
    conn.execute(
        "INSERT INTO reports (filename, upload_date, raw_text, summary, model_used) VALUES (?, datetime('now'), ?, ?, ?)",
        (filename, raw_text, summary, model_used)
    )
    conn.commit()
    conn.close()

def get_all_reports():
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.execute("SELECT id, filename, upload_date, summary FROM reports ORDER BY upload_date DESC")
    rows = cursor.fetchall()
    conn.close()
    return rows

def delete_report(report_id):
    conn = sqlite3.connect(DB_NAME)
    conn.execute("DELETE FROM reports WHERE id=?", (report_id,))
    conn.commit()
    conn.close() 