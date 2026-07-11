"""
Lapisan penyimpanan lokal (SQLite) sesuai Technical Stack pada presentasi:
Web Server (Python) <--driver--> Database Server (SQLite).
"""
import sqlite3
from datetime import datetime
from pathlib import Path

DB_PATH = Path(__file__).parent / "guepedia_ai.db"


def get_conn():
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    conn = get_conn()
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS naskah_analysis (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            file_name TEXT,
            title TEXT,
            word_count INTEGER,
            genre TEXT,
            blurb_1 TEXT,
            blurb_2 TEXT,
            blurb_3 TEXT,
            chosen_blurb INTEGER DEFAULT 0,
            created_at TEXT
        )
        """
    )
    conn.commit()
    conn.close()


def save_analysis(file_name, title, word_count, genre, blurbs):
    conn = get_conn()
    cur = conn.execute(
        """
        INSERT INTO naskah_analysis
            (file_name, title, word_count, genre, blurb_1, blurb_2, blurb_3, created_at)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (
            file_name,
            title,
            word_count,
            genre,
            blurbs[0],
            blurbs[1],
            blurbs[2],
            datetime.now().isoformat(timespec="seconds"),
        ),
    )
    conn.commit()
    new_id = cur.lastrowid
    conn.close()
    return new_id


def update_chosen_blurb(record_id, chosen_index):
    conn = get_conn()
    conn.execute(
        "UPDATE naskah_analysis SET chosen_blurb = ? WHERE id = ?",
        (chosen_index, record_id),
    )
    conn.commit()
    conn.close()


def get_recent(limit=10):
    conn = get_conn()
    rows = conn.execute(
        "SELECT * FROM naskah_analysis ORDER BY id DESC LIMIT ?", (limit,)
    ).fetchall()
    conn.close()
    return rows
