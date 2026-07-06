import sqlite3
import os

DB_PATH = os.path.join(os.path.dirname(__file__), 'equipment_qr.db')

def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn

def init_db():
    conn = get_db()
    c = conn.cursor()

    c.execute('''
        CREATE TABLE IF NOT EXISTS equipment (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            eq_number TEXT NOT NULL,
            eq_name TEXT NOT NULL,
            section TEXT NOT NULL,
            location TEXT,
            note TEXT,
            created_at DATETIME DEFAULT (datetime('now','localtime'))
        )
    ''')

    c.execute('''
        CREATE TABLE IF NOT EXISTS fault_history (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            equipment_id INTEGER NOT NULL,
            eq_number TEXT NOT NULL,
            eq_name TEXT NOT NULL,
            occurred_at DATETIME DEFAULT (datetime('now','localtime')),
            symptom TEXT,
            cause TEXT,
            action_detail TEXT,
            worker TEXT,
            completed_at DATE,
            grade TEXT,
            photo_fault TEXT,
            photo_action TEXT,
            created_at DATETIME DEFAULT (datetime('now','localtime')),
            FOREIGN KEY (equipment_id) REFERENCES equipment(id)
        )
    ''')

    c.execute('''
        CREATE TABLE IF NOT EXISTS used_parts (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            fault_id INTEGER NOT NULL,
            part_name TEXT NOT NULL,
            quantity INTEGER DEFAULT 1,
            note TEXT,
            FOREIGN KEY (fault_id) REFERENCES fault_history(id) ON DELETE CASCADE
        )
    ''')

    c.execute('''
        CREATE TABLE IF NOT EXISTS fault_photos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            fault_id INTEGER NOT NULL,
            filename TEXT NOT NULL,
            label TEXT,
            photo_type TEXT DEFAULT 'fault',
            FOREIGN KEY (fault_id) REFERENCES fault_history(id) ON DELETE CASCADE
        )
    ''')

    c.execute('''
        CREATE TABLE IF NOT EXISTS workers (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL UNIQUE
        )
    ''')

    conn.commit()

    # 기존 고장이력에 기록된 작업자 이름을 workers 테이블로 이관
    try:
        c.execute("""
            INSERT OR IGNORE INTO workers (name)
            SELECT DISTINCT TRIM(worker) FROM fault_history
            WHERE worker IS NOT NULL AND TRIM(worker) != ''
        """)
        conn.commit()
    except Exception:
        pass

    # status 컬럼 마이그레이션
    try:
        c.execute("ALTER TABLE fault_history ADD COLUMN status TEXT DEFAULT '미조치'")
        conn.commit()
        c.execute("""UPDATE fault_history SET status='완료'
                     WHERE completed_at IS NOT NULL AND completed_at != ''
                     AND (status IS NULL OR status='미조치')""")
        c.execute("""UPDATE fault_history SET status='진행중'
                     WHERE (completed_at IS NULL OR completed_at='')
                     AND action_detail IS NOT NULL AND action_detail != ''
                     AND (status IS NULL OR status='미조치')""")
        conn.commit()
    except Exception:
        pass

    # grade_note 컬럼 마이그레이션 (D등급 기타 내용)
    try:
        c.execute("ALTER TABLE fault_history ADD COLUMN grade_note TEXT")
        conn.commit()
    except Exception:
        pass

    # status_note 컬럼 마이그레이션 (대기 상태 사유)
    try:
        c.execute("ALTER TABLE fault_history ADD COLUMN status_note TEXT")
        conn.commit()
    except Exception:
        pass

    # fault_photos.photo_type 컬럼 마이그레이션
    try:
        c.execute("ALTER TABLE fault_photos ADD COLUMN photo_type TEXT DEFAULT 'fault'")
        conn.commit()
    except Exception:
        pass

    conn.close()
    print("DB 초기화 완료:", DB_PATH)
