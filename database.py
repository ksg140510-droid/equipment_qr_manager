import sqlite3
import os

DB_PATH = os.path.join(os.path.dirname(__file__), 'equipment_qr.db')

def get_db():
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    conn.execute("PRAGMA journal_mode = WAL")
    conn.execute("PRAGMA busy_timeout = 30000")
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

    c.execute('''
        CREATE TABLE IF NOT EXISTS audit_log (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            actor TEXT,
            action TEXT NOT NULL,
            target_type TEXT NOT NULL,
            target_id INTEGER,
            detail TEXT,
            created_at DATETIME DEFAULT (datetime('now','localtime'))
        )
    ''')

    c.execute('''
        CREATE TABLE IF NOT EXISTS production_lots (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            model_name TEXT NOT NULL,
            lot_number TEXT NOT NULL,
            started_at DATETIME NOT NULL,
            ended_at DATETIME,
            status TEXT NOT NULL DEFAULT '진행중',
            started_by TEXT,
            created_at DATETIME DEFAULT (datetime('now','localtime'))
        )
    ''')

    c.execute('''
        CREATE TABLE IF NOT EXISTS production_lot_section_stats (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            lot_id INTEGER NOT NULL,
            section TEXT NOT NULL,
            available_seconds INTEGER NOT NULL,
            fault_seconds INTEGER NOT NULL,
            fault_count INTEGER NOT NULL,
            uptime_rate REAL NOT NULL,
            FOREIGN KEY (lot_id) REFERENCES production_lots(id)
        )
    ''')

    c.execute('''
        CREATE TABLE IF NOT EXISTS break_schedules (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL UNIQUE,
            start_time TEXT NOT NULL,
            end_time TEXT NOT NULL
        )
    ''')

    c.execute('''
        CREATE TABLE IF NOT EXISTS production_lot_sections (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            lot_id INTEGER NOT NULL,
            section TEXT NOT NULL,
            started_at DATETIME NOT NULL,
            ended_at DATETIME,
            status TEXT NOT NULL DEFAULT '진행중',
            available_seconds INTEGER,
            fault_seconds INTEGER,
            fault_count INTEGER,
            uptime_rate REAL,
            FOREIGN KEY (lot_id) REFERENCES production_lots(id)
        )
    ''')

    c.execute('''
        CREATE TABLE IF NOT EXISTS production_lot_section_pauses (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            lot_section_id INTEGER NOT NULL,
            paused_at DATETIME NOT NULL,
            resumed_at DATETIME,
            reason TEXT,
            FOREIGN KEY (lot_section_id) REFERENCES production_lot_sections(id)
        )
    ''')

    conn.commit()

    # 휴식시간 기본 항목 시딩 (최초 1회, 실제 값은 설정 화면에서 수정)
    try:
        c.execute("""
            INSERT OR IGNORE INTO break_schedules (name, start_time, end_time) VALUES
            ('오전휴식', '10:00', '10:10'),
            ('점심시간', '12:00', '13:00'),
            ('오후휴식', '15:00', '15:10'),
            ('잔업', '18:00', '18:10')
        """)
        conn.commit()
    except Exception:
        pass

    # fault_history.lot_id 컬럼 마이그레이션 (LOT과 고장 연결)
    try:
        c.execute("ALTER TABLE fault_history ADD COLUMN lot_id INTEGER")
        conn.commit()
    except Exception:
        pass

    # fault_history.lot_section_id 컬럼 마이그레이션 (섹션별 가동시간 구간과 고장 연결)
    try:
        c.execute("ALTER TABLE fault_history ADD COLUMN lot_section_id INTEGER")
        conn.commit()
    except Exception:
        pass

    # production_lot_section_pauses.reason 컬럼 마이그레이션 (가동정지 사유)
    try:
        c.execute("ALTER TABLE production_lot_section_pauses ADD COLUMN reason TEXT")
        conn.commit()
    except Exception:
        pass

    # production_lot_sections.pause_downtime_seconds 컬럼 마이그레이션
    # (퇴근 외 사유의 가동정지 시간 - 설비가동률에 반영되는 다운타임)
    try:
        c.execute("ALTER TABLE production_lot_sections ADD COLUMN pause_downtime_seconds INTEGER")
        conn.commit()
    except Exception:
        pass

    # 조회 성능을 위한 인덱스 (생산LOT 기능 추가로 조회 빈도가 늘어난 컬럼들)
    try:
        c.execute("CREATE INDEX IF NOT EXISTS idx_fault_equipment_id ON fault_history(equipment_id)")
        c.execute("CREATE INDEX IF NOT EXISTS idx_fault_lot_id ON fault_history(lot_id)")
        c.execute("CREATE INDEX IF NOT EXISTS idx_fault_lot_section_id ON fault_history(lot_section_id)")
        c.execute("CREATE INDEX IF NOT EXISTS idx_lot_sections_lot_id ON production_lot_sections(lot_id)")
        c.execute("CREATE INDEX IF NOT EXISTS idx_lot_sections_section_status ON production_lot_sections(section, status)")
        c.execute("CREATE INDEX IF NOT EXISTS idx_lot_section_pauses_section_id ON production_lot_section_pauses(lot_section_id)")
        conn.commit()
    except Exception:
        pass

    # 섹션당 동시 가동(진행중/일시정지)은 하나만 존재하도록 DB 레벨에서 강제
    # (동시 요청으로 같은 섹션이 중복 시작되는 경쟁 상태 방지)
    try:
        c.execute("""
            CREATE UNIQUE INDEX IF NOT EXISTS ux_lot_sections_active_per_section
            ON production_lot_sections(section)
            WHERE status IN ('진행중','일시정지')
        """)
        conn.commit()
    except Exception:
        pass

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

    # equipment.qr_reset_at / qr_last_ip 컬럼 마이그레이션
    # (QR코드가 실제로 재생성될 때마다 시각을 기록해, 화면에 표시된 QR과 이미 인쇄된 QR이
    # 서로 달라졌을 수 있음을 다른 작업자들도 알 수 있게 함)
    try:
        c.execute("ALTER TABLE equipment ADD COLUMN qr_reset_at DATETIME")
        conn.commit()
    except Exception:
        pass
    try:
        c.execute("ALTER TABLE equipment ADD COLUMN qr_last_ip TEXT")
        conn.commit()
    except Exception:
        pass

    conn.close()
    print("DB 초기화 완료:", DB_PATH)
