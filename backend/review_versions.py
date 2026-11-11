"""Các lần Lưu của gói soát xét: mỗi lần bấm Lưu tạo một bản (số liệu bảng + file Excel kèm cột soát xét)
để xem lại / tải lại bản cũ trong tab Lịch sử chỉnh sửa."""
import json
import threading
from pathlib import Path
from typing import Any, Dict, List, Optional

from backend import mep_db
from backend.mep_db import get_connection, now_iso

SNAP_FIELDS = ("rv_quantity", "rv_material", "rv_labor", "rv_machine", "rv_total", "rv_note", "rv_dep_pct", "rv_extra",
               "boq_quantity", "boq_material", "boq_labor", "boq_dep_pct", "boq_price", "boq_stt", "boq_description",
               "boq_unit", "is_added", "sheet_name", "excel_row")


def ensure_schema():
    with get_connection() as conn:
        conn.executescript("""
        CREATE TABLE IF NOT EXISTS mep_review_versions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            package_id INTEGER NOT NULL,
            session_id TEXT,
            no INTEGER NOT NULL,
            saved_at TEXT NOT NULL,
            saved_by TEXT,
            changes INTEGER DEFAULT 0,
            note TEXT,
            ntp_total REAL,
            review_total REAL,
            data_json TEXT,
            file_path TEXT
        );
        CREATE INDEX IF NOT EXISTS idx_rev_versions_pkg ON mep_review_versions(package_id);
        """)


def _dir(pid: int) -> Path:
    d = Path(mep_db.DB_PATH).parent / "review_versions" / f"p{pid}"
    d.mkdir(parents=True, exist_ok=True)
    return d


def create(package: Dict[str, Any], user: str, changes: int, note: str, export_fn, background: bool = True) -> Dict[str, Any]:
    """Chụp bảng hiện tại thành một bản lưu; export_fn(package) → bytes file Excel (lỗi thì bỏ qua file)."""
    items = package.get("items") or []
    data = {str(i["id"]): {k: i.get(k) for k in SNAP_FIELDS} for i in items}
    ntp = sum(float(i.get("ntp_amount") or 0) for i in items)
    rev = sum(float(i.get("rv_amount") or 0) for i in items)
    with get_connection() as conn:
        no = (conn.execute("SELECT MAX(no) FROM mep_review_versions WHERE package_id=?", (package["id"],)).fetchone()[0] or 0) + 1
        cur = conn.execute("""INSERT INTO mep_review_versions (package_id, session_id, no, saved_at, saved_by, changes, note,
                              ntp_total, review_total, data_json) VALUES (?,?,?,?,?,?,?,?,?,?)""",
                           (package["id"], package.get("session_id"), no, now_iso(), user, changes, note, ntp, rev,
                            json.dumps(data, ensure_ascii=False)))
        vid = cur.lastrowid
    def _write_file():
        try:
            content = export_fn(package)
            if content:
                path = _dir(package["id"]) / f"lan-luu-{no:03d}.xlsx"
                path.write_bytes(content)
                with get_connection() as conn:
                    conn.execute("UPDATE mep_review_versions SET file_path=? WHERE id=?", (str(path), vid))
        except Exception:  # file gốc mất / lỗi xuất: vẫn giữ bản số liệu
            pass
    # Xuất file Excel của bản lưu chạy nền để nút Lưu trả về ngay (file sẵn sau vài giây)
    if background:
        threading.Thread(target=_write_file, daemon=True).start()
    else:
        _write_file()
    return {"id": vid, "no": no}


def list_versions(pid: int) -> List[Dict[str, Any]]:
    with get_connection() as conn:
        rows = conn.execute("""SELECT id, no, saved_at, saved_by, changes, note, ntp_total, review_total,
                               file_path IS NOT NULL AS has_file FROM mep_review_versions WHERE package_id=? ORDER BY no DESC""",
                            (pid,)).fetchall()
    return [dict(r) for r in rows]


def get(pid: int, vid: int) -> Optional[Dict[str, Any]]:
    with get_connection() as conn:
        r = conn.execute("SELECT * FROM mep_review_versions WHERE package_id=? AND id=?", (pid, vid)).fetchone()
    if not r:
        return None
    d = dict(r)
    d["data"] = json.loads(d.pop("data_json") or "{}")
    return d
