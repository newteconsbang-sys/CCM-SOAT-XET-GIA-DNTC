"""Database models and persistence for CCM-AI MEP Price Review."""
import sqlite3
import json
import uuid
from datetime import datetime, timezone
from pathlib import Path
from typing import List, Dict, Any, Optional

ROOT = Path(__file__).resolve().parents[1]
DB_PATH = ROOT / "data" / "ccm.sqlite3"

def now_iso() -> str:
    return datetime.now(timezone.utc).isoformat()

def get_connection() -> sqlite3.Connection:
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn

def init_mep_db():
    """Initializes tables for MEP price review system."""
    with get_connection() as conn:
        conn.executescript("""
        CREATE TABLE IF NOT EXISTS mep_libraries (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            system TEXT NOT NULL DEFAULT 'Khác',
            code TEXT,
            description TEXT NOT NULL,
            specification TEXT,
            brand TEXT,
            origin TEXT,
            model TEXT,
            unit TEXT,
            material_price REAL,
            labor_price REAL,
            total_price REAL NOT NULL,
            source TEXT,
            vendor TEXT,
            project TEXT,
            date TEXT,
            version INTEGER DEFAULT 1,
            normalized_desc TEXT,
            created_at TEXT
        );

        CREATE INDEX IF NOT EXISTS idx_mep_lib_sys ON mep_libraries(system);
        CREATE INDEX IF NOT EXISTS idx_mep_lib_code ON mep_libraries(code);
        CREATE INDEX IF NOT EXISTS idx_mep_lib_desc ON mep_libraries(description);

        CREATE TABLE IF NOT EXISTS mep_price_history (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            library_id INTEGER,
            old_price REAL,
            new_price REAL,
            changed_by TEXT,
            reason TEXT,
            timestamp TEXT
        );

        CREATE TABLE IF NOT EXISTS mep_learning_memory (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            boq_code TEXT,
            boq_description TEXT NOT NULL,
            boq_unit TEXT,
            library_code TEXT,
            library_description TEXT NOT NULL,
            library_unit TEXT,
            confirmed_price REAL,
            notes TEXT,
            created_at TEXT
        );

        CREATE INDEX IF NOT EXISTS idx_mep_mem_boq_desc ON mep_learning_memory(boq_description);

        CREATE TABLE IF NOT EXISTS mep_workbooks (
            id TEXT PRIMARY KEY,
            filename TEXT NOT NULL,
            original_filename TEXT NOT NULL,
            file_path TEXT NOT NULL,
            backup_path TEXT NOT NULL,
            file_type TEXT NOT NULL,
            sheets_info TEXT,
            uploaded_at TEXT
        );

        CREATE TABLE IF NOT EXISTS mep_review_sessions (
            id TEXT PRIMARY KEY,
            workbook_id TEXT NOT NULL,
            sheet_name TEXT NOT NULL,
            threshold_percent REAL DEFAULT 10.0,
            mapping_config TEXT,
            total_items INTEGER DEFAULT 0,
            passed_count INTEGER DEFAULT 0,
            higher_count INTEGER DEFAULT 0,
            lower_count INTEGER DEFAULT 0,
            not_found_count INTEGER DEFAULT 0,
            spec_mismatch_count INTEGER DEFAULT 0,
            manual_check_count INTEGER DEFAULT 0,
            total_variance REAL DEFAULT 0.0,
            status TEXT DEFAULT 'completed',
            created_at TEXT
        );

        CREATE TABLE IF NOT EXISTS mep_review_items (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            session_id TEXT NOT NULL,
            sheet_name TEXT NOT NULL,
            excel_row INTEGER NOT NULL,
            excel_col INTEGER NOT NULL,
            boq_code TEXT,
            boq_description TEXT NOT NULL,
            boq_spec TEXT,
            boq_brand TEXT,
            boq_model TEXT,
            boq_unit TEXT,
            boq_quantity REAL,
            boq_price REAL NOT NULL,
            lib_id INTEGER,
            lib_code TEXT,
            lib_description TEXT,
            lib_spec TEXT,
            lib_brand TEXT,
            lib_model TEXT,
            lib_unit TEXT,
            lib_price REAL,
            lib_source TEXT,
            difference REAL,
            diff_percent REAL,
            status TEXT NOT NULL,
            match_method TEXT,
            match_score REAL,
            warning_color TEXT,
            user_action TEXT DEFAULT 'none',
            user_note TEXT,
            edited_price REAL,
            candidates_json TEXT,
            updated_at TEXT
        );

        CREATE INDEX IF NOT EXISTS idx_mep_rev_sess ON mep_review_items(session_id);
        CREATE INDEX IF NOT EXISTS idx_mep_rev_row ON mep_review_items(sheet_name, excel_row);

        CREATE TABLE IF NOT EXISTS mep_settings (
            key TEXT PRIMARY KEY,
            value TEXT NOT NULL
        );
        """)

init_mep_db()

