import json
import os
import sqlite3
import uuid
from datetime import datetime

DB_FILE = os.path.join(os.path.dirname(__file__), "history.db")
JSON_HISTORY_FILE = os.path.join(os.path.dirname(__file__), "history.json")

def get_db_connection():
    conn = sqlite3.connect(DB_FILE)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    """Initializes the SQLite database and migrates existing JSON history if found."""
    with get_db_connection() as conn:
        conn.execute("""
            CREATE TABLE IF NOT EXISTS history (
                id TEXT PRIMARY KEY,
                timestamp TEXT,
                job_role TEXT,
                skills TEXT,
                experience_level TEXT,
                interview_type TEXT,
                num_questions INTEGER,
                company_context TEXT,
                data TEXT
            )
        """)
        conn.commit()

    # Perform automatic data migration from history.json if it exists
    if os.path.exists(JSON_HISTORY_FILE):
        try:
            with open(JSON_HISTORY_FILE, "r", encoding="utf-8") as f:
                records = json.load(f)
            
            if isinstance(records, list) and len(records) > 0:
                with get_db_connection() as conn:
                    # Insert in reverse order (oldest first) so that timestamps/order are preserved correctly
                    for record in reversed(records):
                        # Check if already exists in DB to prevent duplicates
                        cur = conn.cursor()
                        cur.execute("SELECT 1 FROM history WHERE id = ?", (record.get("id"),))
                        if not cur.fetchone():
                            conn.execute("""
                                INSERT INTO history (
                                    id, timestamp, job_role, skills, experience_level, 
                                    interview_type, num_questions, company_context, data
                                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                            """, (
                                record.get("id"),
                                record.get("timestamp"),
                                record.get("job_role", "").strip(),
                                record.get("skills", "").strip(),
                                record.get("experience_level", "intermediate"),
                                record.get("interview_type", "mixed"),
                                int(record.get("num_questions", 10)),
                                record.get("company_context", "").strip(),
                                json.dumps(record.get("data", {}), ensure_ascii=False)
                            ))
                    conn.commit()
            
            # Backup/rename the old file to prevent running migration again
            bak_file = JSON_HISTORY_FILE + ".bak"
            if os.path.exists(bak_file):
                os.remove(bak_file)
            os.rename(JSON_HISTORY_FILE, bak_file)
            print(f"[SQLite Migration] Successfully migrated history from JSON to SQLite and backed up history.json.")
        except Exception as e:
            print(f"[SQLite Migration Error] Failed to migrate history.json: {str(e)}")

def save_history_item(metadata: dict, questions_data: dict) -> dict:
    item_id = str(uuid.uuid4())
    timestamp = datetime.utcnow().isoformat() + "Z"
    
    # Store questions data as JSON string in the DB
    data_str = json.dumps(questions_data, ensure_ascii=False)
    
    with get_db_connection() as conn:
        conn.execute("""
            INSERT INTO history (
                id, timestamp, job_role, skills, experience_level, 
                interview_type, num_questions, company_context, data
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            item_id,
            timestamp,
            metadata.get("job_role", "").strip(),
            metadata.get("skills", "").strip(),
            metadata.get("experience_level", "intermediate"),
            metadata.get("interview_type", "mixed"),
            int(metadata.get("num_questions", 10)),
            metadata.get("company_context", "").strip(),
            data_str
        ))
        conn.commit()
        
    return {
        "id": item_id,
        "timestamp": timestamp,
        "job_role": metadata.get("job_role", "").strip(),
        "skills": metadata.get("skills", "").strip(),
        "experience_level": metadata.get("experience_level", "intermediate"),
        "interview_type": metadata.get("interview_type", "mixed"),
        "num_questions": int(metadata.get("num_questions", 10)),
        "company_context": metadata.get("company_context", "").strip(),
        "data": questions_data
    }

def get_history_list() -> list:
    with get_db_connection() as conn:
        cursor = conn.cursor()
        # Omit full raw 'data' column for quick transfer of list
        cursor.execute("""
            SELECT id, timestamp, job_role, skills, experience_level, 
                   interview_type, num_questions, company_context 
            FROM history 
            ORDER BY timestamp DESC
        """)
        rows = cursor.fetchall()
        
        metadata_list = []
        for row in rows:
            metadata_list.append({
                "id": row["id"],
                "timestamp": row["timestamp"],
                "job_role": row["job_role"],
                "skills": row["skills"],
                "experience_level": row["experience_level"],
                "interview_type": row["interview_type"],
                "num_questions": row["num_questions"],
                "company_context": row["company_context"]
            })
        return metadata_list

def get_history_item(item_id: str) -> dict:
    with get_db_connection() as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM history WHERE id = ?", (item_id,))
        row = cursor.fetchone()
        
        if not row:
            return None
            
        return {
            "id": row["id"],
            "timestamp": row["timestamp"],
            "job_role": row["job_role"],
            "skills": row["skills"],
            "experience_level": row["experience_level"],
            "interview_type": row["interview_type"],
            "num_questions": row["num_questions"],
            "company_context": row["company_context"],
            "data": json.loads(row["data"]) if row["data"] else {}
        }

def delete_history_item(item_id: str) -> bool:
    with get_db_connection() as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT 1 FROM history WHERE id = ?", (item_id,))
        if not cursor.fetchone():
            return False
            
        conn.execute("DELETE FROM history WHERE id = ?", (item_id,))
        conn.commit()
        return True

# Initialize database on module load
init_db()
