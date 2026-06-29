# -*- coding: utf-8 -*-
"""
数据库模块 - PostgreSQL
自动建表，提供与JSON存储等价的接口
"""
import os
import json
import uuid
import psycopg2
import psycopg2.extras
from datetime import datetime

# 数据库配置（从环境变量读取）
DB_HOST = os.environ.get("DB_HOST", "")
DB_PORT = os.environ.get("DB_PORT", "5432")
DB_NAME = os.environ.get("DB_NAME", "")
DB_USER = os.environ.get("DB_USER", "")
DB_PASS = os.environ.get("DB_PASS", "")

DB_ENABLED = bool(DB_HOST and DB_NAME and DB_USER)

_pool = None

def get_conn():
    """获取数据库连接"""
    return psycopg2.connect(
        host=DB_HOST, port=DB_PORT, dbname=DB_NAME,
        user=DB_USER, password=DB_PASS, connect_timeout=10
    )

def init_db():
    """初始化数据库表"""
    if not DB_ENABLED:
        return
    conn = get_conn()
    cur = conn.cursor()
    # 需求记录
    cur.execute("""
        CREATE TABLE IF NOT EXISTS bot_feedback (
            id TEXT PRIMARY KEY,
            created_time TIMESTAMP,
            org TEXT, contact TEXT, req_type TEXT,
            description TEXT, status TEXT DEFAULT '待处理', extra TEXT
        );
    """)
    # 访问记录
    cur.execute("""
        CREATE TABLE IF NOT EXISTS bot_visits (
            id SERIAL PRIMARY KEY,
            visit_time TIMESTAMP, visit_date DATE,
            ip TEXT, path TEXT, ua TEXT
        );
    """)
    # 问答日志
    cur.execute("""
        CREATE TABLE IF NOT EXISTS bot_chat_logs (
            id TEXT PRIMARY KEY,
            log_time TIMESTAMP, log_date DATE,
            ip TEXT, user_msg TEXT, bot_msg TEXT,
            matched BOOLEAN, source TEXT,
            duration_ms INTEGER, error TEXT
        );
    """)
    # 文档
    cur.execute("""
        CREATE TABLE IF NOT EXISTS bot_docs (
            id TEXT PRIMARY KEY,
            name TEXT, file_path TEXT, file_size BIGINT,
            content TEXT, upload_time TIMESTAMP
        );
    """)
    # 知识库FAQ
    cur.execute("""
        CREATE TABLE IF NOT EXISTS bot_kb_faq (
            id TEXT PRIMARY KEY,
            category TEXT, question TEXT,
            keywords TEXT, answer TEXT
        );
    """)
    conn.commit()
    conn.close()
    print(f"[DB] 数据库已连接: {DB_HOST}:{DB_PORT}/{DB_NAME}")
    print("[DB] 表结构初始化完成")


# ============================================================
# 需求记录
# ============================================================
def db_load_requests():
    conn = get_conn()
    cur = conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
    cur.execute("SELECT id, to_char(created_time,'YYYY-MM-DD HH24:MI:SS') as time, org, contact, req_type as type, description, status, extra FROM bot_feedback ORDER BY created_time")
    rows = cur.fetchall()
    conn.close()
    return [dict(r) for r in rows]

def db_save_requests(data):
    """全量覆盖（删除后重新插入）"""
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("DELETE FROM bot_feedback")
    for r in data:
        cur.execute("""INSERT INTO bot_feedback (id, created_time, org, contact, req_type, description, status, extra)
                       VALUES (%s, %s, %s, %s, %s, %s, %s, %s)""",
                    (r["id"], r.get("time", datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
                     r.get("org",""), r.get("contact",""), r.get("type",""),
                     r.get("description",""), r.get("status","待处理"), r.get("extra","")))
    conn.commit()
    conn.close()

def db_add_request(org, contact, rtype, desc, extra=""):
    rid = str(int(datetime.now().timestamp() * 1000)) + uuid.uuid4().hex[:4]
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("""INSERT INTO bot_feedback (id, created_time, org, contact, req_type, description, status, extra)
                   VALUES (%s, %s, %s, %s, %s, %s, '待处理', %s)""",
                (rid, now, org, contact, rtype, desc, extra))
    conn.commit()
    conn.close()
    return {"id": rid, "time": now, "org": org, "contact": contact, "type": rtype, "description": desc, "status": "待处理", "extra": extra}

def db_update_request_status(rid, status):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("UPDATE bot_feedback SET status=%s WHERE id=%s", (status, rid))
    affected = cur.rowcount
    conn.commit()
    conn.close()
    return affected > 0

def db_delete_request(rid):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("DELETE FROM bot_feedback WHERE id=%s", (rid,))
    affected = cur.rowcount
    conn.commit()
    conn.close()
    return affected > 0


# ============================================================
# 访问记录
# ============================================================
def db_record_visit(ip, path, ua):
    now = datetime.now()
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("""INSERT INTO bot_visits (visit_time, visit_date, ip, path, ua)
                   VALUES (%s, %s, %s, %s, %s)""",
                (now, now.date(), ip[:100], path[:200], ua[:200]))
    conn.commit()
    conn.close()

def db_load_visits():
    conn = get_conn()
    cur = conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
    cur.execute("SELECT to_char(visit_time,'YYYY-MM-DD HH24:MI:SS') as time, to_char(visit_date,'YYYY-MM-DD') as date, ip, path, ua FROM bot_visits ORDER BY visit_time DESC LIMIT 5000")
    rows = cur.fetchall()
    conn.close()
    return [dict(r) for r in rows]


# ============================================================
# 问答日志
# ============================================================
def db_record_chat_log(user_msg, bot_answer, matched, source, duration_ms, ip, error=None):
    lid = uuid.uuid4().hex[:12]
    now = datetime.now()
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("""INSERT INTO bot_chat_logs (id, log_time, log_date, ip, user_msg, bot_msg, matched, source, duration_ms, error)
                   VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s)""",
                (lid, now, now.date(), ip[:100], user_msg, bot_answer, matched, source, duration_ms, error))
    conn.commit()
    conn.close()

def db_load_chat_logs():
    conn = get_conn()
    cur = conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
    cur.execute("""SELECT id, to_char(log_time,'YYYY-MM-DD HH24:MI:SS') as time,
                   to_char(log_date,'YYYY-MM-DD') as date, ip, user_msg as user, bot_msg as bot,
                   matched, source, duration_ms, error
                   FROM bot_chat_logs ORDER BY log_time DESC LIMIT 5000""")
    rows = cur.fetchall()
    conn.close()
    return [dict(r) for r in rows]

def db_clear_chat_logs():
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("DELETE FROM bot_chat_logs")
    conn.commit()
    conn.close()


# ============================================================
# 文档
# ============================================================
def db_load_docs():
    conn = get_conn()
    cur = conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
    cur.execute("SELECT id, name, file_path, file_size as size, content as text, to_char(upload_time,'YYYY-MM-DD HH24:MI:SS') as upload_time FROM bot_docs ORDER BY upload_time DESC")
    rows = cur.fetchall()
    conn.close()
    return [dict(r) for r in rows]

def db_add_doc(doc_id, name, file_path, size, text):
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("""INSERT INTO bot_docs (id, name, file_path, file_size, content, upload_time)
                   VALUES (%s, %s, %s, %s, %s, %s)""",
                (doc_id, name, file_path, size, text, now))
    conn.commit()
    conn.close()

def db_delete_doc(doc_id):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("SELECT file_path FROM bot_docs WHERE id=%s", (doc_id,))
    row = cur.fetchone()
    cur.execute("DELETE FROM bot_docs WHERE id=%s", (doc_id,))
    affected = cur.rowcount
    conn.commit()
    conn.close()
    return affected > 0, (row[0] if row else None)


# ============================================================
# 知识库FAQ
# ============================================================
def db_load_kb():
    conn = get_conn()
    cur = conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
    cur.execute("SELECT id, category, question, keywords, answer FROM bot_kb_faq ORDER BY id")
    rows = cur.fetchall()
    conn.close()
    faq = []
    for r in rows:
        r = dict(r)
        # keywords存为JSON文本，解析回列表
        try:
            r["keywords"] = json.loads(r["keywords"]) if r["keywords"] else []
        except:
            r["keywords"] = []
        faq.append(r)
    return {"meta": {"name": "AI任务中台", "version": "2.0"}, "faq": faq}

def db_save_kb(kb_data):
    """全量覆盖知识库"""
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("DELETE FROM bot_kb_faq")
    for item in kb_data.get("faq", []):
        cur.execute("""INSERT INTO bot_kb_faq (id, category, question, keywords, answer)
                       VALUES (%s, %s, %s, %s, %s)""",
                    (item["id"], item.get("category",""), item.get("question",""),
                     json.dumps(item.get("keywords",[]), ensure_ascii=False), item.get("answer","")))
    conn.commit()
    conn.close()

def db_add_kb_item(item):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("""INSERT INTO bot_kb_faq (id, category, question, keywords, answer)
                   VALUES (%s, %s, %s, %s, %s)""",
                (item["id"], item.get("category",""), item.get("question",""),
                 json.dumps(item.get("keywords",[]), ensure_ascii=False), item.get("answer","")))
    conn.commit()
    conn.close()

def db_update_kb_item(kid, data):
    conn = get_conn()
    cur = conn.cursor()
    sets = []
    vals = []
    if "question" in data:
        sets.append("question=%s"); vals.append(data["question"].strip())
    if "answer" in data:
        sets.append("answer=%s"); vals.append(data["answer"].strip())
    if "category" in data:
        sets.append("category=%s"); vals.append(data["category"].strip())
    if "keywords" in data:
        kws = data["keywords"]
        if isinstance(kws, str):
            import re
            kws = [k.strip() for k in re.split(r"[，,、\n]", kws) if k.strip()]
        sets.append("keywords=%s"); vals.append(json.dumps(kws, ensure_ascii=False))
    if not sets:
        return False
    vals.append(kid)
    cur.execute(f"UPDATE bot_kb_faq SET {','.join(sets)} WHERE id=%s", vals)
    affected = cur.rowcount
    conn.commit()
    conn.close()
    return affected > 0

def db_delete_kb_item(kid):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("DELETE FROM bot_kb_faq WHERE id=%s", (kid,))
    affected = cur.rowcount
    conn.commit()
    conn.close()
    return affected > 0
