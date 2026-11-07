import json
import os
import re
import shutil
import sqlite3
import sys
import threading
import time
import uuid
from concurrent.futures import ThreadPoolExecutor, as_completed

import requests
from flask import Flask, jsonify, request, send_from_directory

if getattr(sys, "frozen", False):
    BASE_DIR = os.path.dirname(sys.executable)
    STATIC_DIR = os.path.join(sys._MEIPASS, "static")
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    STATIC_DIR = os.path.join(BASE_DIR, "static")
DB_PATH = os.path.join(BASE_DIR, "data.db")
CONFIG_PATH = os.path.join(BASE_DIR, "config.json")
BACKUP_DIR = os.path.join(BASE_DIR, "backups")

CONFIG_KEYS = ("provider", "api_base", "api_key", "model")
DEFAULT_CONFIG = {k: "" for k in CONFIG_KEYS}

app = Flask(__name__, static_folder=STATIC_DIR)


def load_config_file():
    if not os.path.exists(CONFIG_PATH):
        return {}
    with open(CONFIG_PATH, "r", encoding="utf-8") as f:
        return json.load(f)


def load_models():
    raw = load_config_file()
    if "models" in raw:
        return raw["models"], raw.get("active_id")
    if raw.get("api_base"):
        old = {k: raw.get(k) or "" for k in CONFIG_KEYS}
        old["id"] = "default"
        old["name"] = old["provider"] or "模型"
        return [old], "default"
    return [], None


def save_models(models, active_id):
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump({"models": models, "active_id": active_id}, f, ensure_ascii=False, indent=2)


def load_config():
    models, active = load_models()
    current = next((m for m in models if m.get("id") == active), models[0] if models else {})
    return {**DEFAULT_CONFIG, **{k: current.get(k) or "" for k in CONFIG_KEYS}}


def backup_db(tag=""):
    if not os.path.exists(DB_PATH):
        return None
    os.makedirs(BACKUP_DIR, exist_ok=True)
    name = "data_" + time.strftime("%Y%m%d_%H%M%S") + ("_" + tag if tag else "") + ".db"
    shutil.copy2(DB_PATH, os.path.join(BACKUP_DIR, name))
    files = [os.path.join(BACKUP_DIR, f) for f in os.listdir(BACKUP_DIR) if f.endswith(".db")]
    files.sort(key=os.path.getmtime)
    for old in files[:-10]:
        os.remove(old)
    return name


def db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    with db() as conn:
        conn.execute(
            """CREATE TABLE IF NOT EXISTS banks(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT UNIQUE NOT NULL,
                created_at REAL
            )"""
        )
        conn.execute(
            """CREATE TABLE IF NOT EXISTS questions(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                type TEXT NOT NULL,
                question TEXT NOT NULL,
                options TEXT,
                answer TEXT,
                analysis TEXT,
                source TEXT DEFAULT 'import',
                created_at REAL
            )"""
        )
        conn.execute(
            """CREATE TABLE IF NOT EXISTS records(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                qid INTEGER NOT NULL,
                correct INTEGER NOT NULL,
                answered_at REAL
            )"""
        )
        conn.execute(
            """CREATE TABLE IF NOT EXISTS exams(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                bank_id INTEGER NOT NULL,
                score REAL,
                total REAL,
                pass_line REAL,
                answered INTEGER,
                correct INTEGER,
                duration INTEGER,
                detail TEXT,
                created_at REAL
            )"""
        )
        conn.execute(
            """CREATE TABLE IF NOT EXISTS favorites(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                qid INTEGER UNIQUE NOT NULL,
                created_at REAL
            )"""
        )
        cols = [r[1] for r in conn.execute("PRAGMA table_info(questions)")]
        if "bank_id" not in cols:
            conn.execute("ALTER TABLE questions ADD COLUMN bank_id INTEGER")
            conn.execute(
                "INSERT OR IGNORE INTO banks(id,name,created_at) VALUES(1,'默认题库',?)",
                (time.time(),),
            )
            conn.execute("UPDATE questions SET bank_id=1 WHERE bank_id IS NULL")
        if "kind" not in cols:
            conn.execute("ALTER TABLE questions ADD COLUMN kind TEXT")
        old = conn.execute(
            "SELECT id,options,answer FROM questions WHERE kind IS NULL OR kind=''"
        ).fetchall()
        for r in old:
            opts = json.loads(r["options"]) if r["options"] else None
            conn.execute(
                "UPDATE questions SET kind=? WHERE id=?",
                (infer_kind(opts, r["answer"]), r["id"]),
            )


def get_or_create_bank(conn, name):
    name = (name or "").strip() or "未命名题库"
    row = conn.execute("SELECT id FROM banks WHERE name=?", (name,)).fetchone()
    if row:
        return row["id"]
    cur = conn.execute("INSERT INTO banks(name,created_at) VALUES(?,?)", (name, time.time()))
    return cur.lastrowid


def row_to_q(row):
    return {
        "id": row["id"],
        "type": row["type"],
        "kind": row["kind"],
        "question": row["question"],
        "options": json.loads(row["options"]) if row["options"] else None,
        "answer": row["answer"],
        "analysis": row["analysis"],
    }


def extract_text_from_file(path, filename):
    ext = os.path.splitext(filename)[1].lower()
    if ext == ".docx":
        from docx import Document

        doc = Document(path)
        parts = [p.text for p in doc.paragraphs if p.text.strip()]
        for table in doc.tables:
            for row in table.rows:
                cells = [c.text.strip() for c in row.cells if c.text.strip()]
                if cells:
                    parts.append(" | ".join(cells))
        return "\n".join(parts)
    if ext == ".pdf":
        import fitz

        with fitz.open(path) as doc:
            return "\n".join(page.get_text() for page in doc)
    if ext == ".xlsx":
        from openpyxl import load_workbook

        wb = load_workbook(path, read_only=True, data_only=True)
        parts = []
        for ws in wb.worksheets:
            for row in ws.iter_rows(values_only=True):
                cells = [str(v).strip() for v in row if v is not None and str(v).strip()]
                if cells:
                    parts.append(" | ".join(cells))
        wb.close()
        return "\n".join(parts)
    if ext in (".txt", ".md", ".csv"):
        with open(path, "r", encoding="utf-8-sig", errors="ignore") as f:
            return f.read()
    raise ValueError(f"不支持的文件类型: {ext}（支持 .docx / .xlsx / .csv / .pdf / .txt）")


def _table_columns(header):
    cols = {"options": {}}
    for i, name in enumerate(header):
        letter = re.search(r"[A-Za-z]", name)
        if name.startswith("选项") and letter:
            cols["options"][i] = letter.group(0).upper()
        elif "题干" in name or "试题内容" in name or "题目内容" in name:
            cols.setdefault("question", i)
        elif "答案" in name:
            cols.setdefault("answer", i)
        elif "题目类型" in name or "题型" in name:
            cols.setdefault("kind", i)
        elif "鉴定点" in name or "知识点" in name or "分类" in name:
            cols.setdefault("type", i)
        elif "解析" in name:
            cols.setdefault("analysis", i)
    if "question" not in cols or "answer" not in cols:
        return None
    return cols


def _row_to_question(cells, cols):
    def cell(key):
        i = cols.get(key)
        return cells[i] if i is not None and i < len(cells) else ""

    question = cell("question")
    if not question:
        return None
    options = {cols["options"][i]: cells[i] for i in cols["options"] if i < len(cells) and cells[i]}
    return {
        "type": cell("type") or "未分类",
        "kind": cell("kind"),
        "question": question,
        "options": options or None,
        "answer": cell("answer"),
        "analysis": cell("analysis"),
    }


def parse_question_table(path):
    from openpyxl import load_workbook

    wb = load_workbook(path, read_only=True, data_only=True)
    questions = []
    try:
        for ws in wb.worksheets:
            cols = None
            for row in ws.iter_rows(values_only=True):
                cells = [str(v).strip() if v is not None else "" for v in row]
                if not any(cells):
                    continue
                if cols is None:
                    cols = _table_columns(cells)
                    continue
                item = _row_to_question(cells, cols)
                q = clean_question(item) if item else None
                if q:
                    questions.append(q)
    finally:
        wb.close()
    return questions


QNUM_RE = re.compile(r"^\s*(?:第\s*)?(?:\d{1,3}|[一二三四五六七八九十]{1,3})\s*[.、．)）]\s*")
ANS_RE = re.compile(r"^(?:答案|正确答案|参考答案|标准答案)\s*[：:]\s*(.*)$")
ANA_RE = re.compile(r"^(?:解析|试题解析)\s*[：:]\s*(.*)$")
OPT_SPLIT = re.compile(r"([A-F])\s*[.、．)）]\s*")
TRUE_FALSE = ("正确", "错误", "对", "错", "√", "×", "是", "否")


def _split_option_line(line, allow_single):
    ms = list(OPT_SPLIT.finditer(line))
    if not ms:
        return line, {}
    letters = [m.group(1) for m in ms]
    if len(letters) == 1:
        if letters[0] != "A" and not allow_single:
            return line, {}
    elif letters != [chr(65 + i) for i in range(len(letters))]:
        return line, {}
    options = {}
    for i, m in enumerate(ms):
        end = ms[i + 1].start() if i + 1 < len(ms) else len(line)
        options[m.group(1)] = line[m.end():end].strip()
    return line[:ms[0].start()].strip(), options


def _normalize_answer(raw, options):
    raw = raw.strip()
    if not raw:
        return None
    if re.fullmatch(r"[A-Fa-f、,，;；\s]+", raw):
        letters = sorted({c.upper() for c in re.findall(r"[A-Fa-f]", raw)})
        return "".join(letters) or None
    for k, v in options.items():
        if raw == v or raw in v or v in raw:
            return k
    return raw


def _question_kind(options, answer):
    if not options:
        return "填空题"
    if re.fullmatch(r"[A-Fa-f]+", answer or "") and len(answer) > 1:
        return "多选题"
    if len(options) == 2 and all(v in TRUE_FALSE for v in options.values()):
        return "判断题"
    return "单选题"


def _parse_question_block(block):
    lines = [l.strip() for l in block.splitlines() if l.strip()]
    if not lines:
        return None
    qtext, options, answer, analysis = [], {}, None, None
    for i, line in enumerate(lines):
        m = ANS_RE.match(line)
        if m:
            answer = m.group(1)
            continue
        m = ANA_RE.match(line)
        if m:
            analysis = m.group(1)
            continue
        if i == 0:
            line = QNUM_RE.sub("", line)
        prefix, opts = _split_option_line(line, bool(options))
        if opts:
            options.update(opts)
            if prefix:
                qtext.append(prefix)
        elif options:
            continue
        else:
            qtext.append(line)
    question = " ".join(qtext).strip()
    answer = _normalize_answer(answer or "", options)
    if not question or not answer:
        return None
    return clean_question({
        "type": _question_kind(options, answer),
        "question": question,
        "options": options or None,
        "answer": answer,
        "analysis": analysis,
    })


def parse_plain_questions(text):
    blocks, cur = [], []
    for line in text.splitlines():
        if QNUM_RE.match(line.strip()) and any(l.strip() for l in cur):
            blocks.append("\n".join(cur))
            cur = []
        cur.append(line)
    if any(l.strip() for l in cur):
        blocks.append("\n".join(cur))
    questions, leftover = [], []
    for block in blocks:
        lines = block.splitlines()
        first = next((i for i, l in enumerate(lines) if QNUM_RE.match(l.strip())), 0)
        q = _parse_question_block("\n".join(lines[first:]))
        if q:
            questions.append(q)
        elif block.strip():
            leftover.append(block.strip())
    return questions, leftover


def chat(cfg, messages, timeout=180):
    resp = requests.post(
        cfg["api_base"].rstrip("/") + "/chat/completions",
        headers={"Authorization": f"Bearer {cfg['api_key']}"},
        json={
            "model": cfg["model"],
            "messages": messages,
            "temperature": 0.2,
            "max_tokens": 8192,
        },
        timeout=timeout,
    )
    resp.raise_for_status()
    return resp.json()["choices"][0]["message"]["content"]


def parse_json_array(text):
    text = re.sub(r"^```(json)?|```$", "", text.strip(), flags=re.M).strip()
    start, end = text.find("["), text.rfind("]")
    if start == -1 or end == -1:
        raise ValueError("返回内容中没有找到 JSON 数组：" + text[:200])
    return json.loads(text[start:end + 1])


def normalize_options(options):
    if isinstance(options, list):
        return {chr(65 + i): str(v) for i, v in enumerate(options)} if options else None
    if isinstance(options, dict):
        return {str(k).strip().upper()[:1]: str(v) for k, v in options.items()} or None
    return None


def infer_kind(options, answer):
    if not options:
        return "填空题"
    if len(options) == 2 and all(v in TRUE_FALSE for v in options.values()):
        return "判断题"
    if re.fullmatch(r"[A-Fa-f]{2,}", answer or ""):
        return "多选题"
    return "单选题"


def clean_question(item):
    question = str(item.get("question", "")).strip()
    if not question:
        return None
    options = normalize_options(item.get("options"))
    answer = str(item.get("answer") or "").strip() or None
    return {
        "type": str(item.get("type", "未分类")).strip() or "未分类",
        "kind": str(item.get("kind", "")).strip() or infer_kind(options, answer),
        "question": question,
        "options": options,
        "answer": answer,
        "analysis": str(item.get("analysis") or "").strip() or None,
    }


EXTRACT_PROMPT = """你是题目整理助手。下面是从复习文档中提取的原始文本，可能包含多道题目（也可能有答案混在一起）。
请把每道题目完整提取出来，按"知识点/题型"分类（例如：选择-计算机网络、填空-数据结构、编程题-循环 等，分类名要简洁统一）。

严格只输出 JSON 数组，不要输出任何其他文字：
[{"type":"分类名","question":"题干(完整,含题号可去掉)","options":{"A":"...","B":"..."} 或 null,"answer":"正确答案(选择题填字母)"}]

要求：
- 非选择题 options 一律为 null
- 原文给出了答案就原样使用，绝不修改；原文没有答案才自己解答
- 不要生成解析
- 不要遗漏题目，不要编造原文没有的题目

原始文本：
"""


def pack_blocks(blocks, limit=2500):
    chunks, buf, size = [], [], 0
    for block in blocks:
        block = block.strip()
        if not block:
            continue
        if len(block) > limit:
            if buf:
                chunks.append("\n".join(buf))
                buf, size = [], 0
            chunks.extend(block[i:i + limit] for i in range(0, len(block), limit))
            continue
        if size + len(block) + 1 > limit and buf:
            chunks.append("\n".join(buf))
            buf, size = [], 0
        buf.append(block)
        size += len(block) + 1
    if buf:
        chunks.append("\n".join(buf))
    return chunks


def count_types(questions):
    types = {}
    for q in questions:
        key = q.get("kind") or "未分类"
        types[key] = types.get(key, 0) + 1
    return types


def normalize_question(text):
    return re.sub(r"\s+", "", str(text)).strip()


def save_questions(bank_name, questions):
    with db() as conn:
        bank_id = get_or_create_bank(conn, bank_name)
        inserted = insert_questions(conn, questions, bank_id)
    return bank_id, inserted, len(questions) - inserted


def extract_chunk(cfg, chunk):
    last = None
    for i in range(3):
        try:
            content = chat(cfg, [{"role": "user", "content": EXTRACT_PROMPT + chunk}])
            return parse_json_array(content)
        except Exception as e:
            last = e
            time.sleep(2 + i * 3)
    raise last


TASKS = {}


def start_task(bank_name, local, chunks, cfg):
    task_id = uuid.uuid4().hex
    if local:
        bank_id, inserted, skipped = save_questions(bank_name, local)
    else:
        bank_id, inserted, skipped = None, 0, 0
    task = {
        "done": 0,
        "total": len(chunks),
        "count": inserted,
        "skipped": skipped,
        "types": count_types(local),
        "errors": 0,
        "finished": False,
        "bank_id": bank_id,
    }
    TASKS[task_id] = task

    def run():
        try:
            with ThreadPoolExecutor(max_workers=3) as pool:
                futures = [pool.submit(extract_chunk, cfg, c) for c in chunks]
                for fut in as_completed(futures):
                    try:
                        items = fut.result()
                        questions = [q for q in (clean_question(i) for i in items) if q]
                    except Exception:
                        questions = []
                        task["errors"] += 1
                    if questions:
                        bid, ins, skip = save_questions(bank_name, questions)
                        task["bank_id"] = bid or task["bank_id"]
                        task["count"] += ins
                        task["skipped"] += skip
                        for t, c in count_types(questions).items():
                            task["types"][t] = task["types"].get(t, 0) + c
                    task["done"] += 1
        finally:
            task["finished"] = True

    threading.Thread(target=run, daemon=True).start()
    return task_id


def insert_questions(conn, questions, bank_id):
    existing = {
        normalize_question(r["question"])
        for r in conn.execute("SELECT question FROM questions WHERE bank_id=?", (bank_id,))
    }
    inserted = 0
    for q in questions:
        key = normalize_question(q["question"])
        if key in existing:
            continue
        conn.execute(
            """INSERT INTO questions(type,kind,question,options,answer,analysis,created_at,bank_id)
               VALUES(?,?,?,?,?,?,?,?)""",
            (
                q["type"],
                q["kind"],
                q["question"],
                json.dumps(q["options"], ensure_ascii=False) if q["options"] else None,
                q["answer"],
                q["analysis"],
                time.time(),
                bank_id,
            ),
        )
        existing.add(key)
        inserted += 1
    return inserted


def mask_key(key):
    return key[:6] + "***" if key else ""


@app.route("/")
def index():
    return send_from_directory(app.static_folder, "index.html")


@app.route("/api/config", methods=["GET"])
def get_config():
    cfg = load_config()
    return jsonify({**cfg, "api_key": mask_key(cfg["api_key"])})


@app.route("/api/config", methods=["POST"])
def set_config():
    body = request.json or {}
    api_base = str(body.get("api_base") or "").strip()
    model = str(body.get("model") or "").strip()
    api_key = str(body.get("api_key") or "").strip()
    if not api_base or not model:
        return jsonify({"error": "接口地址和模型名称都要填写"}), 400
    provider = str(body.get("provider") or "").strip()
    name = str(body.get("name") or provider or "模型").strip()
    models, active = load_models()
    same = next((m for m in models if m.get("api_base") == api_base and m.get("model") == model), None)
    if same:
        if api_key:
            same["api_key"] = api_key
        same["provider"] = provider or same.get("provider", "")
        same["name"] = name or same.get("name", "")
        active = same["id"]
    else:
        if not api_key:
            return jsonify({"error": "新模型需要填写 API Key"}), 400
        item = {
            "id": uuid.uuid4().hex,
            "provider": provider,
            "name": name,
            "api_base": api_base,
            "api_key": api_key,
            "model": model,
        }
        models.append(item)
        active = item["id"]
    save_models(models, active)
    return jsonify({"ok": True})


@app.route("/api/saved-models")
def saved_models():
    models, active = load_models()
    return jsonify([
        {
            "id": m.get("id", ""),
            "provider": m.get("provider", ""),
            "name": m.get("name") or m.get("provider") or "模型",
            "api_base": m.get("api_base", ""),
            "model": m.get("model", ""),
            "key": mask_key(m.get("api_key", "")),
            "active": m.get("id") == active,
        }
        for m in models
    ])


@app.route("/api/saved-models/<mid>", methods=["DELETE"])
def delete_model(mid):
    models, active = load_models()
    models = [m for m in models if m.get("id") != mid]
    if active == mid:
        active = models[0]["id"] if models else None
    save_models(models, active)
    return jsonify({"ok": True})


@app.route("/api/saved-models/activate", methods=["POST"])
def activate_model():
    mid = str((request.json or {}).get("id") or "")
    models, _ = load_models()
    if not any(m.get("id") == mid for m in models):
        return jsonify({"error": "模型不存在"}), 404
    save_models(models, mid)
    return jsonify({"ok": True})


@app.route("/api/backup", methods=["POST"])
def make_backup():
    return jsonify({"ok": True, "name": backup_db()})


@app.route("/api/backups")
def list_backups():
    items = []
    if os.path.isdir(BACKUP_DIR):
        for f in os.listdir(BACKUP_DIR):
            if not f.endswith(".db"):
                continue
            p = os.path.join(BACKUP_DIR, f)
            items.append({"name": f, "size": os.path.getsize(p), "time": os.path.getmtime(p)})
        items.sort(key=lambda x: x["time"], reverse=True)
    return jsonify(items)


@app.route("/api/backups/<name>/restore", methods=["POST"])
def restore_backup(name):
    path = os.path.join(BACKUP_DIR, os.path.basename(name))
    if not os.path.exists(path):
        return jsonify({"error": "备份不存在"}), 404
    backup_db("before-restore")
    src = sqlite3.connect(path)
    dst = sqlite3.connect(DB_PATH)
    try:
        src.backup(dst)
    finally:
        src.close()
        dst.close()
    init_db()
    return jsonify({"ok": True})


@app.route("/api/records/clear", methods=["POST"])
def clear_records():
    with db() as conn:
        conn.execute("DELETE FROM records")
        conn.execute("DELETE FROM exams")
    return jsonify({"ok": True})


@app.route("/api/config/test", methods=["POST"])
def test_config():
    cfg = load_config()
    if not cfg["api_key"]:
        return jsonify({"ok": False, "error": "还没有配置模型，先到「设置 → 添加模型」添加一个"}), 400
    try:
        reply = chat(cfg, [{"role": "user", "content": "只回复两个字：正常"}], timeout=30)
        return jsonify({"ok": True, "reply": reply.strip()[:50]})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)[:300]})


@app.route("/api/models", methods=["POST"])
def list_models():
    body = request.json or {}
    api_base = str(body.get("api_base") or "").strip()
    api_key = str(body.get("api_key") or "").strip()
    if not api_key or api_key.endswith("***"):
        return jsonify({"error": "请先粘贴完整的 API Key"}), 400
    if not api_base:
        return jsonify({"error": "缺少接口地址"}), 400
    try:
        r = requests.get(
            api_base.rstrip("/") + "/models",
            headers={"Authorization": f"Bearer {api_key}"},
            timeout=10,
        )
        if r.status_code in (401, 403):
            return jsonify({"error": "Key 无效或没有权限，检查是否复制完整"}), 401
        r.raise_for_status()
        models = sorted({m["id"] for m in r.json().get("data", []) if m.get("id")})
        if not models:
            return jsonify({"error": "平台没有返回模型列表"}), 404
        return jsonify({"ok": True, "models": models})
    except Exception as e:
        return jsonify({"error": f"获取失败：{str(e)[:200]}"}), 502


@app.route("/api/upload", methods=["POST"])
def upload():
    f = request.files.get("file")
    if not f:
        return jsonify({"error": "没有收到文件"}), 400
    path = os.path.join(BASE_DIR, "_upload_" + f.filename)
    f.save(path)
    try:
        questions = parse_question_table(path) if f.filename.lower().endswith(".xlsx") else []
        if questions:
            return jsonify({"text": "", "questions": questions})
        return jsonify({"text": extract_text_from_file(path, f.filename)})
    except Exception as e:
        return jsonify({"error": str(e)}), 400
    finally:
        try:
            os.remove(path)
        except OSError:
            pass


@app.route("/api/import", methods=["POST"])
def import_text():
    body = request.json or {}
    bank_name = body.get("bank_name")
    items = body.get("questions") or []
    questions = [q for q in (clean_question(i) for i in items) if q]
    if questions:
        bank_id, inserted, skipped = save_questions(bank_name, questions)
        return jsonify({
            "ok": True,
            "count": inserted,
            "skipped": skipped,
            "types": count_types(questions),
            "bank_id": bank_id,
        })

    text = str(body.get("text") or "").strip()
    if len(text) < 10:
        return jsonify({"error": "文本太短，先上传文件或粘贴题目"}), 400
    local, leftover = parse_plain_questions(text)
    if not leftover:
        if not local:
            return jsonify({"error": "没有识别出题目，检查文档内容"}), 400
        bank_id, inserted, skipped = save_questions(bank_name, local)
        return jsonify({
            "ok": True,
            "count": inserted,
            "skipped": skipped,
            "types": count_types(local),
            "bank_id": bank_id,
        })

    cfg = load_config()
    if not cfg["api_key"]:
        if not local:
            return jsonify({"error": "这些题目本地识别不出来，需要配置模型才能继续"}), 400
        bank_id, inserted, skipped = save_questions(bank_name, local)
        return jsonify({
            "ok": True,
            "count": inserted,
            "skipped": skipped,
            "types": count_types(local),
            "bank_id": bank_id,
            "pending": len(leftover),
        })

    task_id = start_task(bank_name, local, pack_blocks(leftover), cfg)
    return jsonify({"ok": True, "task_id": task_id, "local_count": len(local)})


@app.route("/api/import/progress/<task_id>")
def import_progress(task_id):
    task = TASKS.get(task_id)
    if not task:
        return jsonify({"error": "任务不存在或已结束"}), 404
    return jsonify(task)


WRONG_COND = """EXISTS(SELECT 1 FROM records r WHERE r.qid=q.id AND r.correct=0)
                   AND NOT EXISTS(
                     SELECT 1 FROM records r2 WHERE r2.qid=q.id AND r2.id=(
                       SELECT id FROM records WHERE qid=q.id ORDER BY id DESC LIMIT 1
                     ) AND r2.correct=1)"""


@app.route("/api/banks")
def list_banks():
    with db() as conn:
        rows = conn.execute(
            """SELECT b.id, b.name,
                 (SELECT COUNT(*) FROM questions q WHERE q.bank_id=b.id) qcount,
                 (SELECT COUNT(DISTINCT q.kind) FROM questions q WHERE q.bank_id=b.id) tcount,
                 (SELECT COUNT(*) FROM questions q WHERE q.bank_id=b.id AND """ + WRONG_COND + """) wcount
               FROM banks b ORDER BY b.id"""
        ).fetchall()
    return jsonify([dict(r) for r in rows])


@app.route("/api/banks/<int:bank_id>", methods=["PUT"])
def rename_bank(bank_id):
    body = request.json or {}
    name = str(body.get("name") or "").strip()
    if not name:
        return jsonify({"error": "题库名不能为空"}), 400
    with db() as conn:
        if conn.execute("SELECT id FROM banks WHERE name=?", (name,)).fetchone():
            return jsonify({"error": "已经有同名题库"}), 400
        conn.execute("UPDATE banks SET name=? WHERE id=?", (name, bank_id))
    return jsonify({"ok": True})


@app.route("/api/banks/<int:bank_id>", methods=["DELETE"])
def delete_bank(bank_id):
    with db() as conn:
        conn.execute(
            "DELETE FROM records WHERE qid IN (SELECT id FROM questions WHERE bank_id=?)",
            (bank_id,),
        )
        conn.execute("DELETE FROM questions WHERE bank_id=?", (bank_id,))
        conn.execute("DELETE FROM banks WHERE id=?", (bank_id,))
    return jsonify({"ok": True})


@app.route("/api/bank/<int:bank_id>/last")
def last_results(bank_id):
    with db() as conn:
        rows = conn.execute(
            """SELECT r.qid, r.correct FROM records r
               JOIN (SELECT qid, MAX(id) mid FROM records GROUP BY qid) m ON m.mid=r.id
               JOIN questions q ON q.id=r.qid AND q.bank_id=?""",
            (bank_id,),
        ).fetchall()
    return jsonify({str(r["qid"]): r["correct"] for r in rows})


@app.route("/api/favorites")
def list_favorites():
    bank_id = request.args.get("bank_id", type=int)
    if not bank_id:
        return jsonify({"error": "缺少 bank_id"}), 400
    with db() as conn:
        rows = conn.execute(
            """SELECT q.* FROM questions q JOIN favorites f ON f.qid=q.id
               WHERE q.bank_id=? ORDER BY f.id DESC""",
            (bank_id,),
        ).fetchall()
    return jsonify([row_to_q(r) for r in rows])


@app.route("/api/favorites/ids")
def favorite_ids():
    with db() as conn:
        rows = conn.execute("SELECT qid FROM favorites").fetchall()
    return jsonify([r["qid"] for r in rows])


@app.route("/api/favorite/<int:qid>", methods=["POST"])
def add_favorite(qid):
    with db() as conn:
        conn.execute(
            "INSERT OR IGNORE INTO favorites(qid,created_at) VALUES(?,?)",
            (qid, time.time()),
        )
    return jsonify({"ok": True})


@app.route("/api/favorite/<int:qid>", methods=["DELETE"])
def remove_favorite(qid):
    with db() as conn:
        conn.execute("DELETE FROM favorites WHERE qid=?", (qid,))
    return jsonify({"ok": True})


@app.route("/api/practice")
def practice():
    bank_id = request.args.get("bank_id", type=int)
    if not bank_id:
        return jsonify({"error": "缺少 bank_id"}), 400
    qtype = request.args.get("type", "__all__")
    count = int(request.args.get("count", 20))
    order = "ORDER BY id" if request.args.get("order") == "seq" else "ORDER BY RANDOM()"
    limit = " LIMIT ?" if count > 0 else ""
    with db() as conn:
        if qtype == "__wrong__":
            sql = f"SELECT q.* FROM questions q WHERE q.bank_id=? AND {WRONG_COND} {order}{limit}"
            params = [bank_id]
        elif qtype == "__all__":
            sql = f"SELECT * FROM questions WHERE bank_id=? {order}{limit}"
            params = [bank_id]
        else:
            sql = f"SELECT * FROM questions WHERE bank_id=? AND kind=? {order}{limit}"
            params = [bank_id, qtype]
        if limit:
            params.append(count)
        rows = conn.execute(sql, params).fetchall()
    return jsonify([row_to_q(r) for r in rows])


EXPLAIN_PROMPT = """你是考试辅导老师。请用中文简洁讲解这道题。

题干：{question}
选项：
{options}
正确答案：{answer}
学生选择：{chosen}

要求：
1. 说明正确答案为什么对
2. 如果学生选错了，指出他选的为什么错
3. 不要复述题目原文，控制在 200 字以内
"""


@app.route("/api/explain", methods=["POST"])
def explain_question():
    body = request.json or {}
    qid = body.get("qid")
    chosen = body.get("chosen") or []
    if not qid:
        return jsonify({"error": "缺少题目 id"}), 400
    with db() as conn:
        row = conn.execute("SELECT * FROM questions WHERE id=?", (qid,)).fetchone()
    if not row:
        return jsonify({"error": "题目不存在"}), 404
    q = row_to_q(row)
    if q["analysis"]:
        return jsonify({"ok": True, "text": q["analysis"]})
    cfg = load_config()
    if not cfg["api_key"]:
        return jsonify({"error": "还没有配置模型，先到「设置 → 添加模型」添加一个"}), 400
    opts = "\n".join(f"{k}. {v}" for k, v in (q["options"] or {}).items())
    prompt = EXPLAIN_PROMPT.format(
        question=q["question"],
        options=opts or "（无选项）",
        answer=q["answer"] or "（无）",
        chosen="".join(chosen) or "未作答",
    )
    try:
        text = chat(cfg, [{"role": "user", "content": prompt}], timeout=60).strip()
    except Exception as e:
        return jsonify({"error": f"讲解失败：{str(e)[:200]}"}), 502
    with db() as conn:
        conn.execute("UPDATE questions SET analysis=? WHERE id=?", (text, qid))
    return jsonify({"ok": True, "text": text})


@app.route("/api/exam", methods=["POST"])
def exam_pick():
    body = request.json or {}
    bank_id = body.get("bank_id")
    counts = body.get("types") or {}
    if not bank_id:
        return jsonify({"error": "缺少 bank_id"}), 400
    picked = []
    with db() as conn:
        for kind, num in counts.items():
            num = int(num or 0)
            if num <= 0:
                continue
            rows = conn.execute(
                "SELECT * FROM questions WHERE bank_id=? AND kind=? ORDER BY RANDOM() LIMIT ?",
                (bank_id, kind, num),
            ).fetchall()
            picked.extend(row_to_q(r) for r in rows)
    return jsonify({"questions": picked, "total": len(picked)})


@app.route("/api/exam/records/<int:rid>")
def exam_record_detail(rid):
    with db() as conn:
        r = conn.execute("SELECT * FROM exams WHERE id=?", (rid,)).fetchone()
        if not r:
            return jsonify({"error": "记录不存在"}), 404
        detail = json.loads(r["detail"]) if r["detail"] else {}
        stats = detail.get("stats", detail)
        wrong = detail.get("wrong") or []
        wrong_items = []
        ids = [w.get("qid") for w in wrong if w.get("qid")]
        if ids:
            rows = conn.execute(
                "SELECT * FROM questions WHERE id IN ({})".format(",".join("?" * len(ids))),
                ids,
            ).fetchall()
            qmap = {x["id"]: row_to_q(x) for x in rows}
            for w in wrong:
                q = qmap.get(w.get("qid"))
                if q:
                    wrong_items.append({**q, "chosen": w.get("chosen") or []})
    return jsonify({
        "id": r["id"],
        "score": r["score"],
        "total": r["total"],
        "pass_line": r["pass_line"],
        "answered": r["answered"],
        "correct": r["correct"],
        "duration": r["duration"],
        "stats": stats,
        "wrong": wrong_items,
        "created_at": r["created_at"],
    })


@app.route("/api/exam/records", methods=["POST"])
def save_exam_record():
    body = request.json or {}
    bank_id = body.get("bank_id")
    if not bank_id:
        return jsonify({"error": "缺少 bank_id"}), 400
    with db() as conn:
        conn.execute(
            """INSERT INTO exams(bank_id,score,total,pass_line,answered,correct,duration,detail,created_at)
               VALUES(?,?,?,?,?,?,?,?,?)""",
            (
                bank_id,
                body.get("score") or 0,
                body.get("total") or 0,
                body.get("pass") or 0,
                body.get("answered") or 0,
                body.get("correct") or 0,
                body.get("duration") or 0,
                json.dumps({
                    "stats": body.get("stats") or {},
                    "wrong": body.get("wrong") or [],
                }, ensure_ascii=False),
                time.time(),
            ),
        )
    return jsonify({"ok": True})


@app.route("/api/exam/records")
def exam_records():
    bank_id = request.args.get("bank_id", type=int)
    if not bank_id:
        return jsonify({"error": "缺少 bank_id"}), 400
    with db() as conn:
        rows = conn.execute(
            "SELECT * FROM exams WHERE bank_id=? ORDER BY id DESC LIMIT 20", (bank_id,)
        ).fetchall()
    return jsonify([
        {
            "id": r["id"],
            "score": r["score"],
            "total": r["total"],
            "pass_line": r["pass_line"],
            "answered": r["answered"],
            "correct": r["correct"],
            "duration": r["duration"],
            "stats": json.loads(r["detail"]) if r["detail"] else {},
            "created_at": r["created_at"],
        }
        for r in rows
    ])


@app.route("/api/record/batch", methods=["POST"])
def record_batch():
    items = (request.json or {}).get("items") or []
    with db() as conn:
        for item in items:
            qid = item.get("qid")
            if not qid:
                continue
            conn.execute(
                "INSERT INTO records(qid,correct,answered_at) VALUES(?,?,?)",
                (qid, 1 if item.get("correct") else 0, time.time()),
            )
    return jsonify({"ok": True})


@app.route("/api/record", methods=["POST"])
def record():
    body = request.json or {}
    qid = body.get("qid")
    if not qid:
        return jsonify({"error": "缺少 qid"}), 400
    with db() as conn:
        conn.execute(
            "INSERT INTO records(qid,correct,answered_at) VALUES(?,?,?)",
            (qid, 1 if body.get("correct") else 0, time.time()),
        )
    return jsonify({"ok": True})


@app.route("/api/stats")
def stats():
    bank_id = request.args.get("bank_id", type=int)
    if not bank_id:
        return jsonify({"error": "缺少 bank_id"}), 400
    with db() as conn:
        by_type = conn.execute(
            "SELECT kind, COUNT(*) total FROM questions WHERE bank_id=? GROUP BY kind",
            (bank_id,),
        ).fetchall()
        answered = conn.execute(
            """SELECT q.kind t, SUM(r.correct) right_cnt, COUNT(*) total_cnt
               FROM records r JOIN questions q ON q.id=r.qid
               WHERE q.bank_id=? GROUP BY q.kind""",
            (bank_id,),
        ).fetchall()
        wrong_rows = conn.execute(
            """SELECT * FROM questions q WHERE q.bank_id=? AND EXISTS(
                 SELECT 1 FROM records r WHERE r.qid=q.id AND r.correct=0)
               AND NOT EXISTS(
                 SELECT 1 FROM records r2 WHERE r2.qid=q.id AND r2.id=(
                   SELECT MAX(id) FROM records WHERE qid=q.id) AND r2.correct=1)"""
            , (bank_id,),
        ).fetchall()
        fav = conn.execute(
            "SELECT COUNT(*) FROM favorites f JOIN questions q ON q.id=f.qid WHERE q.bank_id=?",
            (bank_id,),
        ).fetchone()[0]
    accuracy = {r["t"]: {"right": r["right_cnt"], "total": r["total_cnt"]} for r in answered}
    return jsonify({
        "types": {(r["kind"] or "未分类"): r["total"] for r in by_type},
        "accuracy": accuracy,
        "wrong": [row_to_q(r) for r in wrong_rows],
        "fav": fav,
    })


init_db()
backup_db("start")

if __name__ == "__main__":
    print("http://127.0.0.1:5000")
    app.run(host="127.0.0.1", port=5000, debug=False)
