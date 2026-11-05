# -*- coding: utf-8 -*-
"""
课堂随机点名器（本地版）
- 双击「启动课堂点名器.bat」即用
- 导入花名册 Excel（.xlsx/.xls/.csv），自动识别「姓名」列和「班级」列
- 支持单人点名（大字滚动）、抽小组、不重复模式、按班级过滤
- 点名记录一键导出 TXT（含时间与班级）
- 全部数据保存在本产品文件夹 data/ 内，不联网。
"""
import json
import os
import random
import re
import socket
import threading
import time
import uuid
import webbrowser
from pathlib import Path

from flask import Flask, jsonify, request, render_template, send_file

BASE = Path(__file__).resolve().parent
DATA_DIR = BASE / "data"
REC_DIR = BASE / "点名记录"
TMP_DIR = DATA_DIR / "_uploads"
STATE_F = DATA_DIR / "state.json"
DATA_DIR.mkdir(exist_ok=True)
REC_DIR.mkdir(exist_ok=True)
TMP_DIR.mkdir(exist_ok=True)

DEFAULT_PORT = 8931

app = Flask(__name__, static_folder="static", template_folder="static")
app.config["MAX_CONTENT_LENGTH"] = 200 * 1024 * 1024

_NAME_PAT = re.compile(r"^[\u4e00-\u9fa5·‧•]{2,6}$")
_NAME_HEADER = re.compile(r"^(姓名|学生姓名|学生|name)$")
_BAD_HEADER = re.compile(r"(父|母|监护|联系|备注|联系)")
_CLASS_HEADER = re.compile(r"(班级|班别|行政班)")
_EXCLUDE_HEADER = re.compile(r"(证件|电话|号码|身份|地址|状态|银行|社保)")


def load_state() -> dict:
    if STATE_F.exists():
        try:
            return json.loads(STATE_F.read_text(encoding="utf-8"))
        except Exception:
            pass
    return {"students": [], "drawn": [], "source": ""}


def save_state(st: dict):
    tmp = STATE_F.with_suffix(".tmp")
    tmp.write_text(json.dumps(st, ensure_ascii=False), encoding="utf-8")
    os.replace(str(tmp), str(STATE_F))


_RESERVED = re.compile(r"^(con|prn|aux|nul|com[1-9]|lpt[1-9])(\..*)?$", re.I)


def safe_name(name: str) -> str:
    """清洗文件/目录名：剥路径分隔符、去首尾点号（阻断 .. 穿越）、拦 Windows 保留设备名。"""
    name = os.path.basename(name or "")
    cleaned = re.sub(r'[\\/:*?"<>|\r\n]+', "_", name).strip().strip(".")
    if not cleaned or _RESERVED.match(cleaned):
        cleaned = "_" + (cleaned or "file")
    return cleaned or "file.txt"


def contained(base: Path, p: Path) -> bool:
    """路径容器校验：拼接后的最终路径必须仍在 base 内（防御纵深）。"""
    try:
        return os.path.commonpath([str(base.resolve()), str(p.resolve())]) == str(base.resolve())
    except ValueError:
        return False


def read_table(path: Path):
    """返回 (headers, rows)，rows 为元组列表。xlsx/xls/csv 统一处理。"""
    suffix = path.suffix.lower()
    if suffix == ".csv":
        for enc in ("utf-8-sig", "gbk", "utf-8"):
            try:
                lines = path.read_text(encoding=enc).splitlines()
                break
            except Exception:
                continue
        else:
            raise RuntimeError("CSV 编码无法识别")
        rows = [next(__import__("csv").reader([ln])) for ln in lines if ln.strip()]
        # csv 长度补齐
        width = max(len(r) for r in rows) if rows else 0
        rows = [r + [""] * (width - len(r)) for r in rows]
        return rows[0] if rows else [], rows[1:] if rows else []
    if suffix == ".xls":
        import xlrd
        wb = xlrd.open_workbook(str(path))
        ws = wb.sheet_by_index(0)
        rows = [[ws.cell_value(r, c) for c in range(ws.ncols)] for r in range(ws.nrows)]
        return rows[0] if rows else [], rows[1:] if rows else []
    # xlsx
    import openpyxl
    wb = openpyxl.load_workbook(str(path), read_only=True, data_only=True)
    ws = wb.active
    headers, rows = [], []
    for i, row in enumerate(ws.iter_rows(values_only=True)):
        vals = ["" if v is None else str(v).strip() for v in row]
        if i == 0:
            headers = vals
        else:
            rows.append(vals)
        if len(rows) > 200000:
            break
    wb.close()
    return headers, rows


def guess_name_col(headers, rows=None):
    """按表头识别姓名列：优先精确『姓名』，排除父母/监护人/联系人/证件等。"""
    for i, h in enumerate(headers):
        h = str(h).strip()
        if h and _NAME_HEADER.match(h) and not _BAD_HEADER.search(h):
            return i
    for i, h in enumerate(headers):
        h = str(h).strip()
        if "姓名" in h and not _BAD_HEADER.search(h) and not _EXCLUDE_HEADER.search(h):
            return i
    # 再兜底：中文姓名格式占比最高的列
    if rows:
        best, best_n = None, 0
        for i in range(len(headers)):
            n = sum(1 for r in rows[:5000] if i < len(r) and _NAME_PAT.match(str(r[i]).strip()))
            if n > best_n:
                best, best_n = i, n
        if best_n:
            return best
    return None


def guess_class_col(headers):
    for i, h in enumerate(headers):
        h = str(h).strip()
        if h and _CLASS_HEADER.search(h) and not _EXCLUDE_HEADER.search(h):
            return i
    return None


def parse_students(headers, rows, name_col, class_col):
    """同名学生（真实名册存在）各自独立成条，靠 id 区分；
    只有整行完全相同的重复录入才合并。"""
    students, seen_rows = [], set()
    for r in rows:
        name = str(r[name_col]).strip() if name_col < len(r) else ""
        name = re.sub(r"\s+", "", name)
        if not name:
            continue
        # 表头行混进数据时跳过
        if name in ("姓名", "学生姓名"):
            continue
        key = tuple(str(c).strip() for c in r)
        if key in seen_rows:
            continue
        seen_rows.add(key)
        cls = str(r[class_col]).strip() if class_col is not None and class_col < len(r) else ""
        students.append({"id": len(students), "name": name, "cls": cls})
    return students


@app.get("/")
def index():
    return render_template("index.html")


@app.post("/api/import")
def api_import():
    f = request.files.get("file")
    if not f or not f.filename:
        return jsonify({"error": "未收到文件"}), 400
    fn = f"{uuid.uuid4().hex}_{safe_name(f.filename)}"
    p = TMP_DIR / fn
    f.save(str(p))
    try:
        headers, rows = read_table(p)
    except Exception as e:
        p.unlink(missing_ok=True)
        return jsonify({"error": f"解析失败：{e}"}), 400
    if not headers:
        return jsonify({"error": "表格为空"}), 400
    # 去掉整行为空的记录（学籍导出常有几万空行），控制回传体积
    rows = [r for r in rows if any(str(c).strip() for c in r)]
    return jsonify({
        "file": fn,
        "filename": safe_name(f.filename),
        "columns": [{"idx": i, "header": str(h).strip() or f"列{i+1}"} for i, h in enumerate(headers)],
        "name_col_guess": guess_name_col(headers, rows),
        "class_col_guess": guess_class_col(headers),
        "row_count": len(rows),
        "_headers": headers,
        "_rows": rows[:50000],
    })


@app.post("/api/apply")
def api_apply():
    data = request.get_json(force=True)
    headers = data.get("_headers") or []
    rows = data.get("_rows") or []
    name_col = data.get("name_col")
    class_col = data.get("class_col")
    if name_col is None or name_col >= len(headers):
        return jsonify({"error": "请选择姓名列"}), 400
    students = parse_students(headers, rows, int(name_col),
                              int(class_col) if class_col not in (None, "") else None)
    if not students:
        return jsonify({"error": "该列没有解析出中文姓名，请换一列试试"}), 400
    st = load_state()
    st["students"] = students
    st["drawn"] = []
    st["source"] = data.get("filename", "手动导入")
    save_state(st)
    cleanup_tmp()
    classes = sorted({s["cls"] for s in students if s["cls"]})
    return jsonify({"ok": True, "count": len(students), "classes": classes})


def cleanup_tmp():
    """清掉上传临时目录（名册属敏感数据，不留副本）。"""
    for f in TMP_DIR.iterdir():
        try:
            if f.is_file():
                f.unlink()
        except Exception:
            pass


@app.post("/api/paste")
def api_paste():
    text = (request.get_json(force=True) or {}).get("text", "")
    raw = re.split(r"[\s,，、;；]+", text)
    students, seen = [], set()
    for n in raw:
        n = re.sub(r"\s+", "", n)
        if _NAME_PAT.match(n) and n not in seen:
            seen.add(n)
            students.append({"id": len(students), "name": n, "cls": ""})
    if not students:
        return jsonify({"error": "没解析出姓名（用换行、空格或顿号分隔）"}), 400
    st = load_state()
    st["students"] = students
    st["drawn"] = []
    st["source"] = "手动粘贴"
    save_state(st)
    return jsonify({"ok": True, "count": len(students), "classes": []})


@app.get("/api/state")
def api_state():
    st = load_state()
    classes = sorted({s["cls"] for s in st["students"] if s["cls"]})
    return jsonify({"count": len(st["students"]), "drawn": st["drawn"],
                    "source": st.get("source", ""), "classes": classes})


@app.post("/api/draw")
def api_draw():
    data = request.get_json(force=True)
    n = max(1, min(int(data.get("n", 1)), 12))
    cls = data.get("cls") or ""
    unique = bool(data.get("unique", True))
    st = load_state()
    pool = [s for s in st["students"] if (not cls or s["cls"] == cls)]
    if unique:
        drawn_ids = {d["id"] for d in st["drawn"]}
        pool = [s for s in pool if s["id"] not in drawn_ids]
    if not pool:
        return jsonify({"error": "可选的人抽完了：请重置本轮或关闭『不重复』"}), 400
    n = min(n, len(pool))
    picked = random.sample(pool, n)
    now = time.strftime("%H:%M:%S")
    for s in picked:
        st["drawn"].append({"id": s["id"], "name": s["name"], "cls": s["cls"], "time": now})
    save_state(st)
    return jsonify({"picked": picked, "drawn": st["drawn"], "remain": len(st["students"]) - len(st["drawn"])})


@app.post("/api/reset")
def api_reset():
    st = load_state()
    st["drawn"] = []
    save_state(st)
    return jsonify({"ok": True})


@app.post("/api/undo")
def api_undo():
    st = load_state()
    if st["drawn"]:
        st["drawn"].pop()
        save_state(st)
    return jsonify({"ok": True, "drawn": st["drawn"]})


@app.post("/api/clear")
def api_clear():
    st = load_state()
    st["students"] = []
    st["drawn"] = []
    st["source"] = ""
    save_state(st)
    cleanup_tmp()
    return jsonify({"ok": True})


@app.post("/api/export")
def api_export():
    st = load_state()
    if not st["drawn"]:
        return jsonify({"error": "本轮还没有点名记录"}), 400
    fn = REC_DIR / f"点名记录_{time.strftime('%Y%m%d_%H%M%S')}.txt"
    lines = [f"课堂随机点名记录（来源：{st.get('source','')}，共 {len(st['drawn'])} 人次）",
             f"导出时间：{time.strftime('%Y-%m-%d %H:%M:%S')}", "-" * 40]
    for i, d in enumerate(st["drawn"], 1):
        lines.append(f"{i:>3}. {d['name']}" + (f"（{d['cls']}）" if d["cls"] else "") + f"  {d['time']}")
    fn.write_text("\n".join(lines), encoding="utf-8-sig")
    return jsonify({"filename": fn.name, "count": len(st["drawn"])})


@app.get("/api/open/<name>")
def api_open(name):
    p = REC_DIR / safe_name(name)
    if not contained(REC_DIR, p) or not p.is_file():
        return jsonify({"error": "不存在"}), 404
    return send_file(str(p), as_attachment=True, download_name=p.name)


def pick_port(start):
    for port in range(start, start + 30):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            try:
                s.bind(("127.0.0.1", port))
                return port
            except OSError:
                continue
    raise RuntimeError("端口都被占用")


if __name__ == "__main__":
    port = pick_port(DEFAULT_PORT)
    url = f"http://127.0.0.1:{port}"
    print(f"课堂随机点名器 已启动：{url}（本窗口最小化即可，关闭即退出）", flush=True)
    if not os.environ.get("ZC_NO_BROWSER"):
        threading.Timer(1.0, lambda: webbrowser.open(url)).start()
    app.run(host="127.0.0.1", port=port, debug=False)
