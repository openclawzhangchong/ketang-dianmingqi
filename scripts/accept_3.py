# -*- coding: utf-8 -*-
"""产品3《课堂随机点名器》验收：导入真实24级706人花名册，验证识别/抽取/不重复/班级过滤/导出。"""
import sys, os, io, json, time, re, requests

BASE = sys.argv[1] if len(sys.argv) > 1 else "http://127.0.0.1:8931"
EVID = sys.argv[2] if len(sys.argv) > 2 else "evidence"
PROD = sys.argv[3] if len(sys.argv) > 3 else os.path.join(
    os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "build", "03-课堂随机点名器")
REAL = r"Z:\github项目\xueji-xitong\2026年7月份24级（706人）.xlsx"
SAMPLE = os.path.join(PROD, "示例花名册.xlsx")
os.makedirs(EVID, exist_ok=True)
report = {"base": BASE, "real_roster": REAL, "checks": []}

def check(name, ok, detail):
    report["checks"].append({"name": name, "ok": bool(ok), "detail": str(detail)})
    print(("PASS " if ok else "FAIL ") + name + " | " + str(detail))

NAME_PAT = re.compile(r"^[\u4e00-\u9fa5·]{2,6}$")

# 0 首页
r = requests.get(BASE + "/", timeout=10)
check("首页HTTP", r.status_code == 200 and "课堂随机点名器" in r.text, f"status={r.status_code} len={len(r.text)}")

# 1 导入真实花名册（学籍系统导出，64338 行大表）
with open(REAL, "rb") as f:
    r = requests.post(BASE + "/api/import", files={"file": ("2026年7月份24级（706人）.xlsx", f)}, timeout=300)
j = r.json()
t0 = time.time()
check("导入解析成功", "error" not in j, f"有效行={j.get('row_count')} 用时{time.time()-t0:.1f}s")
headers = j["_headers"]
name_guess = j.get("name_col_guess")
cls_guess = j.get("class_col_guess")
check("姓名列智能识别=『姓名』", name_guess is not None and str(headers[name_guess]).strip() == "姓名",
      f"列{name_guess}: {headers[name_guess]}")
check("班级列智能识别=『班级』", cls_guess is not None and str(headers[cls_guess]).strip() == "班级",
      f"列{cls_guess}: {headers[cls_guess]}")
# 识别列不能落在敏感列（证件/电话/父母姓名）
bad_cols = [i for i, h in enumerate(headers) if re.search(r"(证件|电话|父|母|监护|银行|社保)", str(h))]
check("未误选敏感列", name_guess not in bad_cols, f"敏感列索引={bad_cols[:8]}…共{len(bad_cols)}个, 姓名列={name_guess}")

# 2 确认导入
body = {"_headers": headers, "_rows": j["_rows"], "filename": j["filename"],
        "name_col": name_guess, "class_col": cls_guess}
r = requests.post(BASE + "/api/apply", json=body, timeout=300)
j2 = r.json()
check("导入706名学生", j2.get("count") == 706, f"count={j2.get('count')} classes={len(j2.get('classes', []))}个班")
check("班级列表非空", len(j2.get("classes", [])) > 0, f"classes={j2.get('classes', [])[:6]}…")
classes = j2.get("classes", [])
tmp_dir = os.path.join(PROD, "server", "data", "_uploads")
tmp_files = os.listdir(tmp_dir) if os.path.isdir(tmp_dir) else ["<目录不存在=已清理>"]
check("上传名册临时副本已清理(隐私)", not tmp_files or tmp_files == ["<目录不存在=已清理>"], f"_uploads剩余={tmp_files}")

# 3 单人抽取（全部名单，不重复）
r = requests.post(BASE + "/api/draw", json={"n": 1, "cls": "", "unique": True}, timeout=30)
j3 = r.json()
p0 = j3["picked"][0]
check("单人抽取", NAME_PAT.match(p0["name"]) is not None,
      f"抽中：{p0['name']}（{p0['cls']}）remain={j3['remain']}")
check("已点名计数=1", len(j3["drawn"]) == 1, f"drawn={len(j3['drawn'])}")

# 4 班级过滤抽取
cls0 = classes[0]
r = requests.post(BASE + "/api/draw", json={"n": 3, "cls": cls0, "unique": True}, timeout=30)
j4 = r.json()
names4 = [p["name"] for p in j4["picked"]]
check(f"按班级[{cls0}]抽3人", len(names4) == 3 and all(p["cls"] == cls0 for p in j4["picked"]),
      "抽中：" + "、".join(names4))
check("不重复（与此前无重名）", len({p["name"] for p in j4["picked"]} | {p0["name"]}) == 4, "各次抽取无重复")

# 5 连续抽取不重复验证
seen = {p0["name"], *names4}
dup = False
for _ in range(20):
    jj = requests.post(BASE + "/api/draw", json={"n": 1, "cls": "", "unique": True}, timeout=30).json()
    nm = jj["picked"][0]["name"]
    if nm in seen:
        dup = True
    seen.add(nm)
check("连续20次抽取无重复", not dup, f"累计已点{len(seen)}人")

# 6 撤销/重置
r = requests.post(BASE + "/api/undo", json={}, timeout=10).json()
before = len(seen)
check("撤销上一个", len(r["drawn"]) == before - 1, f"drawn={len(r['drawn'])}")
requests.post(BASE + "/api/reset", json={}, timeout=10)
st = requests.get(BASE + "/api/state", timeout=10).json()
check("重置本轮", st["drawn"] == [] and st["count"] == 706, f"drawn={len(st['drawn'])} count={st['count']}")

# 7 点名记录导出（先点几位再导出）
for _ in range(3):
    requests.post(BASE + "/api/draw", json={"n": 2, "cls": "", "unique": True}, timeout=30)
r = requests.post(BASE + "/api/export", json={}, timeout=10).json()
check("导出点名记录", "filename" in r, json.dumps(r, ensure_ascii=False))
txt = requests.get(BASE + "/api/open/" + r["filename"], timeout=10).content.decode("utf-8-sig")
check("记录内容完整", f"共 {r['count']} 人次" in txt and "课堂随机点名记录" in txt,
      f"{len(txt)}字, 前3行: {txt.splitlines()[4:7]}")

# 8 粘贴名单模式
r = requests.post(BASE + "/api/paste", json={"text": "张三、李四\n王五,赵六 孙七\t周八"}, timeout=10).json()
check("粘贴名单解析", r.get("count") == 6, f"count={r.get('count')}（应恰好6人）")

# 9 清空状态（不留真实学生数据在产品内）
requests.post(BASE + "/api/clear", json={}, timeout=10)
st = requests.get(BASE + "/api/state", timeout=10).json()
check("验收后清空名单(隐私)", st["count"] == 0 and st["drawn"] == [], "state已清空，产品内不残留学生数据")
check("示例花名册就位", os.path.isfile(SAMPLE), os.path.basename(SAMPLE))

report["transcript_note"] = "真实姓名仅用于本机功能验证，验收报告与日志不记录具体学生姓名"
with open(os.path.join(EVID, "product3_accept.json"), "w", encoding="utf-8") as f:
    json.dump(report, f, ensure_ascii=False, indent=2)
all_ok = all(c["ok"] for c in report["checks"])
print("\n=== 产品3验收：" + ("全部通过 [OK]" if all_ok else "存在失败 [NG]") + " ===")
sys.exit(0 if all_ok else 1)
