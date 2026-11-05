# -*- coding: utf-8 -*-
"""产品1/产品3 边界渗透测试"""
import io, os, sys, json, time, requests

which = sys.argv[1]  # 1 or 3
results = []
def check(name, ok, detail=""):
    results.append({"name": name, "ok": bool(ok), "detail": str(detail)[:150]})
    print(("PASS " if ok else "**BUG** ") + name + " | " + str(detail)[:120])

if which == "1":
    BASE = "http://127.0.0.1:8765"
    print("========== 产品1 边界测试 ==========")
    r = requests.post(BASE + "/api/upload", files={"file": ("假音频.txt", io.BytesIO(("这不是音频" * 100).encode("utf-8")))}, data={"lang": "zh"}, timeout=30)
    jid = r.json().get("job_id")
    st = None
    for _ in range(30):
        st = requests.get(BASE + f"/api/job/{jid}", timeout=10).json()
        if st["status"] in ("done", "error"): break
        time.sleep(1)
    check("非音频文件 → 干净报错", st and st["status"] == "error" and st.get("error"), f"error={st.get('error') if st else None}")
    r = requests.post(BASE + "/api/upload", files={"file": ("空文件.m4a", io.BytesIO(b""))}, data={"lang": "zh"}, timeout=30)
    jid = r.json().get("job_id")
    st = None
    for _ in range(30):
        st = requests.get(BASE + f"/api/job/{jid}", timeout=10).json()
        if st["status"] in ("done", "error"): break
        time.sleep(1)
    check("0字节音频 → 干净报错", st and st["status"] == "error", f"error={st.get('error') if st else None}")
    check("任务号不存在 → 404", requests.get(BASE + "/api/job/does-not-exist", timeout=10).status_code == 404, "")
    check("导出未完成任务 → 400", requests.post(BASE + "/api/export", json={"job_id": "nope", "format": "txt"}, timeout=10).status_code == 400, "")
    r = requests.get(BASE + "/api/open/..%2F..%2Fserver.py", timeout=10)
    leak = "Flask" in r.text
    check("导出下载穿越被拦截", r.status_code in (400, 404) and not leak, f"HTTP {r.status_code}")
    r = requests.get(BASE + "/api/open/..", timeout=10)
    check("导出下载 .. 被拦截", r.status_code in (400, 404), f"HTTP {r.status_code}")
    # 并发两个任务
    ids = []
    for i in range(2):
        r = requests.post(BASE + "/api/upload", files={"file": (f"并发测试{i}.m4a", open(r"C:\Users\Administrator\Documents\录音\录音.m4a", "rb").read())}, data={"lang": "zh"}, timeout=60)
        ids.append(r.json()["job_id"])
    done = 0
    for _ in range(120):
        done = sum(1 for j in ids if requests.get(BASE + f"/api/job/{j}", timeout=10).json()["status"] == "done")
        if done == 2: break
        time.sleep(1)
    check("并发2个转写任务全部完成", done == 2, f"done={done}")
    h = requests.get(BASE + "/api/history", timeout=10).json()
    check("历史不含失败任务垃圾", all("假音频" not in x["name"] and "空文件" not in x["name"] for x in h), f"{len(h)}条")
elif which == "3":
    BASE = "http://127.0.0.1:8931"
    print("========== 产品3 边界测试 ==========")
    # .xls 老格式（真实文件，只看计数不外泄姓名）
    real_xls = r"Z:\github项目\xueji-xitong\24级学生（676）.xls"
    with open(real_xls, "rb") as f:
        r = requests.post(BASE + "/api/import", files={"file": ("old.xls", f)}, timeout=300)
    j = r.json()
    ok_import = "error" not in j and j.get("row_count", 0) >= 600
    check("老格式 .xls 导入解析", ok_import, f"row_count={j.get('row_count')} name_col={j.get('name_col_guess')}")
    if ok_import:
        body = {"_headers": j["_headers"], "_rows": j["_rows"], "filename": j["filename"],
                "name_col": j["name_col_guess"], "class_col": j["class_col_guess"]}
        j2 = requests.post(BASE + "/api/apply", json=body, timeout=300).json()
        check(".xls 导入学生数≥600", j2.get("count", 0) >= 600, f"count={j2.get('count')}")
    requests.post(BASE + "/api/clear", json={}, timeout=10)
    # 坏文件
    r = requests.post(BASE + "/api/import", files={"file": ("garbage.xlsx", io.BytesIO(os.urandom(4096)))}, timeout=60)
    check("乱码 xlsx → 干净报错", r.status_code == 400 and "解析失败" in r.json().get("error", ""), f"{r.json().get('error')}")
    # 粘贴边界
    check("粘贴空文本 → 400", requests.post(BASE + "/api/paste", json={"text": "、、、"}, timeout=10).status_code == 400, "")
    check("粘贴英文名 → 0 人", requests.post(BASE + "/api/paste", json={"text": "tom jack mary"}, timeout=10).json().get("count") in (None, 0) or requests.post(BASE + "/api/paste", json={"text": "tom jack mary"}, timeout=10).status_code == 400, "英文名被中文姓名规则拒绝")
    j = requests.post(BASE + "/api/paste", json={"text": "、".join(f"测试名{i}" for i in range(1, 61))}, timeout=10).json()
    check("粘贴60人全收", j.get("count") == 60, f"count={j.get('count')}")
    # 抽取边界
    j = requests.post(BASE + "/api/draw", json={"n": 6, "cls": "不存在的班", "unique": True}, timeout=10)
    check("不存在班级 → 400/空池报错", j.status_code == 400 if hasattr(j, 'status_code') else True, j.json().get("error", ""))
    # 用小名单抽小组：3人名单抽6人 → 服务端 min() 钳制
    requests.post(BASE + "/api/paste", json={"text": "张三 李四 王五"}, timeout=10)
    j = requests.post(BASE + "/api/draw", json={"n": 6, "unique": True}, timeout=10).json()
    check("小组人数>池子 → 钳制为3", len(j.get("picked", [])) == 3, f"picked={len(j.get('picked', []))}")
    # 抽干
    err = None
    for _ in range(5):
        j = requests.post(BASE + "/api/draw", json={"n": 1, "unique": True}, timeout=10)
        err = j.json().get("error")
        if err: break
    check("抽干后明确报错可恢复", err is not None and "重置" in err, f"error={err}")
    # 状态损坏恢复
    st_path = r"C:\Users\Administrator\Desktop\03-课堂随机点名器\server\data\state.json"
    open(st_path, "w", encoding="utf-8").write("{corrupted!!!")
    j = requests.get(BASE + "/api/state", timeout=10).json()
    check("state.json 损坏 → 服务不崩、名单清零", j.get("count") == 0, f"state={j}")
    requests.post(BASE + "/api/clear", json={}, timeout=10)

with open(f"evidence/pento_{which}.json", "w", encoding="utf-8") as f:
    json.dump(results, f, ensure_ascii=False, indent=2)
bugs = [x for x in results if not x["ok"]]
print(f"\n=== 产品{which} 边界测试：{len(results)-len(bugs)}/{len(results)} 通过，{len(bugs)} 个问题 ===")
