# 课堂随机点名器 · 开发手册

> 面向二次开发者与维护者。所有代码片段均摘自当前仓库真实文件，行号以最新 main 为准。
> 配套阅读：`README.md`（用户视角）。

## 1. 项目结构

```
ketang-dianmingqi/
├── .github/
│   └── workflows/pages.yml        # GitHub Actions：push main 后把 web/ 部署到 GitHub Pages
├── .gitignore                     # 忽略 __pycache__、desktop/server/data/、点名记录/ 等
├── desktop/
│   ├── 启动课堂点名器.bat          # 双击入口：cd 到 server/ 后 python server.py（GBK 编码，见 §9）
│   ├── 示例花名册.xlsx             # 虚构演示名册，用于演示与验收对照
│   └── server/
│       ├── server.py              # 桌面版全部后端：Flask API、名册解析、列识别、抽人、状态落盘
│       ├── static/index.html      # 桌面版前端：调用 /api/* 的单页界面（与网页版同构）
│       ├── data/                  # （运行时生成，不入库）state.json、_uploads/ 临时目录
│       └── 点名记录/               # （运行时生成，不入库）/api/export 写出的 TXT
├── docs/
│   ├── DEVELOPMENT.md             # 本手册
│   └── screenshots/               # README 用的界面截图
├── web/
│   ├── index.html                 # 网页版：单文件纯前端（解析表格/抽人/状态全在浏览器本地）
│   └── vendor/xlsx.full.min.js    # vendored SheetJS（xlsx.js 全量版，~0.86MB），解析 xlsx/xls/csv
├── LICENSE                        # MIT
└── README.md                      # 用户文档与截图
```

一句话职责：

| 文件 | 职责 |
|---|---|
| `desktop/server/server.py` | 桌面版 HTTP 服务与全部业务逻辑（11 个端点 + 名册解析 + 安全防护） |
| `desktop/server/static/index.html` | 桌面版界面，纯 fetch 调 API，不含业务算法 |
| `web/index.html` | 网页版界面 + 全部业务逻辑（无后端） |
| `web/vendor/xlsx.full.min.js` | SheetJS 库，供网页版在浏览器本地解析表格 |
| `desktop/启动课堂点名器.bat` | Windows 双击启动器，自动挑端口并开浏览器 |
| `.github/workflows/pages.yml` | CI：把 `web/` 作为站点根发布到 GitHub Pages |
| `docs/screenshots/*` | README 展示图 |

## 2. 技术栈与运行环境

**先划清界限：本项目与 faster-whisper / 语音识别完全无关**（那是同系列其他产品的依赖）。本项目刻意零重依赖：

| 形态 | 技术栈 | 说明 |
|---|---|---|
| 桌面版 | Python 3.10+ / Flask / openpyxl / xlrd | 后端 `server.py` 一个文件；前端原生 HTML/CSS/JS，无构建步骤 |
| 网页版 | 纯前端（原生 JS）+ vendored SheetJS | `xlsx.full.min.js` 直接随仓库分发，**不引入任何 CDN / npm**，保证离线内网可用 |

桌面版 Python 依赖（仓库无 requirements.txt，就这三样）：

```bash
pip install flask openpyxl xlrd
```

- `openpyxl`：读 `.xlsx`（`read_only=True` 流式读大表）
- `xlrd`：读老格式 `.xls`（学籍系统常导出这种）
- `.csv`：标准库 `csv` 模块，无第三方依赖

**浏览器要求**：两端前端都是 ES2017+ 原生 JS（`async/await`、模板字符串、展开运算符、`file.arrayBuffer()`），需要常青浏览器：

- Chrome / Edge ≥ 80，Firefox ≥ 75，Safari ≥ 13.1
- 不支持 IE（机房若只有 IE，用桌面版）
- 网页版依赖 `localStorage`（存名单）与 `Blob` 下载（导记录），无需任何网络请求

## 3. 本地开发

### 3.1 桌面版

```bash
cd desktop/server
python server.py
```

启动逻辑（`server.py` 末尾）：从 **8931 起向上试 30 个端口**（8931~8960），谁空闲用谁，并自动开浏览器：

```python
def pick_port(start):
    for port in range(start, start + 30):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            try:
                s.bind(("127.0.0.1", port))
                return port
            except OSError:
                continue
    raise RuntimeError("端口都被占用")
```

- 只绑 `127.0.0.1`，局域网其他机器访问不到（隐私边界）
- 设 `ZC_NO_BROWSER=1` 环境变量可跳过自动开浏览器（自动化测试用）：

```bash
ZC_NO_BROWSER=1 python server.py   # 然后手动访问 http://127.0.0.1:8931
```

- 运行时数据落在 `desktop/server/data/state.json`（原子写：先写 `.tmp` 再 `os.replace`），上传临时件在 `data/_uploads/`，导出的记录在 `desktop/server/点名记录/`。三者均被 `.gitignore` 忽略。

### 3.2 网页版

```bash
cd web
python -m http.server 8080
# 打开 http://127.0.0.1:8080
```

不启服务直接双击 `index.html`（file://）也能用（localStorage 与 SheetJS 在 file:// 下均可用），但开发时建议走 http，行为与 Pages 线上一致。

### 3.3 改完如何验证

1. **手工冒烟**：导入 `desktop/示例花名册.xlsx` → 单人点名 → 开不重复连点几次 → 2 人小组 → 撤销 / 重置 → 导出 TXT 能下载。
2. **改了列识别 / 去重逻辑**：用学籍系统导出的真实大表验证（见 §6，注意隐私红线）。
3. **改了 API**：跑 `scripts/accept_3.py`（§6）。
4. **改了 web/index.html**：注意**桌面端 `desktop/server/static/index.html` 是另一份代码**，两端都要同步改（见 §4 对照表），否则线上版和桌面版行为分叉。

## 4. 架构与关键实现

### 4.1 双形态同构

同一套产品逻辑在两端各实现一份，**没有共享代码**（网页版必须零依赖单文件，桌面版必须离线），因此改算法时必须两端同步。对应关系：

| 能力 | 桌面版 | 网页版 |
|---|---|---|
| 解析表格 | `server.py: read_table()`（openpyxl/xlrd/csv） | `web/index.html: readTable()`（SheetJS `sheet_to_json`） |
| 姓名列识别 | `guess_name_col(headers, rows)` | `guessNameCol(headers, rows)` |
| 班级列识别 | `guess_class_col(headers)` | `guessClassCol(headers)` |
| 学生解析/去重 | `parse_students()`（`/api/apply` 调用） | `applyBtn.onclick` 内联逻辑 |
| 粘贴名单 | `/api/paste`（服务端正则过滤） | `pasteApply.onclick`（同一套正则） |
| 随机抽人 | `/api/draw` → `random.sample(pool, n)` | `pickN()`（Fisher-Yates 洗牌后取前 n） |
| 不重复模式 | 服务端按 `drawn` 里的 `id` 过滤池子 | 前端按 `drawn` 里的 `id` 过滤池子 |
| 状态存储 | `server/data/state.json`（原子写） | `localStorage["dianmingqi_state_v1"]` |
| 导出记录 | `/api/export` 落盘 + `/api/open/<name>` 下载 | `Blob` + `a.download` 前端直接下载 |
| 撤销/重置 | `/api/undo`、`/api/reset` | `state.drawn.pop()`、`state.drawn=[]` |

两端共享同一组识别正则（一字符不差，改一处必须改另一处）：

| 正则 | 桌面 `server.py` | 网页 `index.html` |
|---|---|---|
| 中文姓名格式 | `^[\u4e00-\u9fa5·‧•]{2,6}$` | 同 |
| 姓名表头精确匹配 | `^(姓名\|学生姓名\|学生\|name)$` | 同（多 `/i`） |
| 脏表头排除 | `(父\|母\|监护\|联系\|备注\|联系)` | `(父\|母\|监护\|联系\|备注)` |
| 班级表头 | `(班级\|班别\|行政班)` | 同 |
| 敏感列排除 | `(证件\|电话\|号码\|身份\|地址\|状态\|银行\|社保)` | 同 |

### 4.2 姓名/班级列识别算法（三级递进）

`server.py: guess_name_col()`，网页版 `guessNameCol()` 逻辑相同：

```python
def guess_name_col(headers, rows=None):
    # 1) 表头精确匹配『姓名/学生姓名/学生/name』，且不命中父母/监护人/联系人等脏表头
    for i, h in enumerate(headers):
        h = str(h).strip()
        if h and _NAME_HEADER.match(h) and not _BAD_HEADER.search(h):
            return i
    # 2) 表头含『姓名』（如"学生姓名"、"幼儿姓名"），排除敏感列
    for i, h in enumerate(headers):
        h = str(h).strip()
        if "姓名" in h and not _BAD_HEADER.search(h) and not _EXCLUDE_HEADER.search(h):
            return i
    # 3) 兜底：逐列统计中文姓名格式占比，取最高的列
    if rows:
        best, best_n = None, 0
        for i in range(len(headers)):
            n = sum(1 for r in rows[:5000] if i < len(r) and _NAME_PAT.match(str(r[i]).strip()))
            if n > best_n:
                best, best_n = i, n
        if best_n:
            return best
    return None
```

- 第 3 级兜底只看**前 5000 行**（学籍导出几万行时控制耗时）
- 兜底能救“表头是空白/乱码”的表：姓名列的值必然大量命中 `^[\u4e00-\u9fa5·‧•]{2,6}$`（2~6 个汉字，允许间隔号）
- `_EXCLUDE_HEADER`（证件/电话/号码/身份/地址/状态/银行/社保）在 2、3 两级都生效——学籍表里“身份证号”这类列即使长得像也绝不选
- 班级列 `guess_class_col()`：表头含 `班级|班别|行政班` 且不命中敏感列即中，没有兜底
- 识别只是**预选**：前端会下拉展示所有列，教师可手动换列，最终以 `/api/apply` 传回的 `name_col`/`class_col` 为准

### 4.3 整行去重与学生 id

**不能按姓名去重**——真实名册存在同名同班的不同学生。所以去重的粒度是“整行完全相同”（即重复录入的同一行）：

```python
def parse_students(headers, rows, name_col, class_col):
    """同名学生（真实名册存在）各自独立成条，靠 id 区分；
    只有整行完全相同的重复录入才合并。"""
    students, seen_rows = [], set()
    for r in rows:
        name = str(r[name_col]).strip() if name_col < len(r) else ""
        name = re.sub(r"\s+", "", name)
        if not name:
            continue
        if name in ("姓名", "学生姓名"):   # 表头行混进数据时跳过
            continue
        key = tuple(str(c).strip() for c in r)
        if key in seen_rows:
            continue
        seen_rows.add(key)
        ...
        students.append({"id": len(students), "name": name, "cls": cls})
```

- 网页版同一逻辑：`const key=r.map(c=>String(c).trim()).join("|")`
- **id 语义**：`id = len(students)` 顺序分配、0 起始、**导入完成后永不改变**。三处分配点必须一致：`parse_students`、`api_paste`、网页版 `applyBtn`。不重复模式、“撤销”全都以 id 为锚，改导入逻辑时不能破坏这一点
- 名单从粘贴进入时同样有 `seen` 集合按姓名去重（粘贴场景没有“行”的概念，同名即视为同人是可接受的近似；Excel 导入路径则严格整行去重）

### 4.4 掷点协议

服务端（`/api/draw`）：

```python
n = max(1, min(int(data.get("n", 1)), 12))      # 人数钳制 1~12
cls = data.get("cls") or ""
unique = bool(data.get("unique", True))
st = load_state()
pool = [s for s in st["students"] if (not cls or s["cls"] == cls)]   # 班级过滤
if unique:
    drawn_ids = {d["id"] for d in st["drawn"]}
    pool = [s for s in pool if s["id"] not in drawn_ids]             # 不重复：按 id 剔除
if not pool:
    return jsonify({"error": "可选的人抽完了：请重置本轮或关闭『不重复』"}), 400
n = min(n, len(pool))                            # 池子不够时钳制
picked = random.sample(pool, n)                  # 不放回均匀抽样
now = time.strftime("%H:%M:%S")
for s in picked:
    st["drawn"].append({"id": s["id"], "name": s["name"], "cls": s["cls"], "time": now})
save_state(st)
return jsonify({"picked": picked, "drawn": st["drawn"], "remain": len(st["students"]) - len(st["drawn"])})
```

网页版在前端等价实现：`pickN()` 用 Fisher-Yates 洗牌后取前 n 个（`random.sample` 的 JS 对应物）：

```js
function pickN(pool,n){
  const arr=pool.slice();
  for(let i=arr.length-1;i>0;i--){const j=Math.floor(Math.random()*(i+1));[arr[i],arr[j]]=[arr[j],arr[i]]}
  return arr.slice(0,n);
}
```

- 前端 UI 的小组人数下拉是 1~6，服务端上限是 12（服务端钳制才是权威边界）
- 两端都有滚动动画：16 次假名滚动，间隔从 60ms 每轮 +8ms，最后定格真名。桌面版定格在 `static/index.html` 的 `singleTick()` else 分支，网页版在 `draw()` 内 `tick()` else 分支（§8b 挂钩点）

## 5. 数据契约

### 5.1 桌面版 HTTP API（全部端点）

| 方法 | 路径 | 入参 | 出参（JSON） |
|---|---|---|---|
| GET | `/` | - | 桌面版页面（render_template） |
| POST | `/api/import` | multipart 字段 `file`（.xlsx/.xls/.csv） | `{file, filename, columns:[{idx,header}], name_col_guess, class_col_guess, row_count, _headers:[...], _rows:[[...]]}`；解析失败 400 `{error}` |
| POST | `/api/apply` | `{_headers, _rows, filename, name_col, class_col}` | `{ok:true, count, classes:[...]}`；无姓名 400 `{error}` |
| POST | `/api/paste` | `{text}` | `{ok:true, count, classes:[]}`；解析不出姓名 400 |
| GET | `/api/state` | - | `{count, drawn:[{id,name,cls,time}], source, classes}` |
| POST | `/api/draw` | `{n, cls, unique}` | `{picked:[{id,name,cls}], drawn:[...], remain}`；池子空 400 |
| POST | `/api/reset` | - | `{ok:true}`（drawn 清零，名单保留） |
| POST | `/api/undo` | - | `{ok:true, drawn}`（弹出一个） |
| POST | `/api/clear` | - | `{ok:true}`（名单+记录全清，并清 `_uploads`） |
| POST | `/api/export` | - | `{filename, count}`；本轮无记录 400；TXT 写入 `点名记录/点名记录_YYYYMMDD_HHMMSS.txt`（`utf-8-sig`，Excel 打开不乱码） |
| GET | `/api/open/<name>` | 路径参数：文件名 | 文件流下载；经 `safe_name` + `contained` 双重校验，非法返回 404 |

要点：

- `/api/import` 返回的 `_rows` **截断到 50000 行**，且已剔除整行为空的记录（学籍导出常见几万空行）；`row_count` 是剔除后的有效行数
- `/api/import` 后文件存到 `data/_uploads/<uuid>_<safe_name>`，`/api/apply` 成功或 `/api/clear` 时 `cleanup_tmp()` 全部删除——名册是敏感数据，不留副本
- `drawn` 条目形状：`{"id": 3, "name": "王五", "cls": "24级1班", "time": "10:23:45"}`，`remain = 名单总数 - len(drawn)`

### 5.2 网页版 localStorage：`dianmingqi_state_v1`

`web/index.html` 中的存取：

```js
const LS_KEY="dianmingqi_state_v1";
function save(){localStorage.setItem(LS_KEY,JSON.stringify(state))}
function load(){
  try{ state=JSON.parse(localStorage.getItem(LS_KEY))||state }catch(e){}
  refresh();
}
```

完整 JSON 结构示例（含同名学生、已点一人）：

```json
{
  "students": [
    { "id": 0, "name": "张三", "cls": "24级软件1班" },
    { "id": 1, "name": "李四", "cls": "24级软件1班" },
    { "id": 2, "name": "张三", "cls": "24级软件2班" }
  ],
  "drawn": [
    { "id": 1, "name": "李四", "cls": "24级软件1班", "time": "08:05:12" }
  ],
  "source": "2026年7月份24级（706人）.xlsx"
}
```

- 顶层三键 `students / drawn / source` 与桌面版 `state.json` 完全一致，`drawn` 条目形状与 API 相同
- `id: 0` 与 `id: 2` 两个“张三”是不同学生——整行去重的意义所在
- 损坏的 JSON 会被 `try/catch` 静默吞掉回到空态（宁可丢状态不可崩页面）

## 6. 测试与验收

验收脚本不在本仓库，在 workspace 的 `scripts/` 目录（产品仓库保持干净）：

| 脚本 | 用途 |
|---|---|
| `scripts/accept_3.py` | 产品 3（本项目）功能验收：导入真实名册 → 验证列识别/去重/不重复/班级过滤/导出 |
| `scripts/pento_1_3.py` | 边界渗透测试（传参数 `3` 跑本项目）：路径穿越、坏文件、粘贴边界、抽人钳制 |

**用法**（先启动桌面版，再跑脚本）：

```bash
# 终端 1
cd desktop/server && ZC_NO_BROWSER=1 python server.py

# 终端 2：验收。三个参数：服务地址、证据目录、产品目录
python scripts/accept_3.py http://127.0.0.1:8931 evidence "C:\path\to\ketang-dianmingqi"

# 渗透：参数 3 = 本产品
python scripts/pento_1_3.py 3
```

`accept_3.py` 会检查：首页 200、真实学籍大表（6 万+ 行）导入成功、姓名列猜中『姓名』、班级列猜中『班级』、**识别列不落在敏感列**、导入 `_uploads` 临时副本已清理（隐私）、706 人计数正确、单人/小组抽取、不重复模式不重样、班级过滤、TXT 导出。

**真实名册验证与隐私红线**：

- 脚本引用的真实花名册路径写在脚本参数/常量里，**只在本机跑**，证据目录（截图、JSON 报告）不得提交进仓库
- **严禁把任何真实学生数据（xlsx/csv/名单文本/截图）提交到本仓库**。`.gitignore` 已忽略 `desktop/server/data/` 与 `点名记录/`，但根目录下临时放的名册文件不会被忽略——提交前 `git status` 过一眼
- 仓库里唯一的表格是 `desktop/示例花名册.xlsx`（虚构名单），构造新测试数据时请用它或再造虚构数据

## 7. 构建与部署

### 7.1 GitHub Pages（网页版）

`.github/workflows/pages.yml` 全文核心就是"把 `web/` 目录当站点根上传"：

```yaml
- uses: actions/upload-pages-artifact@v3
  with:
    path: web
- id: deployment
  uses: actions/deploy-pages@v4
```

**为什么不能走 legacy 分支式部署**：旧式 Pages 只有两种选择——仓库根 `/` 或 `docs/` 目录，没有"自定义子目录"一说。本仓库站点在 `web/` 下，仓库根又有 README/桌面版等不该公开成站点的内容；若走分支式，要么把 `web/` 内容复制到根（仓库结构与源码错位），要么塞进 `docs/`（与截图目录语义混淆）。Actions 工作流（`upload-pages-artifact` 的 `path: web`）直接指定子目录为站点根，push main 即部署，无需维护发布分支。

线上地址：`https://openclawzhangchong.github.io/ketang-dianmingqi/`（以 README 当前链接为准）。注意 Pages 部署的是 `web/`，`desktop/` 不会被发布。

### 7.2 桌面版交付

**复制目录即用**：把整个 `desktop/` 文件夹拷到教师机（需装有 Python 3.10+ 与三个 pip 依赖），双击 `启动课堂点名器.bat`。无安装器、无注册表、无服务，删文件夹即卸载。

`启动课堂点名器.bat`（全文）：

```bat
@echo off
title 课堂随机点名器
cd /d "%~dp0server"
python server.py
if errorlevel 1 (
  echo.
  echo [启动失败] 请截图本窗口内容联系维护人。
  pause
)
```

**为什么 .bat 必须 GBK 编码**：中文 Windows 的 cmd.exe 默认代码页是 936（GBK）。`title` 和 `echo` 的中文若以 UTF-8 存盘，cmd 按 GBK 解读就是乱码（窗口标题变"锟斤拷"类乱串，报错提示不可读）。本仓库的 bat 以 **GBK + CRLF** 提交，编辑它时务必保持该编码（VS Code 右下角选 GBK 保存；不要用默认 UTF-8 覆盖）。

## 8. 二次开发指南

### a. 新增名单格式（如 JSON 名单 / 学籍系统新模板）

改两处，逻辑保持同构：

- **桌面** `desktop/server/server.py` → `read_table()`。它现在按后缀分派（`.csv` → 编码尝试 utf-8-sig/gbk/utf-8；`.xls` → xlrd；其余 → openpyxl），加一个分支即可，例如 `.json`：

```python
if suffix == ".json":
    data = json.loads(path.read_text(encoding="utf-8"))
    # 归一到 (headers, rows)：rows 必须是字符串二维数组，第一行为表头
    headers = list(data[0].keys())
    rows = [[str(r.get(h, "")) for h in headers] for r in data]
    return headers, rows
```

- **网页** `web/index.html` → `readTable(wb)`。它接收 SheetJS workbook，返回 `{headers, rows}`（rows 为字符串二维数组、已剔除空行）。在 `upload(f)` 里按扩展名分流：`f.name.endsWith(".json")` 时直接 `JSON.parse` 归一成同样形状，不走 `XLSX.read`
- **注意事项**：
  - 归一目标形状固定：`headers` 一维字符串数组 + `rows` 二维字符串数组。后续 `guessNameCol` / 去重 / id 分配全部基于这个形状，别返回对象数组直接穿透
  - 若新模板的"姓名"表头不在 `_NAME_HEADER` 精确列表里（如"幼儿姓名"），现有第 2 级"含姓名"规则已能兜住；若表头完全不同，在**两端**同步扩 `_NAME_HEADER`/`BAD_HEADER` 正则（对照表见 §4.1）
  - `guessNameCol` 有一端改了，另一端必须同步，否则同一份名册在网页版和桌面版选出不同列

### b. 加点名音效 / 语音播报

挂钩点在**滚动动画定格处**，两端各一处：

- **网页** `web/index.html` → `draw()` 内 `tick()` 的 else 分支（定格、写 `state.drawn` 的地方）：

```js
else{
  $("name").textContent=names[0];$("name").classList.add("done");
  /* ← 在这里挂钩：播音效 + speechSynthesis 播 names.join("、") */
  $("note").textContent=n>1?("小组已抽出（"+n+" 人）"):"就是你了！";
  ...
  picked.forEach(p=>state.drawn.push({id:p.id,name:p.name,cls:p.cls,time:now}));
  save();rolling=false;refresh();
}
```

- **桌面** `desktop/server/static/index.html` → `draw()` 内 `singleTick()` 的 else 分支（同样位置）
- **实现建议**：浏览器原生 `speechSynthesis.speak(new SpeechSynthesisUtterance(名字))` 即可，`lang="zh-CN"`、`rate: 0.9`；音效文件放 `web/assets/`（桌面版放 `static/assets/`），用 `new Audio("...")` 播。原生 API，零依赖，符合本项目"不引第三方"的纪律
- **注意事项**：
  - **两端都要加**（§4.1 同构原则），只改 web 会导致桌面版没声音
  - 播报放在 `rolling=false` 之前的定格处即可；不要在每次 `tick()` 里播（16 次滚动会念一串假名）
  - 自动播放策略：Chrome 要求用户先与页面交互；点名本身是点击/空格触发，满足条件，但**不要**试图在页面加载时预播
  - 桌面版若把音频做成可配置，配置也存 `state.json` / localStorage（`dianmingqi_state_v1` 加新键），注意旧数据没有该键时的默认值兜底

### c. 接入学籍系统在线数据

建议在**桌面版服务端**加只读代理端点，而不是让前端直接调学籍系统：

```python
@app.post("/api/fetch_roster")
def api_fetch_roster():
    data = request.get_json(force=True)
    url = data.get("url", "")
    if not url.startswith("https://xueji.example.edu/"):   # 1) 域名白名单
        return jsonify({"error": "仅支持学籍系统地址"}), 400
    import urllib.request
    req = urllib.request.Request(url, headers={"Authorization": os.environ.get("XUEJI_TOKEN", "")})
    with urllib.request.urlopen(req, timeout=30) as resp:  # 2) 凭据走环境变量，不进前端/不进仓库
        raw = resp.read()
    # 3) 复用现有管线：解析→guess_name_col→parse_students→save_state，不落盘原始文件
    ...
```

- **为什么放服务端**：跨域（CORS）、内网认证、证书都好处理，且凭据不暴露给浏览器
- **信任边界（认真读）**：桌面版只绑 `127.0.0.1`，这只挡**外部机器**，挡不住本机上的其他程序和恶意网页。本机任意网页都能向 `http://127.0.0.1:8931` 发跨域请求（Flask 默认无 CSRF 防护；`Content-Type: application/json` 的 POST 虽会触发预检，但简单请求如 form/multipart 不会）。因此新端点必须：只读、白名单域名、不提供任意 URL 转发（防 SSRF 打内网）、不落盘原始数据、超时短。**绝不要**做"通用下载代理"这种端点
- 网页版（GitHub Pages 纯前端）**不建议**接在线学籍：需要把凭据放进浏览器，且违背"名单不出本机"的隐私承诺；如确要做，走用户本机提供的 token + 用户显式操作触发

### d. 换抽点算法（加权 / 勿点某生）

- **服务端**：`desktop/server/server.py` 的 `/api/draw`，就一行 `picked = random.sample(pool, n)`（§4.4）。加权抽点可改为按权重不放回抽样（如循环 `random.choices` + 从池中移除，或按权重排序取前 n）；**勿点某生**只需在 `pool` 过滤处加条件
- **网页端**：`web/index.html` 的 `pickN(pool,n)`（Fisher-Yates）。两端必须实现同一分布，否则同一份名册在两端概率不同
- **id 语义**：`id` 在导入时顺序分配后不再变（§4.3）。加权表 / 排除名单应以 `name+cls` 或 `id` 为键在**抽人时查表**，不要在导入时改写 `students` 数组（那会移动所有人的 id，导致旧 `drawn` 记录对不上号）
- **不重复集合**：`unique` 模式用 `{d["id"] for d in st["drawn"]}` 过滤。若引入"权重随被点次数衰减"，衰减状态要么并入 `drawn` 条目（向后兼容新增键），要么存独立键，并保证 `state.json`（桌面）与 `dianmingqi_state_v1`（网页）同步演进
- 保留 `n = min(n, len(pool))` 与空池 400 的钳制行为——渗透脚本 `pento_1_3.py` 对此有断言

## 9. 常见坑（真实踩过）

1. **中文 .bat 必须 GBK**：UTF-8 存盘后 `title`/`echo` 中文全乱（§7.2）。同理 bat 要 CRLF。若终端出现 `warning: LF will be replaced by CRLF`，是 git autocrlf 的正常提示，**无害**，不必处理；但**不要**为了消警告把 bat 转成 LF + UTF-8。
2. **同名学生必须整行去重**：按姓名去重会把同名同班的两个真实学生合并成一人，点名永远点不到第二个。唯一正确粒度是"整行完全相同才合并"（§4.3）。反过来，粘贴名单入口按姓名去重是有意为之，别"顺手统一"成整行（粘贴没有行概念）。
3. **Windows 路径穿越**：`/api/open/<name>` 曾是攻击面（`..%2F..%2Fserver.py` 这类载荷，`pento_1_3.py` 有回归用例）。防线有二，缺一不可：
   - `safe_name()`：`os.path.basename` 剥目录 → 非法字符替换 → **剥首尾点号（阻断 `..`）** → Windows 保留设备名（con/prn/aux/nul/com1-9/lpt1-9）加 `_` 前缀；
   - `contained()`：`os.path.commonpath` 校验最终路径仍在 `REC_DIR` 内（防御纵深）。
   新增任何"按文件名取文件"的端点都必须复用这对组合。
4. **read_only 大表性能**：学籍导出 6 万+ 行，openpyxl 必须开 `read_only=True, data_only=True` 流式读（`server.py` 已如此），且超 200000 行主动 break；`/api/import` 回传 `_rows` 截 50000 行、剔除全空行——这些都是实测后加的，别在"重构"中删掉，也别把 `iter_rows(values_only=True)` 换成整表 `list(ws.values)`。
5. **两端分叉**：web 版与桌面版是两份独立实现（§4.1），改一处忘另一处是本项目最易发生的回归。每次改动后对照 §4.1 的表逐行核对。
6. **桌面版 static 前端插值未转义**：`desktop/server/static/index.html` 的 `drawnList`/班级下拉直接内插 `d.name`/班级名，未走 `esc()`（web 版有 `esc()`）。数据源是教师自己导入的名册，风险低，但若接入外部数据源（§8c），先补转义。
7. **state.json 损坏容错**：`load_state()` 解析失败静默回空态。手工改 state 调试时先备份；写状态永远走 `save_state()`（tmp + `os.replace` 原子替换），别直接 `open(STATE_F,"w")`。

## 10. 维护与发布流程清单

1. **改码**：按 §8 对照表两端同步改；遵循现有风格（server.py 单文件、web/index.html 单文件、无构建步骤）。
2. **本地验证**：`python server.py` 手工冒烟（§3.3）；动过解析/去重/抽人逻辑的，跑真实名册验收：
   ```bash
   python scripts/accept_3.py http://127.0.0.1:8931 evidence "<产品目录>"
   python scripts/pento_1_3.py 3
   ```
   全 PASS 再继续；证据目录不入库。
3. **同步桌面副本**：`web/index.html` 的改动若涉及同构逻辑，镜像到 `desktop/server/static/index.html`（UI 文案/样式两端可有意不同，算法不行）。
4. **提交前自检**：`git status` 确认没有真实学生数据、没有 `desktop/server/data/`、没有证据目录混入；bat 改动确认仍是 GBK。
5. **commit + push**：commit message 用 `feat:/fix:/docs:/ci:` 前缀（见 `git log` 惯例）。
6. **Actions 部署**：push main 后 Pages 工作流自动跑（仓库 Actions 页可看进度），只发布 `web/`。
7. **线上验证**：打开 Pages 地址，导入示例花名册点两轮；若班级下拉/导出行为与本地不一致，先查是否两端分叉（第 3 步漏做）。
8. **回滚**：Pages 用的是 main 代码，回滚即 `git revert` 后再 push；桌面版无版本概念，发布即"把 main 的 `desktop/` 打包发给教师"。
