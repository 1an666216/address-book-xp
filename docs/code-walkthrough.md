# 代码讲解卡（Code Walkthrough）

> 用途：答辩/评审被问到代码时能答上来。按“请求怎么走 → 每个函数干什么 → 为什么这么设计 → 高频问题”组织。
> 读一遍约 10 分钟，配合 `app.py` 边看边对照。

## 1. 一次请求的完整路径

```
浏览器点击「新建联系人」
  → app.js: saveContact() 收集表单 → fetch POST /api/contacts（JSON）
  → app.py: Handler.do_POST → _create_contact()
      → normalize_contact() 校验+整理 → 重名检查 → save_contacts() 原子写盘
  → 返回 201 {ok:true, contact:{...}}
  → app.js: refresh() 重新拉列表 → renderList() + renderDetail() 重绘
```

## 2. `app.py` 的五个核心函数

| 函数 | 做什么 | 为什么这么设计 |
| --- | --- | --- |
| `normalize_contact(raw, existing)` | 统一生成存储结构：补 id、时间戳、布尔收藏、清洗 methods；更新时保留 `createdAt` | 新建和编辑共用一个函数，避免两套校验逻辑漂移 |
| `_clean_methods(raw)` | 丢掉空值；type 不在白名单的归为 `other` | 前端可以随便传，后端兜底，防脏数据 |
| `contacts_to_workbook(contacts)` | openpyxl 生成 xlsx：表头蓝底白字、列宽、冻结首行、自动筛选；同类多值用 `；` 拼一格 | 作业要求“一行一人、一类一列”，多值合并进一格才不会破坏行结构 |
| `workbook_to_contacts(blob)` | 按**表头名称**建列映射（支持中英文别名），逐行解析，`；` 拆多值 | 列顺序可乱、多余列忽略，比固定列号健壮 |
| `merge_imported(contacts, imported)` | 按姓名去重：新名字新增；老名字只追加缺失的 (type,value)，并合并收藏/备注 | “导入不能把已有数据冲掉”，这是最容易被扣分的行为 |

## 3. 存储与并发

```python
tmp = DATA_FILE + ".tmp"
json.dump(..., tmp)
os.replace(tmp, DATA_FILE)   # 原子替换，写一半断电也不会坏文件
```

- 所有读写都在 `_lock = threading.Lock()` 里，因为 `ThreadingHTTPServer` 是多线程的。
- 为什么不用数据库：作业一个人 2 天，JSON 足够；评审 `python app.py` 就能跑，零依赖。

## 4. `static/app.js` 的组织

- `state`：唯一可变状态（contacts / keyword / onlyFav / selectedId），改完调用 `render*()` 重绘。
- `api()`：统一 fetch 包装，把 `{ok:false,error}` 转成异常，页面只处理 `toast()`。
- `renderList/renderDetail`：纯字符串模板 + `escapeHtml()`，所有用户输入都转义（防 XSS）。
- 编辑弹窗用 `<dialog>` + `<template>` 动态加行，不引入框架。

## 5. 高频问题与参考回答

**Q：为什么 `methods` 用数组，不用 `phone1/phone2`？**
A：需求是“一人多种联系方式”。数组天然支持任意条数、任意类型；加一种类型只在 `METHOD_TYPES` 加一行，导出列、前端下拉、校验都会自动跟随。

**Q：导入按姓名去重，重名的两个人怎么办？**
A：这是刻意的简化，模板的“填写说明”里写明了。真实系统会用手机号或外部 ID 做唯一键；姓名只作为弱匹配。

**Q：导出为什么把多个手机号拼在一格？**
A：作业要求“每个联系人一行”。若一人两个手机号就占两行，行数和人数就对不上；用 `；` 分隔后导入端再拆开，往返无损。

**Q：上传非 xlsx 为什么会 400 而不是 500？**
A：`workbook_to_contacts` 里 `load_workbook` 包了 try/except 转 `ValueError`，HTTP 层把 `ValueError` 统一映射成 400。这是 e2e 测试里专门断言的一条。

**Q：表头 `.value` 报错是怎么回事？**
A：`ws.iter_rows(values_only=True)` 返回纯值，`_header_map` 却按单元格对象取 `.value`。修复是兼容两种输入：`cell.value if hasattr(cell,"value") else cell`。

**Q：收藏为什么能排在最前？**
A：`contacts.sort(key=lambda c: (not c.get("bookmarked"), c.get("name","")))`——False 排在 True 前面，所以收藏的在前，同组内按姓名。

**Q：前端搜索怎么做到“搜微信也能搜到”？**
A：`_list_contacts` 的匹配串里同时放了 `method.value` 和该类型的**中文名**（`METHOD_TYPES[type][0]`），所以输入“微信”能命中所有带微信的人。

**Q：数据存在哪？重启会丢吗？**
A：`data/contacts.json`，启动时读、每次写操作原子落盘，重启不丢。该文件在 `.gitignore` 里，避免把演示数据提交进仓库。

**Q：怎么证明功能真的能用？**
A：`tests/e2e_test.py` 27 条断言，真实走 HTTP，覆盖增删改查、收藏开关与过滤、导入导出往返、坏文件 400、清理还原。跑法：`python tests/e2e_test.py --base http://127.0.0.1:8000`。

## 6. 文件职责对照

本项目独立完成，各文件职责如下，便于答辩时按模块解释：

| 文件 | 关键点 |
| --- | --- |
| `static/index.html`、`style.css`、`app.js` | 双栏布局、状态管理、渲染、收藏交互、导入导出入口 |
| `app.py`、`seed.py` | 数据模型、CRUD、Excel 读写、去重合并、端到端测试所需的领域逻辑 |
| `tests/`、`tools/` | 端到端测试、导入模板生成 |
