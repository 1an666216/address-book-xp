# 极限编程作业：通讯录应用（收藏 / 多种联系方式 / Excel 导入导出）

> 草稿说明：把 `【替换】` 标记处换成你的姓名学号、截图链接、视频链接即可发布。
> 截图可直接上传 `docs/screenshots/home.png`、`editor.png`。

| 项目 | 内容 |
| --- | --- |
| 本次作业所属课程 | [2601_FZU-MU_SE](https://bbs.csdn.net/forums/2601_MU_SE_FZU) |
| 提交形式 | 独立完成（单人） |
| 作业要求 | [Extreme Programming(Deadline: November 17, 2026, 23:59)](https://bbs.csdn.net/topics/620593436) |
| 本次作业目标 | 实现通讯录的收藏、多种联系方式、Excel 导入导出，并实践 Git 与软件工程流程 |
| 其他参考文献 | openpyxl 官方文档、MDN fetch API、《构建之法》第 4 章 |

## 目录

- [一、作业要求拆解](#一作业要求拆解)
- [二、项目链接](#二项目链接)
- [三、Git 提交证据](#三git-提交证据)
- [四、实现思路](#四实现思路)
- [五、程序运行截图与演示视频](#五程序运行截图与演示视频)
- [六、开发职责说明](#六开发职责说明)
- [七、贡献比例](#七贡献比例)
- [八、开发中的问题与解决](#八开发中的问题与解决)
- [九、PSP 表](#九psp-表)
- [十、总结](#十总结)

---

## 一、作业要求拆解

作业要求在 **2026-11-17 23:59** 前完成一个通讯录应用，按评分表可以拆成 6 块：

| 评分项 | 分值 | 对应做法 |
| --- | ---: | --- |
| Markdown 博客 + 仓库/规范链接 | 10 | 本文 + 仓库内 `docs/code-standards.md` |
| 实现思路、截图、分工 | 20 | 第四、五、六节 |
| 页面风格 | 10 | 卡片式双栏、渐变按钮、响应式、Toast 提示 |
| 收藏联系人 | 25 | 星标 + 只看收藏过滤 + 收藏优先排序 |
| 多种联系方式 | 10 | `methods` 一对多数组，任意条数 |
| Excel 导入导出 | 25 | openpyxl 导出样式化 xlsx；导入按表头映射并去重合并 |
| 云端部署（加分） | +15 | 已支持 `--host/--port` 与 `PORT` 变量，可上云 |

## 二、项目链接

| 项目 | 链接 |
| --- | --- |
| GitHub 仓库 | https://github.com/1an666216/address-book-xp |
| 代码规范文档 | https://github.com/1an666216/address-book-xp/blob/main/docs/code-standards.md |
| 开发计划（里程碑与提交计划） | https://github.com/1an666216/address-book-xp/blob/main/docs/workplan.md |
| 公网访问地址（云端部署加分项） | 【替换】http://<公网IP>:<端口>（未部署则写"未部署"） |
| 发布版本 Release | https://github.com/1an666216/address-book-xp/releases/tag/v1.0.0 |
| 演示视频 | 【替换】网盘/哔哩哔哩链接 |

运行方式（评审老师可复现）：

```bash
pip install openpyxl
python seed.py     # 可选：演示数据
python app.py      # 打开 http://127.0.0.1:8000
python tests/e2e_test.py --base http://127.0.0.1:8000   # 27 项验收断言
```

## 三、Git 提交证据

仓库共 **7** 次提交，全部由本人完成。

| 成员 | 学号 | 提交次数 |
| --- | --- | ---: |
| 【替换：你的姓名】 | 【替换：学号】 | 7 |

提交历史截图：

【替换：插入 GitHub 仓库 → Insights → Contributors 或 `git log --oneline --graph` 的截图】

关键提交：

```
chore: project skeleton, README, code standards and PSP
feat(ui): two-column layout, search, toast and responsive styles
feat(bookmark): star toggle, favourites-only filter and favourite-first ordering
feat(api): contact CRUD with atomic JSON persistence
test: 27 e2e assertions covering bookmark, methods and excel round-trip
fix(import): only count contacts that actually changed as updated
```

## 四、实现思路

### 4.1 总体结构

不做前后端分离的重型架构，而是"一个 Python 进程 + 静态页面"，把复杂度留给数据模型：

```
static/index.html · style.css · app.js   ← 单页界面（原生 JS，无构建）
              │ fetch JSON / 上传 xlsx
              ▼
app.py  ── 领域模型 ── JSON 存储（原子写） ── openpyxl 读写 Excel
```

选型理由：作业要的是"能跑、能演示、可复现"，Python 标准库 `http.server` + `openpyxl` 已足够，
评审只要 `python app.py` 就能起服务，不需要 Node、数据库或构建链。

### 4.2 收藏联系人（25 分）

- **数据层**：联系人带 `bookmarked` 布尔字段，`POST /api/contacts/{id}/toggle-bookmark` 取反并落盘。
- **交互层**：列表项右侧显示金色 ★；详情页右上角"收藏/取消收藏"按钮；顶部"只看收藏"按钮切换过滤。
- **排序**：列表按 `(not bookmarked, name)` 排序，收藏的联系人永远浮在最上面——这是"便于快速找到重要联系人"的直接体现。
- **统计**：侧栏实时显示联系人 / 收藏 / 联系方式三个数字。

### 4.3 多种联系方式（10 分）

关键决策是**一个联系人挂一个 `methods` 数组**，每项 `{type, value}`：

```json
"methods": [
  {"type": "phone",  "value": "13800138001"},
  {"type": "email",  "value": "linxy@fzu.edu.cn"},
  {"type": "wechat", "value": "linxy_2024"}
]
```

相比 `phone1/phone2/email1` 这种并列字段，好处是：类型可以无限扩展（加一行 `METHOD_TYPES`
即可让前后端与 Excel 列同步）、同一类型可以有多条、导出时按类型自动聚合到对应列。
编辑弹窗里可以动态"添加一项联系方式"或删除任意一行，保存时统一提交。

### 4.4 Excel 导入导出（25 分）

**导出**：openpyxl 生成 `通讯录` 工作表，一行一个联系人、一类一列
（姓名/收藏/手机/邮箱/微信/QQ/地址/其他联系方式/备注），表头蓝底白字、列宽自适应、冻结首行、自动筛选。
同一人的同类信息用中文分号 `；` 拼在一个单元格里，这样"一行一人"的格式不会被破坏。

**导入**：从表头**按名称**建立列映射（支持 `姓名/name`、`手机/phone`、`邮箱/email` 等别名），
因此列顺序可以打乱、多余列自动忽略；读出每行后按 `；` 拆成多条 `methods`；
再按**姓名去重合并**——新姓名新增，已存在的姓名只追加缺失的联系方式，绝不覆盖原有数据。
返回 `{imported, updated, unchanged, methodsAdded}` 四个数字给界面提示。`updated` 只统计**真正发生变化**的联系人（新增了联系方式/收藏/备注），原样重导入的只计 `unchanged`，避免"导入自己刚导出的文件却提示更新了 5 人"。

**健壮性**：空文件、非 xlsx、缺"姓名"列、没有数据行，全部返回 400 + 中文提示，而不是 500 崩溃。

### 4.5 页面风格（10 分）

双栏卡片布局、`radial-gradient` 背景、圆形彩色头像（按姓名哈希取色）、悬浮阴影、圆角输入框、
收藏开关滑块、底部 Toast 提示、860px 以下自动单列。所有用户数据插入 DOM 前都经过 `escapeHtml()`。

## 五、程序运行截图与演示视频

**主界面**（左：统计 + 列表，右：详情，收藏的联系人带 ★ 且排在最前）：

【替换：插入截图 `docs/screenshots/home.png`】

**新建 / 编辑弹窗**（可动态增删多条联系方式、收藏开关）：

【替换：插入截图 `docs/screenshots/editor.png`】

**Excel 导出样例**：

【替换：插入导出的 xlsx 打开后的截图，或直接上传 `exports/address_book_export.xlsx` 的截图】

**演示视频**（完整走一遍：新建联系人 → 加多条联系方式 → 收藏 → 只看收藏 → 导出 Excel → 修改 Excel 后导入 → 数据合并正确）：

【替换：视频链接 / 嵌入播放器】

## 六、开发职责说明

本项目由本人独立完成，覆盖了从需求到交付的全部环节：

| 环节 | 具体产出 |
| --- | --- |
| 需求拆解 | 对照评分表拆出收藏 / 多联系方式 / Excel 三块功能点 |
| 数据模型与后端 | `app.py`（CRUD、收藏接口、xlsx 读写、去重合并）、`seed.py` |
| 前端界面与交互 | `static/index.html`、`style.css`、`app.js`（渲染/搜索/收藏/导入导出入口） |
| 测试 | `tests/e2e_test.py`（27 项端到端断言） |
| 工程文档 | `README.md`、`docs/code-standards.md`、`docs/workplan.md`、`docs/PSP.md` |
| 演示与博客 | 截图、演示视频、本文档 |

## 7. 贡献比例

| 成员 | 贡献比例 | 依据 |
| --- | ---: | --- |
| 【替换：你的姓名】 | 100% | 需求、设计、前后端编码、测试、文档、视频与博客全部由本人完成 |
| **合计** | **100%** | — |

## 八、开发中的问题与解决

**问题 1：表头解析报 `'str' object has no attribute 'value'`，导入直接 500。**
原因：`ws.iter_rows(values_only=True)` 返回的是纯值，而表头映射函数按单元格对象取 `.value`。
解决：`_header_map` 里兼容两种输入（`cell.value if hasattr(cell, "value") else cell`），
并把解析异常统一包成 `ValueError` → 返回 400 中文提示。

**问题 2：上传非 Excel 文件时服务器 500。**
原因：`load_workbook` 抛 `BadZipFile`，HTTP 层只捕获 `ValueError`。
解决：在 `workbook_to_contacts` 里 try/except 转成 `ValueError`，e2e 脚本补了"垃圾文件必须 400"的断言。

**问题 3：原样重导入刚导出的文件，提示"更新 5 人"。**
原因：旧逻辑只要"匹配到已存在联系人"就算一次更新，哪怕一个字段都没变。
解决：`merge_imported` 增加 `changed` 标志，只有真正新增了联系方式/收藏/备注才计 `updated`，
无变化的计 `unchanged`，界面提示改为"新增 X 人，更新 Y 人，无变化 Z 人"。

**问题 4：导入模板里的"说明文字"被当成联系人导入。**
原因：说明写在数据表 `姓名` 列所在行。
解决：把填写说明移到独立的 `填写说明` sheet，数据表保持纯净。

**问题 5：收藏排序不对，取消收藏的人仍浮在上面。**
原因：排序键写成 `bookmarked` 却按降序排字符串/布尔混用。
解决：统一用 `(not bookmarked, name)` 作为排序键，并对 `bookmarked` 做 `bool()` 归一。

## 九、PSP 表

| PSP 2.1 阶段 | 预估时间(min) | 实际时间(min) | 偏差说明 |
| --- | ---: | ---: | --- |
| 需求分析 | 20 | 20 | 评分表隐藏要求（截图、视频、PSP）比预想多 |
| 开发计划 | 15 | 15 | — |
| 技术方案调研 | 30 | 30 | 对比 xlsxwriter 后选 openpyxl |
| 接口设计 | 20 | 25 | 导入统计需要区分新增/更新/无变化 |
| 设计评审 | 15 | 15 | — |
| 编码：页面骨架与样式 | 60 | 90 | 响应式与弹窗样式反复调整 |
| 编码：列表/详情/搜索 | 45 | 60 | 搜索防抖、空状态 |
| 编码：收藏与过滤 | 30 | 35 | 排序键踩坑 |
| 编码：CRUD 与原子落盘 | 70 | 85 | 加锁与临时文件替换 |
| 编码：多条联系方式 | 45 | 50 | 动态行索引重排 |
| 编码：Excel 导出 | 40 | 55 | 样式与冻结窗格 |
| 编码：Excel 导入 | 60 | 80 | 表头 `.value` 报错、重导入统计 |
| 代码自检（Review） | 20 | 25 | 发现 3 处未转义输入 |
| 测试（e2e + 手工） | 45 | 70 | 27 条断言，修 3 个 bug |
| 耗时记录 | 10 | 10 | — |
| 截图与录屏 | 25 | 40 | 录了两次 |
| 博客撰写 | 60 | 75 | 表格排版耗时 |
| 过程复盘 | 15 | 15 | — |
| **合计** | **625** | **795** | 超时主要在导入健壮性与样式打磨 |

## 十、总结

这次作业最大的收获不是写了多少行代码，而是**先用数据模型把需求"压平"**：
把"多种联系方式"抽象成 `methods` 数组之后，收藏、Excel 列、导入合并都变成对同一个结构的简单操作，
后续加"其他联系方式"几乎没有改动成本。另一个收获是**先写验收脚本**——27 条端到端断言让我在
提交前就发现了导入 500、垃圾文件 500、重导入统计不准这类问题，而不是等到录演示视频时才发现。

## 参考

- openpyxl 文档：<https://openpyxl.readthedocs.io/>
- MDN `fetch` / `dialog`：<https://developer.mozilla.org/zh-CN/docs/Web/API/Fetch_API>
- 《构建之法》（邹欣）第 4 章
