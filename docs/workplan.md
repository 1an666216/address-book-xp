# 开发计划（Work Plan）

> 作业要求“先计划再动手”：本文件记录里程碑、提交计划与验收清单，提交进仓库作为过程证据。
> 课程：EE308FZ 软件工程（2601_FZU-MU_SE）· 作业：极限编程 · 截止：2026-11-17 23:59
> 提交形式：独立完成（单人）。

## 1. 职责范围（独立完成）

| 模块 | 主要文件 | 说明 |
| --- | --- | --- |
| 数据模型与后端 | `app.py`、`seed.py` | 联系人 CRUD、收藏接口、Excel 导入导出、去重合并、原子落盘 |
| 前端界面与交互 | `static/index.html`、`style.css`、`app.js` | 双栏布局、列表/详情渲染、搜索、收藏交互、导入导出入口 |
| 测试 | `tests/e2e_test.py`、`tools/make_template.py` | 27 项端到端断言、导入模板生成 |
| 文档 | `README.md`、`docs/*.md` | 代码规范、PSP、代码讲解 |

## 2. 里程碑

| 时间 | 事项 | 产出 |
| --- | --- | --- |
| 11-15 | 需求拆解、数据模型、页面骨架与样式 | 前端提交 |
| 11-15 | 后端接口与 Excel 导入导出 | 后端提交 |
| 11-15 | 端到端测试跑通 | 测试提交 |
| 11-16 | 录制演示视频、博客定稿、修复导入统计 | 视频链接、`BLOG_DRAFT.md` |
| 11-16 | 合并 PR、发布 `v1.0.0` Release | PR + Release 链接 |
| 11-17 | 复核：仓库公开、链接可访问、视频可播、PSP 完整；提交博客链接 | 作业提交完成 |

## 3. 提交计划（Conventional Commits）

| # | 提交信息 | 涉及文件 |
| --- | --- | --- |
| 1 | `chore: project skeleton, README, code standards and PSP` | `.gitignore`、`README.md`、`docs/`、`Dockerfile`、`start.bat` |
| 2 | `feat(ui): two-column layout, search, toast and responsive styles` | `static/index.html`、`static/style.css` |
| 3 | `feat(bookmark): star toggle, favourites-only filter and favourite-first ordering` | `static/app.js` |
| 4 | `feat(api): contact CRUD with atomic JSON persistence` | `app.py`、`seed.py` |
| 5 | `test: 27 e2e assertions covering bookmark, methods and excel round-trip` | `tests/e2e_test.py`、`tools/make_template.py` |
| 6 | `fix(import): only count contacts that actually changed as updated` | `app.py`、`static/app.js` |

> 一次提交只做一件事；提交前跑 `python tests/e2e_test.py`，必须全绿。

## 4. 开发约定

1. **分支命名**：`feat/xxx`、`fix/xxx`、`docs/xxx`。
2. **一次提交只做一件事**，提交信息用祈使句、≤ 50 字。
3. **发布流程**：功能分支 → Pull Request → 合并 → 打 tag → Release。
4. **不可提交**：`data/contacts.json`（含演示数据）、`__pycache__/`（已在 `.gitignore`）。

## 5. 验收清单

- [ ] `python app.py` 启动后首页正常、能增删改查
- [ ] 收藏：★ 切换、只看收藏、收藏优先排序
- [ ] 多联系方式：一个联系人可挂 4 条以上、可增删
- [ ] Excel 导出：一行一人、一类一列、表头齐全
- [ ] Excel 导入：新增 + 合并、无变化不重复计、坏文件返回 400
- [ ] `python tests/e2e_test.py --base http://127.0.0.1:8000` → 27 passed
- [ ] 仓库 Public、≥3 次提交
- [ ] Release `v1.0.0` 已发布、PR 已合并
