# 代码规范（Code Standards）

> 本文件是作业要求中"代码规范文档"的交付物，可直接在博客里贴 GitHub 链接。
> 适用于 `FZU_EE308_XP_AddressBook` 全部源代码（Python + HTML/CSS/JS）。

## 1. 通用原则

1. **可读性优先于技巧**：一段代码若需要注释解释"做了什么"，先考虑重命名。
2. **单一职责**：一个函数只做一件事；超过 ~50 行就拆分。
3. **不重复**：同一段逻辑出现第三次，提取函数或常量。
4. **显式优于隐式**：不依赖隐式类型转换，不吞掉异常（除非有注释说明原因）。
5. **一致性**：同一文件内命名、缩进、引号风格统一，不做无意义的风格混用。

## 2. Python 规范

| 项目 | 约定 |
| --- | --- |
| 基础 | 遵循 PEP 8；每行 ≤ 110 列；文件头 `#!/usr/bin/env python` + `# -*- coding: utf-8 -*-` |
| 命名 | 模块/函数/变量 `snake_case`，常量 `UPPER_SNAKE_CASE`，类 `PascalCase` |
| 私有函数 | 模块内私有函数以 `_` 开头（如 `_clean_methods`、`_header_map`） |
| 类型注解 | 公共函数标注参数与返回值（`def normalize_contact(raw: dict, existing: dict | None = None) -> dict`） |
| 文档字符串 | 每个模块、每个公共函数写一句中文/英文 docstring，说明"为什么"而非"是什么" |
| 异常 | 领域错误抛 `ValueError`，由 HTTP 层统一转 400；不 `except: pass` |
| 日志 | 使用 `self.log_message` 统一输出，禁止调试用 `print` 残留在请求路径中 |
| 导入顺序 | 标准库 → 第三方库（`openpyxl`）→ 本地模块，每组之间空一行 |
| 常量 | 魔法值抽成模块级常量（`METHOD_TYPES`、`EXCEL_HEADERS`、`HEADER_ALIASES`） |

## 3. JavaScript 规范

| 项目 | 约定 |
| --- | --- |
| 基础 | ES2020，无框架、无构建步骤；`const`/`let`，禁止 `var` |
| 命名 | 函数与变量 `camelCase`，类/构造器 `PascalCase`，常量对象 `UPPER_SNAKE_CASE` |
| 分号与引号 | 语句结尾加分号；字符串统一双引号（模板字符串用反引号） |
| 函数 | 优先纯函数；DOM 操作集中在 `render*` / `selectContact` 一类函数中 |
| 转义 | 所有插入 DOM 的用户数据必须经过 `escapeHtml()`，防止 XSS |
| 异步 | 统一 `async/await` + `try/catch`；网络错误必须给用户可见的 Toast 提示 |
| 事件 | 使用 `addEventListener`，禁止内联 `onclick`；事件委托优先 |
| 状态 | 全部可变状态集中在 `state` 对象，修改后调用 `render*()` 重绘 |

## 4. HTML / CSS 规范

- HTML5 语义化标签：`header / main / aside / section / dialog / template`；
- 结构、样式、行为彻底分离：不写行内 `style`（动态颜色等运行时值除外），不写内联事件；
- CSS 使用 `:root` 变量统一色板与圆角（`--brand`、`--line`、`--radius`），禁止散落硬编码颜色；
- 类名 `kebab-case`，按"块-元素"命名（`method-row`、`detail-head`、`editor-foot`）；
- 必须有响应式断点（本项目 `max-width: 860px` 切换单列）。

## 5. 数据与接口规范

- 所有接口返回 JSON：成功 `{"ok": true, ...}`，失败 `{"ok": false, "error": "人类可读信息"}`；
- HTTP 状态码语义化：新建 201，参数错误 400，不存在 404，服务端异常 500；
- 写操作必须校验输入（姓名非空、重名拒绝、方法类型白名单）；
- 数据落盘使用"临时文件 + `os.replace`"的原子写，避免写一半损坏；
- 共享状态（`load_contacts` / `save_contacts`）必须在 `_lock` 内操作。

## 6. 测试规范

- 每个评分点至少一条端到端断言（见 `tests/e2e_test.py`）；
- 测试用例自带清理，跑完把数据还原到基线（不污染演示数据）；
- 断言既检查 HTTP 状态码，也检查业务内容（如导出的表头、行数、合并结果）；
- 提交前必须全绿：`27 passed, 0 failed`。

## 7. 提交信息规范（Conventional Commits）

```
<type>(<scope>): <摘要，祈使句，≤ 50 字>

type: feat | fix | docs | style | refactor | test | chore
scope: bookmark | contact | excel | ui | api | data
```

示例：

- `feat(bookmark): star contacts, filter and favourite-first ordering`
- `feat(excel): styled xlsx export and header-based import with merge`
- `fix(excel): map header row when read with values_only=True`
- `test: e2e coverage for bookmark, methods and excel round-trip`

一次提交只做一件事；不提交 `data/contacts.json`、`__pycache__/`、临时文件（见 `.gitignore`）。
