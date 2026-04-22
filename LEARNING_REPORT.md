# Report Generator — 本地化工程师学习报告

> 你是谁：Localization Engineer，vibecoding 背景，维护一套 51Talk Business English 课后复习卡的生成系统。

---

## 一、这个项目到底是干什么的

**一句话**：把课程原始资料（Excel / Markdown）自动转成漂亮的复习卡片（PDF + PNG + HTML）。

**输入**：
- Excel 工作簿（`KSA BE 20260324.xlsx`）— 每行是一课，第5列是课程内容
- Markdown 文件（`Unit1_Lesson1.md`）— 课程原始文本

**输出**：
- `.html` — 可直接在浏览器打开的学习卡片
- `.pdf` — 扁平化后的卡片，适合打印/发家长
- `.png` — 卡片首页预览图
- `.json` — AI 提取的结构化数据（方便调试）

**两种卡片类型**：
1. **Lesson Card**（单课）— 蓝色主题，一课一张
2. **Unit Summary**（单元总结）— 紫色主题，汇总整个单元的词汇、语法、成就

---

## 二、项目结构（重构后 — 4 文件方案）

```
report_generator/
├── main.py              # 入口 + CLI 交互菜单（非交互 Excel 批量也在这）
├── ai.py                # AI 调用 + JSON 解析 + 并发安全打印
├── engine.py            # 渲染、导出、记录、流程、批量、工具（合并原 core/ 6 个模块）
├── config.py            # API 配置 + AI System Prompts
├── template/
│   ├── lesson_template.html
│   └── unit_template.html
└── test/
    └── ...
```

**为什么从 11 个文件压到 4 个？**
- 原 `core/` 下 6 个模块高度耦合（pipeline 调 render，render 调 utils，batch 调 pipeline），分开反而增加跳转成本
- `cli.py` 和 `main.py` 始终一起改，合并减少一层委托
- `engine.py` 按「函数职责」分组（用注释分隔），而非「文件边界」，定位更快

---

## 三、技术栈全景

| 技术 | 作用 | 你维护时需要懂什么 |
|---|---|---|
| **Python 3.13** | 主语言 | async/await、pathlib |
| **openai** | 调用 AI 接口 | OpenAI-compatible SDK，不是官方 OpenAI |
| **openpyxl** | 读 Excel | 只读，不修改原文件 |
| **PyMuPDF (fitz)** | PDF 处理 + PNG 导出 | 把 PDF 转成图片、栅格化 |
| **Edge/Chrome Headless** | HTML → PDF | 命令行参数、浏览器路径查找 |
| **HTML + CSS** | 卡片样式 | 打印分页、RTL 右对齐、flex/grid |

**没有用的技术**：没有数据库、没有 Web 框架、没有 Docker。

---

## 四、核心数据流

```
Excel 第5列 或 Markdown 文件
    │
    ▼
┌─────────────────┐
│  拼接双语指令   │  ← engine.py process_file / process_unit
│                 │     L1-L3 强制要求 AI 输出英阿双语
└────────┬────────┘
         │
         ▼
┌─────────────────┐
│  调用 AI (GPT)  │  ← ai.py extract_json_from_markdown
│                 │     System Prompt 控制输出格式
└────────┬────────┘
         │
         ▼
┌─────────────────┐
│   JSON 规范化   │  ← engine.py build_record / build_unit_record
│   (normalize)   │
└────────┬────────┘
         │
         ▼
┌─────────────────┐
│   渲染 HTML     │  ← engine.py render_html
│  (模板 + 数据)  │     注入 RTL 阿拉伯语、logo base64
└────────┬────────┘
         │
         ▼
┌─────────────────┐
│  Edge → PDF     │  ← engine.py generate_pdf
│  (headless)     │     浏览器把 HTML "打印" 成 PDF
└────────┬────────┘
         │
    ┌────┴────┐
    ▼         ▼
┌───────┐ ┌────────┐
│flatten│ │generate│
│  PDF  │ │  PNG   │
└───────┘ └────────┘
```

**关键认知**：AI 生成 JSON → 程序转成 HTML → 浏览器把 HTML 转成 PDF。

---

## 五、每个模块是干什么的

### `main.py` — 入口 + CLI
- `main()`：`argparse` 解析 → 非交互 Excel 模式 或 交互菜单模式
- `_menu()`：交互模式选择器（单文件 / 批量文件夹 / Excel / 单元总结）
- 所有 CLI 交互函数：`_mode_single_file`、`_mode_batch_folder`、`_mode_excel`、`_mode_unit_summary`

### `ai.py` — AI 大脑
- `extract_json_from_markdown()`：调用 API，带 3 次重试和退避延迟
  - JSON 解析错误退避 `[5, 15, 30]` 秒
  - API 错误退避 `[10, 30, 60]` 秒
- `_parse_json_response()`：去掉 AI 返回的 ` ```json ` 代码块 fence
- `tprint()`：带锁的并发安全打印（`_PRINT_LOCK`）

### `engine.py` — 整个引擎（6 大职责，按代码顺序）

**① 工具函数**
- `logo_data_uri()`：把 logo PNG 转 base64 data URI
- `infer_course()`：从文件名推断 BE/GE
- `existing_lessons()`：扫描已有 PDF，支持增量跳过
- `CONCURRENCY_LIMIT`：批量并发上限（默认 12）

**② RTL 双语排版**
- `_rtl_arabic_in_html()`：检测阿拉伯语字符，包上 `<span dir="rtl">`
- 区分「卡片模式」（`compact=True`，紧凑无空行）和「段落模式」（双语对之间插 `<br>`）

**③ 渲染**
- `render_html()`：把 JSON 数据塞进 HTML 模板，双兼容（同时传 lesson 和 unit 的所有变量）
- `_esc_br()`：`html.escape()` 后恢复 `<br>` 标签

**④ 导出**
- `generate_pdf()`：Edge headless `--print-to-pdf`**（必须用绝对路径 `.resolve()`）**
- `_flatten_pdf()`：**栅格化** — 把 PDF 每页渲染成 2× 高清图片再拼回 PDF，防字体丢失
- `generate_png()`：PyMuPDF 导出首页 PNG

**⑤ 记录构建与验证**
- `build_record()` / `build_unit_record()`：把 AI 的脏 JSON 洗成统一字典
- `validate()`：检查缺失标题、禁用词（Teacher:, Option A/B 等）
- `extract_unit_title()`：从 markdown 里提取单元标题

**⑥ 流程与批量**
- `process_file()`：单课完整流水线（读文件 → AI → 渲染 → PDF/PNG → 验证）
- `process_unit()`：单元总结完整流水线（含双语注入 + Next Unit Preview）
- `process_excel_cell()`：Excel 单单元格处理
- `_run_batch()`：文件夹批量处理，带 `Semaphore(CONCURRENCY_LIMIT)` 并发控制
- `_run_excel_batch()`：Excel 批量处理
- `write_report()`：生成 `_batch_report.json` / `.md`

---

## 六、5 个设计亮点

### 1. PDF 扁平化（Rasterization）
浏览器生成 PDF 后，`_flatten_pdf()` 用 PyMuPDF 把每一页渲染成 2× 高清图片，再拼回新 PDF。结果是纯图片 PDF，任何设备打开都长一样。

### 2. 双语 = Prompt Engineering
不另外调翻译 API。直接在 System Prompt 里告诉 AI "L1-L3 必须输出阿拉伯语"。AI 自己生成双语内容，程序只做 RTL 排版。

### 3. 模板双兼容
`render_html()` 同时传入所有可能的变量（`lesson_code`, `unit_code`, `word_count` 等），一个函数同时兼容 lesson 和 unit 两个模板，不用 if/else 分支。

### 4. 并发控制
批量处理时 `Semaphore(12)` 限制并发数。阻塞操作（浏览器子进程、PyMuPDF）用 `asyncio.to_thread()` 包一层，避免卡住事件循环。

### 5. Virtual Time Budget
Chrome flag `--virtual-time-budget=5000` 强制浏览器在捕获 PDF 前完成所有布局和字体加载，防止布局闪烁。

---

## 七、本地化工程师必须掌握的 Python 知识点

### 7.1 async/await
```python
async def process_file(...):
    ai_data = await extract_json_from_markdown(...)   # 网络请求 → await
    html_text = render_html(record, template)           # 纯计算 → 直接做
    await asyncio.to_thread(generate_pdf, ...)          # 阻塞操作 → to_thread
```

### 7.2 正则表达式
- 从文件名提取 Level/Unit/Lesson：`re.match(r"^L(\d+)$", ...)`
- 检测阿拉伯语：`[\u0600-\u06FF]`
- 拆 `<br>` 标签：`re.split(r"(<br\s*/?>)", html_text)`

### 7.3 pathlib（不要再用 os.path）
```python
output_dir = Path("test_output")
output_dir.mkdir(exist_ok=True)
html_path = output_dir / "test.html"
```

### 7.4 模板字符串
本项目用 Python 原生 `str.format()`，不是 Jinja2：
```html
<div class="card-title">{title}</div>
```
```python
template.format(title="Greetings & Introductions")
```

### 7.5 Unicode 和编码
- 阿拉伯语在 `\u0600-\u06FF`
- 文件读写必须写 `encoding="utf-8"`
- Windows 默认 `cp1252`，不写会炸

---

## 八、vibecoding 最容易踩的坑

| 坑 | 为什么 | 怎么避免 |
|---|---|---|
| 相对路径导致 PDF 静默失败 | Edge `--print-to-pdf` 要绝对路径 | 代码里用 `pdf_path.resolve()` |
| `<br>` 被 html.escape 转义 | `html.escape("<br>")` → `&lt;br&gt;` | `_esc_br()` 把 `&lt;br&gt;` 转回来 |
| AI 返回包在代码块里 | AI 经常输出 ` ```json {...} ``` ` | `_parse_json_response()` 去掉 fence |
| 阿拉伯语没有右对齐 | 默认 LTR，句号出现在句首 | `dir="rtl"` + `text-align: right` |
| 卡片被分页截断 | PDF 打印时卡片上下分在两页 | CSS `break-inside: avoid` |
| 浏览器找不到 | 不同电脑路径不同 | `_find_browser()` 多路径探测 |

---

## 九、改系统从哪下手

### 改样式
- `template/lesson_template.html` — 单课卡片
- `template/unit_template.html` — 单元总结

### 改 AI 输出内容
- `config.py` — `SYSTEM_PROMPT` / `UNIT_SYSTEM_PROMPT`

### 改输出格式 / 字段
- `engine.py` — `build_record()` / `build_unit_record()` 改 JSON 规范化字段
- `engine.py` — `render_html()` 改模板变量注入逻辑

### 改导出行为
- `engine.py` — `generate_pdf()` / `_flatten_pdf()` / `generate_png()`

### 加新语种
- `engine.py` — 扩展 `_ARABIC_RE` 正则匹配新语种
- `engine.py` — `process_file()` / `process_unit()` 添加新的 level 判断
- 模板 CSS — 加 `direction` 和 `text-align` 规则

---

## 十、文件速查表

| 文件 | 什么时候打开它 |
|---|---|
| `main.py` | 改菜单、改交互流程、改 CLI 参数 |
| `ai.py` | 改重试逻辑、改 AI 调用参数、改打印行为 |
| `engine.py` | **最常打开** — 改渲染、改导出、改记录字段、改批量逻辑、改工具函数 |
| `config.py` | 改 AI 提示词、改 API 配置 |
| `template/*.html` | 改颜色、改布局、改卡片大小 |

---

## 十一、学习优先级

1. **最高**：理解 `async/await` + `asyncio.to_thread()` — 主流程骨架
2. **高**：理解 Prompt → JSON → HTML → PDF 流水线 — 项目骨架
3. **中**：理解 `engine.py` 的双语排版 — 本地化核心
4. **中**：理解 `break-inside: avoid` 和 `page-break-after: avoid` — PDF 分页
5. **低**：理解 `_flatten_pdf()` 栅格化 — 知道有这回事即可
