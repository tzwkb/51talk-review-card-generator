# A1 Review Card Generator

[English](README.md) | 中文


自动生成 51Talk Business English 课后复习卡（PDF + PNG + HTML）。

## 功能

- **单课卡片**（Lesson Card）：蓝色主题，一课一张
- **单元总结**（Unit Summary）：紫色主题，汇总词汇、语法、成就
- **双语支持**：L1-L3 自动输出英文 + 阿拉伯语
- **批量处理**：支持文件夹批量、Excel 批量
- **增量跳过**：已有 PDF 的课程自动跳过

## 安装

```bash
pip install openai openpyxl pymupdf

export REPORT_API_KEY="your-api-key"
```

确保系统已安装 **Microsoft Edge** 或 **Google Chrome**（用于 HTML → PDF）。

## 项目结构

```
report_generator/
├── main.py              # 入口
├── cli.py               # 交互菜单
├── config.py            # API 配置 + AI Prompts
├── core/                # 核心引擎
│   ├── ai.py            # AI 调用
│   ├── render.py        # HTML 渲染 + 阿拉伯语 RTL
│   ├── export.py        # PDF/PNG 导出
│   ├── record.py        # 记录构建 + 验证
│   ├── pipeline.py      # 单课/单元流程
│   ├── batch.py         # 批量处理
│   └── utils.py         # 通用工具
├── template/
│   ├── lesson_template.html
│   └── unit_template.html
├── test/                # 测试脚本
└── KSA BE 20260324.xlsx # 数据源示例
```

## 使用

### 交互模式（推荐）

```bash
python main.py
```

按菜单选择：
1. 处理单个 Markdown 文件
2. 批量处理文件夹
3. 处理 Excel 文件
4. 生成单元总结

### 命令行模式（Excel）

```bash
python main.py --excel "KSA BE 20260324.xlsx" --output-dir "./output"
```

可选参数：
- `--logo <path>` — 自定义 logo 图片
- `--save-json` — 保存中间 JSON
- `--levels 1 2 3` — 只处理指定级别

### 环境变量

| 变量 | 默认值 | 说明 |
|---|---|---|
| `REPORT_SKIP_EXISTING` | `1` | 已有 PDF 时跳过 |

## 输入格式

### Markdown 文件

文件名格式：`Unit{N}_Lesson{M}.md`

```markdown
### Unit 1: Greetings & Introductions

## Lesson 1: Saying Hello

**Goal:** ...

**Vocabulary**
- Hello — A polite greeting.
- ...
```

### Excel 文件

- **第1列** — Unit 编号
- **第2列** — Lesson 编号
- **第5列** — Markdown 课程内容
- Sheet 名匹配 `^L(\d+)$`（如 `L1`, `L2`）

## 输出

每个课程生成：
- `{lesson_code}.html` — 学习卡片
- `{lesson_code}.pdf` — 扁平化 PDF（适合打印）
- `{lesson_code}.png` — 首页预览
- `{lesson_code}.json` — 结构化数据（可选）

单元总结生成：
- `L{level}/U{unit}/{unit_code}_summary.pdf`

## 测试

```bash
cd test

# 快速渲染测试（不调用 AI）
python test_quick_render.py

# 端到端测试（调用 AI）
python test_real_lesson.py
python test_real_unit_excel.py
```

## 技术栈

- Python 3.13 + asyncio
- OpenAI-compatible API（VectorEngine / Gemini）
- openpyxl（Excel 读取）
- PyMuPDF（PDF 栅格化 + PNG 导出）
- Edge/Chrome Headless（HTML → PDF）

## License

MIT
