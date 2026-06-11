# 成绩 Excel 智能清洗 v2.2

本项目用于把不同格式的成绩 Excel 表格清洗成统一结构。当前支持两类课程表格：

- 含课程总分成绩：输出 `学号`、`学生姓名`、`课程名`、`班级名`、`最终成绩`
- 含课程目标成绩：输出 `学号`、`姓名`、`课程目标1-n`、`总分`

## v2.1 功能

- 支持 `.xls`、`.xlsx`、`.xlsb` 批量拖拽上传。
- 表格类型可选 `混合类型识别`、`含课程总分成绩`、`含课程目标成绩`。
- 混合识别会优先用 Python 检测课程目标结构，再由 LLM 判断字段映射和复杂规则。
- 已完成文件默认不会重复解析，可勾选多个待解析文件后统一解析。
- 文件列表固定展示解析结果，切换文件不会重复弹出导出窗口。
- 支持直接导出当前文件，也支持选择性导出多个文件、指定字段、确认文件名和导出格式。
- 字段映射、计算规则、识别依据和异常字段会在页面中展示，便于人工审核。
- 生成结果只保存在当前 Streamlit 会话内存中，关闭或离开页面后会消失，不会在本地保存导出文件。

## v2.2 教学班清洗

- 主模式调整为 `仅解析成绩`、`解析成绩和教学班`，成绩类型通过旁边下拉框选择。
- 教学班固定导出 8 列：`教学班编号`、`教学班名称`、`运行学年`、`学年内学期`、
  `任课老师名称`、`任课老师工号`、`教学班学生姓名`、`教学班学生学号`。
- Python 先从文件名、sheet 名、表头、表尾和关键词邻近单元格提取有限证据，不把整张大表交给 LLM。
- 学生姓名和学号直接复用已清洗的成绩结构；缺失时仍可导出，空列会保留并显示告警。
- 教学班按课程、课程班别名、显式教学班、行政班顺序分组；相同名称但不同编号保持独立编辑。
- 原始教学班编号支持字母数字；缺失时生成跨重复解析稳定的正整数兜底编号。
- 人工修改会锁定对应字段；`识别不准？AI增加再次解析` 只复核未锁定、低置信度字段。
- 学期统一为 `第一学期`、`第二学期`、`第三学期`，并支持 `2023秋`、`2024春`、`23-24-2` 等写法。

## TODO

- 教学班结果批量导出。
- 多文件教学班结果合并到同一个工作簿的多个 sheet。

## 安装

建议使用项目的 conda 环境：

```bash
cd D:\project\zhihuishu\excelformat\grade_excel_cleaner_v1
conda activate excel
pip install -r requirements.txt
```

## 启动

在项目目录 `D:\project\zhihuishu\excelformat\grade_excel_cleaner_v1` 下启动：

```bash
conda activate excel
python -m streamlit run app.py
```

如果需要指定端口：

```bash
python -m streamlit run app.py --server.port 8501 --server.address localhost
```

启动后浏览器打开 `http://localhost:8501/`。如果 8501 已被占用，把 `--server.port` 后面的数字换成其他端口。

## 本地配置

LLM 配置从 `config/local_settings.json` 读取，也可以在页面右上角“设置”里临时修改。该文件已加入 `.gitignore`，不要提交真实 API Key。

示例：

```json
{
  "base_url": "https://llm-service.polymas.com/api/openai/v1",
  "api_key": "YOUR_API_KEY",
  "model": "gpt-5.4",
  "preview_rows": 25,
  "enable_repair": true
}
```

也可以使用环境变量覆盖：

- `OPENAI_BASE_URL`
- `OPENAI_API_KEY`
- `OPENAI_MODEL`
- `GRADE_CLEANER_PREVIEW_ROWS`
- `GRADE_CLEANER_ENABLE_REPAIR`

## 使用流程

1. 打开页面后上传一个或多个 Excel 文件。
2. 选择解析内容和成绩解析类型，默认仅解析成绩并使用混合类型识别。
3. 勾选需要解析的文件，点击“开始解析”。
4. 在结果区查看数据预览、字段映射、识别依据和告警审核。
5. 成绩结果可直接或选择性导出；组合模式可额外导出固定 8 列教学班结果。

## LLM 使用方式

LLM 只负责理解表格结构、字段映射和复杂合并规则，不直接处理整张工作簿。程序会先由 Python 生成压缩预览，包括候选表头、少量样例行、sheet 信息和结构线索，再交给 LLM 生成解析计划。最终的数据读取、字段提取、总分计算、校验和导出都由 Python 执行。

这样可以降低上下文长度，也便于追踪每个字段和规则的来源。

## 测试

```bash
python tests/test_basic.py
python tests/test_teaching_class.py
```

真实样本回归：

```powershell
$env:GRADE_CLEANER_DATABASE="D:\project\zhihuishu\excelformat\database"
$env:GRADE_CLEANER_TEACHING_TEMPLATE="D:\Users\13952\Downloads\教学班模版 (2).xlsx"
python tests/test_teaching_class_samples.py
```
