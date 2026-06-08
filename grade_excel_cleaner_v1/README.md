# 成绩 Excel 智能清洗 v1.0

本项目用于把非标准成绩 Excel 表格清洗成统一 CSV 结构。LLM 只负责根据表格预览生成 `extraction_plan`，最终数据读取、清洗、校验和导出都由 Python 确定性完成。

## 安装

```bash
pip install -r requirements.txt
```

## 启动

```bash
streamlit run app.py
```

默认配置：

- LLM Base URL: `https://www.inroi.shop`
- Model Name: `gpt-5.4`
- API Key: 从页面密码框填写，或从环境变量 `OPENAI_API_KEY` / Streamlit secrets 读取

## 使用

1. 在侧边栏填写 LLM Base URL、API Key、Model Name。
2. 上传 `.xlsx`、`.xls` 或 `.xlsb` 成绩表。
3. 点击“解析”。
4. 查看 extraction_plan、warnings、清洗结果表格。
5. 点击“下载 CSV”导出。

最终输出列固定为：

- 学号
- 学生姓名
- 课程名
- 班级名
- 最终成绩

## 设计说明

流程为：

1. Python 读取 workbook 并生成 compact preview。
2. LLM 根据 preview 输出严格 JSON 格式的 extraction_plan。
3. Pydantic 校验 extraction_plan。
4. Python 执行计划并清洗数据。
5. 校验输出结果。
6. 如果失败且启用二次修复，将错误、可用列、样例数据和原 preview 反馈给 LLM 重新生成计划。

## 当前 v1.0 限制

- 不保证所有合并单元格和复杂多层表头都能完美识别。
- 不直接支持 PDF。
- 极复杂表格可能需要二次修复或人工检查。
- LLM 结果存在不确定性，所以 Python 会做 schema 和输出校验。
- `.xlsb` 和 `.xls` 读取依赖本地 Python 包可用性，失败时建议另存为 `.xlsx` 后重试。

## v1.1 可扩展方向

- 增加无需 LLM 的规则兜底识别。
- 支持多 sheet 合并输出。
- 展示可编辑 extraction_plan，让用户人工修正后再执行。
- 保存历史解析记录和配置模板。
