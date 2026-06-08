from __future__ import annotations

import os
import tempfile
from pathlib import Path

import pandas as pd
import streamlit as st

from grade_excel_cleaner.planner import run_workflow
from grade_excel_cleaner.preview_builder import preview_to_json


DEFAULT_BASE_URL = "https://www.inroi.shop"
DEFAULT_MODEL = "gpt-5.4"


def _secret_or_env(name: str, default: str = "") -> str:
    try:
        return str(st.secrets.get(name, default))
    except Exception:
        return os.getenv(name, default)


def main() -> None:
    st.set_page_config(page_title="成绩 Excel 智能清洗 v1.0", layout="wide")
    st.title("成绩 Excel 智能清洗 v1.0")

    with st.sidebar:
        st.subheader("LLM 配置")
        base_url = st.text_input("LLM Base URL", value=_secret_or_env("LLM_BASE_URL", DEFAULT_BASE_URL))
        api_key = st.text_input(
            "API Key",
            value=_secret_or_env("OPENAI_API_KEY", ""),
            type="password",
        )
        model = st.text_input("Model Name", value=_secret_or_env("MODEL_NAME", DEFAULT_MODEL))
        st.subheader("高级参数")
        preview_rows = st.number_input("preview rows", min_value=5, max_value=80, value=25, step=5)
        enable_repair = st.checkbox("启用 LLM 二次修复", value=True)

    uploaded = st.file_uploader("上传成绩 Excel 文件", type=["xls", "xlsx", "xlsb"])
    if not uploaded:
        _show_sample_files()
        return

    if st.button("解析", type="primary"):
        if not api_key:
            st.error("请填写 API Key。")
            return
        suffix = Path(uploaded.name).suffix
        with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
            tmp.write(uploaded.getbuffer())
            tmp_path = tmp.name

        try:
            with st.status("正在解析 Excel 并生成 extraction_plan...", expanded=True) as status:
                st.write(f"当前文件：{uploaded.name}")
                result = run_workflow(
                    file_path=tmp_path,
                    preview_rows=int(preview_rows),
                    base_url=base_url,
                    api_key=api_key,
                    model=model,
                    enable_repair=enable_repair,
                )
                status.update(label="解析完成", state="complete")

            st.subheader("解析过程")
            col_left, col_right = st.columns(2)
            with col_left:
                st.write("检测到的 sheet")
                st.json(list(result.workbook.sheets.keys()))
                st.write("读取引擎")
                st.code(result.workbook.engine_used)
                st.write("课程表类型判断")
                st.code(result.execution.table_type_detected)
            with col_right:
                st.write("extraction_plan JSON")
                st.json(result.plan.model_dump())
                if result.repaired:
                    st.info("已使用 LLM 二次修复后的 extraction_plan。")

            warnings = result.execution.warnings + result.workbook.read_warnings
            if warnings:
                st.warning("\n".join(dict.fromkeys(warnings)))

            with st.expander("workbook preview"):
                st.code(preview_to_json(result.preview), language="json")

            st.subheader("清洗结果")
            st.dataframe(result.execution.output, use_container_width=True)
            csv = result.execution.output.to_csv(index=False).encode("utf-8-sig")
            st.download_button(
                "下载 CSV",
                data=csv,
                file_name=f"{Path(uploaded.name).stem}_cleaned.csv",
                mime="text/csv",
            )
        except Exception as exc:
            st.error(f"解析失败：{exc}")
        finally:
            try:
                os.unlink(tmp_path)
            except OSError:
                pass


def _show_sample_files() -> None:
    sample_dir = Path(__file__).resolve().parent.parent / "database"
    if not sample_dir.exists():
        return
    files = sorted(
        path.name
        for path in sample_dir.iterdir()
        if path.suffix.lower() in {".xls", ".xlsx", ".xlsb"}
    )
    if files:
        st.caption("本地测试表格位于 database 目录：")
        st.dataframe(pd.DataFrame({"文件名": files}), use_container_width=True, hide_index=True)


if __name__ == "__main__":
    main()
