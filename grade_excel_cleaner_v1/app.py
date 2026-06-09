from __future__ import annotations

import hashlib
import os
import tempfile
import zipfile
from datetime import datetime
from io import BytesIO
from pathlib import Path
from typing import Any

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

from grade_excel_cleaner.planner import WorkflowOutput, run_workflow
from grade_excel_cleaner.preview_builder import preview_to_json
from grade_excel_cleaner.settings import load_app_settings, normalize_base_url
from grade_excel_cleaner.target_workflow import TargetWorkflowOutput, run_target_workflow


MODE_AUTO = "混合类型识别"
MODE_TOTAL = "含课程总分成绩"
MODE_TARGET = "含课程目标成绩"
STATUS_PENDING = "待解析"
STATUS_DONE = "已完成"
STATUS_FAILED = "失败"


def main() -> None:
    st.set_page_config(page_title="成绩 Excel 智能清洗 v2.1", layout="wide")
    _inject_style()
    _inject_beforeunload_warning()
    _init_state()

    start_clicked, clear_clicked = _render_header()
    if clear_clicked:
        _reset_state()
        st.rerun()

    st.segmented_control(
        "表格类型",
        [MODE_AUTO, MODE_TOTAL, MODE_TARGET],
        key="score_mode",
        help="默认混合识别会先用 Python 检测课程目标结构；未命中时再调用 LLM 识别总分成绩表。",
    )
    uploaded_files = _render_upload_section()

    if start_clicked:
        if uploaded_files:
            selected_file_ids = _selected_upload_ids(uploaded_files)
            if selected_file_ids:
                _parse_uploaded_files(uploaded_files, selected_file_ids)
                st.rerun()
            else:
                st.info("没有需要解析的文件。已完成文件默认不会重复解析。")
        else:
            st.warning("请先上传 Excel 文件。")

    records = st.session_state.records
    if not uploaded_files and not records:
        _render_empty_state()
        return

    if records:
        _render_results()


def _init_state() -> None:
    if "records" not in st.session_state:
        st.session_state.records = {}
    if "upload_seed" not in st.session_state:
        st.session_state.upload_seed = 0
    if "active_record_id" not in st.session_state:
        st.session_state.active_record_id = None
    if "settings_loaded" not in st.session_state:
        settings = load_app_settings()
        st.session_state.setting_base_url = settings.base_url
        st.session_state.setting_api_key = settings.api_key
        st.session_state.setting_model = settings.model
        st.session_state.setting_preview_rows = settings.preview_rows
        st.session_state.setting_enable_repair = settings.enable_repair
        st.session_state.setting_api_key_editing = False
        st.session_state.setting_api_key_edit_nonce = 0
        st.session_state.settings_loaded = True
    if "score_mode" not in st.session_state:
        st.session_state.score_mode = MODE_AUTO


def _reset_state() -> None:
    st.session_state.records = {}
    st.session_state.active_record_id = None
    st.session_state.upload_seed += 1


def _runtime_settings() -> dict[str, Any]:
    return {
        "base_url": normalize_base_url(st.session_state.setting_base_url),
        "api_key": st.session_state.setting_api_key,
        "model": st.session_state.setting_model,
        "preview_rows": int(st.session_state.setting_preview_rows),
        "enable_repair": bool(st.session_state.setting_enable_repair),
    }


def _mask_api_key(api_key: str) -> str:
    if not api_key:
        return "未配置"
    return "已配置（隐藏）"


def _inject_beforeunload_warning() -> None:
    components.html(
        """
        <script>
        const message = "离开页面后，本次上传与临时解析结果将消失。";
        window.parent.onbeforeunload = function (event) {
            event.preventDefault();
            event.returnValue = message;
            return message;
        };
        </script>
        """,
        height=0,
    )


def _render_header() -> tuple[bool, bool]:
    title_col, guide_col, clear_col, start_col, settings_col = st.columns([6.5, 1.1, 1.1, 1.2, 1.0])
    with title_col:
        st.markdown(
            """
            <div class="app-title">成绩 Excel 智能清洗 v2.1</div>
            <div class="app-subtitle">批量上传、混合识别、字段溯源与人工审核</div>
            """,
            unsafe_allow_html=True,
        )
    with guide_col:
        with st.popover("使用指南", use_container_width=True):
            st.markdown(
                """
                1. 上传一个或多个 `.xls`、`.xlsx`、`.xlsb` 文件。
                2. 默认使用混合识别；也可以指定总分成绩或课程目标成绩。
                3. 解析后查看字段映射、计算规则和告警，再选择字段导出。
                """
            )
    with clear_col:
        clear_clicked = st.button("清空全部", use_container_width=True)
    with start_col:
        start_clicked = st.button("开始解析", type="primary", use_container_width=True)
    with settings_col:
        with st.popover("设置", use_container_width=True):
            st.text_input("LLM Base URL", key="setting_base_url")
            st.caption(f"API Key：{_mask_api_key(st.session_state.setting_api_key)}")
            if not st.session_state.setting_api_key_editing:
                if st.button("修改 API Key", use_container_width=True):
                    st.session_state.setting_api_key_editing = True
                    st.session_state.setting_api_key_edit_nonce += 1
                    st.rerun()
            else:
                draft_key = f"setting_api_key_draft_{st.session_state.setting_api_key_edit_nonce}"
                draft_api_key = st.text_input("新 API Key", type="password", key=draft_key, placeholder="输入新的 API Key")
                key_cols = st.columns(2)
                with key_cols[0]:
                    if st.button("确认修改", type="primary", use_container_width=True):
                        if draft_api_key:
                            st.session_state.setting_api_key = draft_api_key
                        st.session_state.setting_api_key_editing = False
                        st.session_state.setting_api_key_edit_nonce += 1
                        st.rerun()
                with key_cols[1]:
                    if st.button("取消修改", use_container_width=True):
                        st.session_state.setting_api_key_editing = False
                        st.session_state.setting_api_key_edit_nonce += 1
                        st.rerun()
            st.text_input("Model Name", key="setting_model")
            st.number_input("preview rows", min_value=5, max_value=80, step=5, key="setting_preview_rows")
            st.checkbox("启用 LLM 二次修复", key="setting_enable_repair")
    return start_clicked, clear_clicked


def _render_upload_section() -> list[Any]:
    with st.container(border=True):
        uploaded_files = st.file_uploader(
            "拖拽 Excel 文件到上传框开始",
            type=["xls", "xlsx", "xlsb"],
            accept_multiple_files=True,
            key=f"uploader_{st.session_state.upload_seed}",
            help="支持 .xls / .xlsx / .xlsb，支持批量上传",
        )
        st.caption("支持 .xls / .xlsx / .xlsb，单个文件不超过 200MB。勾选需要解析的文件后点击开始解析。")
        _render_upload_queue(uploaded_files)
        return uploaded_files or []


def _render_upload_queue(uploaded_files: list[Any] | None) -> None:
    uploaded_files = uploaded_files or []
    if uploaded_files:
        records = st.session_state.records
        queue_rows: list[dict[str, Any]] = []
        for uploaded in uploaded_files:
            file_id = _file_id(uploaded)
            record = records.get(file_id)
            status = record["status"] if record else STATUS_PENDING
            key = f"parse_selected_{file_id}"
            if key not in st.session_state:
                st.session_state[key] = status != STATUS_DONE
            if status == STATUS_DONE:
                st.session_state[key] = False
            queue_rows.append(
                {
                    "解析": bool(st.session_state[key]),
                    "文件名": uploaded.name,
                    "大小": _format_size(uploaded.size),
                    "状态": status,
                    "识别类型": record.get("detected_mode", "") if record else "",
                    "_file_id": file_id,
                }
            )

        table_height = min(360, 44 + 44 * len(queue_rows))
        edited = st.data_editor(
            pd.DataFrame(queue_rows),
            column_order=["解析", "文件名", "大小", "状态", "识别类型"],
            hide_index=True,
            use_container_width=True,
            height=table_height,
            disabled=["文件名", "大小", "状态", "识别类型"],
            key="upload_queue_editor",
            column_config={
                "解析": st.column_config.CheckboxColumn("解析", width="small"),
                "文件名": st.column_config.TextColumn("文件名", width="large"),
                "大小": st.column_config.TextColumn("大小", width="small"),
                "状态": st.column_config.TextColumn("状态", width="small"),
                "识别类型": st.column_config.TextColumn("识别类型", width="medium"),
            },
        )
        if isinstance(edited, pd.DataFrame):
            for index, row in edited.iterrows():
                file_id = queue_rows[index]["_file_id"]
                record = records.get(file_id)
                if record and record["status"] == STATUS_DONE:
                    st.session_state[f"parse_selected_{file_id}"] = False
                else:
                    st.session_state[f"parse_selected_{file_id}"] = bool(row.get("解析", False))
    else:
        st.caption("当前未选择文件。")


def _selected_upload_ids(uploaded_files: list[Any]) -> set[str]:
    selected: set[str] = set()
    for uploaded in uploaded_files:
        file_id = _file_id(uploaded)
        record = st.session_state.records.get(file_id)
        if record and record["status"] == STATUS_DONE:
            continue
        if st.session_state.get(f"parse_selected_{file_id}", True):
            selected.add(file_id)
    return selected


def _render_empty_state() -> None:
    st.markdown(
        """
        <div class="empty-state">
            <div class="empty-title">项目仓库</div>
            <a href="https://github.com/duck011002/excelformat" target="_blank">
                https://github.com/duck011002/excelformat
            </a>
            <div class="empty-title second">使用说明</div>
            <div class="empty-text">
                默认混合识别会先检测课程目标结构；没有课程目标结构时，再调用 LLM 识别总分成绩表。
                解析完成后可以查看字段映射、计算规则、异常字段和人工审核建议，并按文件和字段导出结果。
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def _parse_uploaded_files(uploaded_files: list[Any], selected_file_ids: set[str]) -> None:
    settings = _runtime_settings()
    pending_uploads = [
        uploaded
        for uploaded in uploaded_files
        if _file_id(uploaded) in selected_file_ids
        and st.session_state.records.get(_file_id(uploaded), {}).get("status") != STATUS_DONE
    ]
    if not pending_uploads:
        st.info("没有需要解析的文件。已完成文件默认不会重复解析。")
        return

    progress = st.progress(0, text="准备解析")
    for index, uploaded in enumerate(pending_uploads, start=1):
        file_id = _file_id(uploaded)
        st.session_state.records[file_id] = {
            "id": file_id,
            "file_name": uploaded.name,
            "size": uploaded.size,
            "status": "解析中",
            "detected_mode": "",
        }
        progress.progress((index - 1) / len(pending_uploads), text=f"正在解析 {uploaded.name}")
        record = _parse_one_file(uploaded, file_id, settings, st.session_state.score_mode)
        st.session_state.records[file_id] = record
        if record["status"] == STATUS_DONE and st.session_state.active_record_id is None:
            st.session_state.active_record_id = file_id
    progress.progress(1.0, text="解析完成")


def _parse_one_file(uploaded: Any, file_id: str, settings: dict[str, Any], score_mode: str) -> dict[str, Any]:
    suffix = Path(uploaded.name).suffix
    with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
        tmp.write(uploaded.getvalue())
        tmp_path = tmp.name

    try:
        if score_mode == MODE_TARGET:
            result = run_target_workflow(file_path=tmp_path)
            return _record_from_target_result(file_id, uploaded, result, "手动指定：课程目标成绩")
        if score_mode == MODE_TOTAL:
            result = _run_total_workflow(tmp_path, settings)
            return _record_from_total_result(file_id, uploaded, result, "手动指定：课程总分成绩")

        try:
            result = run_target_workflow(file_path=tmp_path)
            return _record_from_target_result(file_id, uploaded, result, "混合识别：Python 检测到课程目标结构")
        except Exception as target_error:
            result = _run_total_workflow(tmp_path, settings)
            record = _record_from_total_result(file_id, uploaded, result, "混合识别：未命中课程目标结构，转入 LLM 总分解析")
            record["warnings"].append(f"课程目标结构检测未通过：{target_error}")
            record["audits"] = _build_audit_issues(record["output"], record["warnings"])
            return record
    except Exception as exc:
        return {
            "id": file_id,
            "file_name": uploaded.name,
            "size": uploaded.size,
            "status": STATUS_FAILED,
            "detected_mode": "未识别",
            "detection_reason": "解析失败，无法判断表格类型。",
            "output": pd.DataFrame(),
            "trace": pd.DataFrame(),
            "warnings": [str(exc)],
            "audits": [
                {
                    "类型": "解析失败",
                    "数量": 1,
                    "级别": "需人工审核",
                    "说明": str(exc),
                    "建议": "确认表格类型、表头和文件格式后重试。",
                }
            ],
            "detail": {},
            "finished_at": _now_text(),
        }
    finally:
        try:
            os.unlink(tmp_path)
        except OSError:
            pass


def _run_total_workflow(tmp_path: str, settings: dict[str, Any]) -> WorkflowOutput:
    if not settings["api_key"]:
        raise ValueError("缺少 API Key，请在设置中填写或检查本地配置文件。")
    return run_workflow(
        file_path=tmp_path,
        preview_rows=settings["preview_rows"],
        base_url=settings["base_url"],
        api_key=settings["api_key"],
        model=settings["model"],
        enable_repair=settings["enable_repair"],
    )


def _record_from_total_result(file_id: str, uploaded: Any, result: WorkflowOutput, detection_reason: str) -> dict[str, Any]:
    warnings = list(dict.fromkeys(result.execution.warnings + result.workbook.read_warnings))
    output = result.execution.output
    trace = _prepend_detection_trace(_trace_total_result(result), MODE_TOTAL, detection_reason)
    return {
        "id": file_id,
        "file_name": uploaded.name,
        "size": uploaded.size,
        "status": STATUS_DONE,
        "detected_mode": MODE_TOTAL,
        "detection_reason": detection_reason,
        "output": output,
        "trace": trace,
        "warnings": warnings,
        "audits": _build_audit_issues(output, warnings),
        "detail": {
            "plan": result.plan.model_dump(),
            "repaired": result.repaired,
            "table_type_detected": result.execution.table_type_detected,
            "workbook_preview": preview_to_json(result.preview),
        },
        "finished_at": _now_text(),
    }


def _record_from_target_result(
    file_id: str,
    uploaded: Any,
    result: TargetWorkflowOutput,
    detection_reason: str,
) -> dict[str, Any]:
    warnings = list(dict.fromkeys(result.warnings + result.workbook.read_warnings))
    output = result.output
    trace = _prepend_detection_trace(_trace_target_result(result), MODE_TARGET, detection_reason)
    return {
        "id": file_id,
        "file_name": uploaded.name,
        "size": uploaded.size,
        "status": STATUS_DONE,
        "detected_mode": MODE_TARGET,
        "detection_reason": detection_reason,
        "output": output,
        "trace": trace,
        "warnings": warnings,
        "audits": _build_audit_issues(output, warnings),
        "detail": {
            "sheet_name": result.plan.sheet_name,
            "header_row_index": result.plan.header_row_index,
            "data_start_row_index": result.plan.data_start_row_index,
            "target_count": len(result.plan.target_groups),
            "total_column": result.plan.total_column,
            "target_source_columns": [
                {
                    "课程目标": group.number,
                    "源列": [_excel_column_name(column) for column in group.source_columns],
                    "源列索引": group.source_columns,
                    "满分": group.denominator,
                    "源表头": group.source_headers,
                }
                for group in result.plan.target_groups
            ],
        },
        "finished_at": _now_text(),
    }


def _render_results() -> None:
    records = st.session_state.records
    ordered_ids = list(records.keys())
    done_ids = [record_id for record_id in ordered_ids if records[record_id]["status"] == STATUS_DONE]
    if st.session_state.active_record_id not in ordered_ids:
        st.session_state.active_record_id = done_ids[0] if done_ids else ordered_ids[0]

    left_col, main_col, side_col = st.columns([1.35, 4.7, 1.95], gap="medium")
    with left_col:
        st.markdown("#### 文件列表")
        with st.container(height=620, border=True):
            st.radio(
                "文件列表",
                ordered_ids,
                format_func=lambda record_id: records[record_id]["file_name"],
                key="active_record_id",
                label_visibility="collapsed",
            )

    active_record = records[st.session_state.active_record_id]
    with main_col:
        _render_active_record(active_record)
    with side_col:
        _render_side_panel(active_record, done_ids)


def _render_active_record(record: dict[str, Any]) -> None:
    title_cols = st.columns([5, 1.4])
    with title_cols[0]:
        st.markdown(f"### {record['file_name']}")
        st.caption(f"{record['detected_mode']} · {record['finished_at']}")
    with title_cols[1]:
        status = record["status"]
        st.markdown(f'<div class="result-status">{status}</div>', unsafe_allow_html=True)

    if record["status"] != STATUS_DONE:
        st.error(record["warnings"][0] if record["warnings"] else "解析失败。")
        return

    output = record["output"]
    review_count = _review_count(record)
    metric_cols = st.columns(4)
    metric_cols[0].metric("识别行数", len(output))
    metric_cols[1].metric("有效字段", len(output.columns))
    metric_cols[2].metric("异常字段", _issue_count(record, "异常"))
    metric_cols[3].metric("待人工审核", review_count)

    tab_preview, tab_trace, tab_raw = st.tabs(["数据预览", "字段映射", "原始预览"])
    with tab_preview:
        st.dataframe(output, use_container_width=True, hide_index=True, height=360)
        st.caption(f"显示前 {min(len(output), 500)} 行，共 {len(output)} 行")
    with tab_trace:
        st.dataframe(record["trace"], use_container_width=True, hide_index=True, height=320)
    with tab_raw:
        st.json(record["detail"])


def _render_side_panel(record: dict[str, Any], done_ids: list[str]) -> None:
    st.markdown("#### 告警与审核")
    _render_audit_panel(record)
    st.markdown("#### 导出")
    _render_export_panel(record, done_ids)


def _render_audit_panel(record: dict[str, Any]) -> None:
    issues = record.get("audits", [])
    if record["status"] != STATUS_DONE:
        st.error(record["warnings"][0] if record["warnings"] else "解析失败。")
        return
    if not issues:
        with st.container(height=290, border=True):
            st.success("无待处理告警")
        return
    with st.container(height=290, border=True):
        for issue in issues:
            level = issue.get("级别", "")
            if level == "需人工审核":
                st.error(f"{issue['类型']} · {issue['数量']}")
            elif level == "警告":
                st.warning(f"{issue['类型']} · {issue['数量']}")
            else:
                st.info(f"{issue['类型']} · {issue['数量']}")
            st.caption(issue.get("说明", ""))


def _render_export_panel(record: dict[str, Any], done_ids: list[str]) -> None:
    if record["status"] != STATUS_DONE:
        with st.container(border=True):
            st.caption("当前文件解析完成后可导出。")
        return

    with st.container(border=True):
        payload, file_name, mime = _build_export_payload(
            [record["id"]],
            list(record["output"].columns),
            "xlsx",
            f"{Path(record['file_name']).stem}_cleaned.xlsx",
        )
        st.download_button(
            "直接导出",
            data=payload,
            file_name=file_name,
            mime=mime,
            type="primary",
            use_container_width=True,
        )
        if st.button("选择性导出", use_container_width=True):
            _selective_export_dialog(record["id"], done_ids)


@st.dialog("选择性导出")
def _selective_export_dialog(active_id: str, done_ids: list[str]) -> None:
    records = st.session_state.records
    default_ids = [active_id] if active_id in done_ids else done_ids[:1]
    export_ids = st.multiselect(
        "导出文件",
        done_ids,
        default=default_ids,
        format_func=lambda record_id: records[record_id]["file_name"],
    )
    available_columns = _ordered_union_columns([records[record_id]["output"] for record_id in export_ids])
    selected_columns = st.multiselect("导出字段", available_columns, default=available_columns)
    export_format = st.selectbox("导出格式", ["xlsx", "csv"])
    default_name = "grade_cleaned.xlsx" if len(export_ids) != 1 else f"{Path(records[export_ids[0]]['file_name']).stem}_cleaned.{export_format}"
    output_name = st.text_input("文件名", value=default_name)
    payload, file_name, mime = _build_export_payload(export_ids, selected_columns, export_format, output_name)
    st.download_button(
        "确认导出",
        data=payload,
        file_name=file_name,
        mime=mime,
        type="primary",
        disabled=not export_ids or not selected_columns or not output_name,
        use_container_width=True,
    )
    if st.button("关闭", use_container_width=True):
        st.rerun()


def _trace_total_result(result: WorkflowOutput) -> pd.DataFrame:
    plan = result.plan
    rows = [
        _trace_row("学号", plan.column_mapping.student_id, "LLM 字段映射", "自动映射"),
        _trace_row("学生姓名", plan.column_mapping.student_name, "LLM 字段映射", "自动映射"),
        _trace_row(
            "课程名",
            plan.column_mapping.course_name or plan.metadata_mapping.course_name.source,
            plan.metadata_mapping.course_name.reason or "LLM 元数据推断",
            "自动映射" if plan.column_mapping.course_name else "元数据推断",
        ),
        _trace_row(
            "班级名",
            plan.column_mapping.class_name or plan.metadata_mapping.class_name.source,
            plan.metadata_mapping.class_name.reason or "LLM 元数据推断",
            "自动映射" if plan.column_mapping.class_name else "元数据推断",
        ),
        _trace_row("最终成绩", plan.column_mapping.final_score, plan.final_score_reason or "LLM 字段映射", "自动映射"),
    ]
    if result.repaired:
        rows.append(_trace_row("extraction_plan", "LLM 二次修复", "首次计划执行失败后重试", "需关注"))
    return pd.DataFrame(rows)


def _prepend_detection_trace(trace: pd.DataFrame, detected_mode: str, detection_reason: str) -> pd.DataFrame:
    detection = pd.DataFrame(
        [
            _trace_row(
                "表格类型",
                detected_mode,
                detection_reason,
                "自动判断" if detection_reason.startswith("混合识别") else "手动指定",
            )
        ]
    )
    return pd.concat([detection, trace], ignore_index=True)


def _trace_target_result(result: TargetWorkflowOutput) -> pd.DataFrame:
    rows = [
        _trace_row("学号", _excel_column_name(result.plan.student_id_column), "Python 表头匹配：学号/学生学号", "自动映射"),
        _trace_row("姓名", _excel_column_name(result.plan.student_name_column), "Python 表头匹配：姓名/学生姓名", "自动映射"),
    ]
    for group in sorted(result.plan.target_groups, key=lambda item: item.number):
        source = ", ".join(_excel_column_name(column) for column in group.source_columns)
        rule = "按满分折算百分制" if group.denominator is not None else "单列比例转百分制或直接取分"
        if len(group.source_columns) > 1 and group.denominator is None:
            rule = "多列求和"
        rows.append(
            _trace_row(
                f"课程目标{group.number}",
                source,
                f"{rule}；源表头：{' / '.join(group.source_headers) or '未命名'}",
                "自动计算",
            )
        )
    total_source = _excel_column_name(result.plan.total_column) if result.plan.total_column is not None else "课程目标均权"
    total_rule = "源表总分列" if result.plan.total_column is not None else "未识别总分列，按课程目标均权兜底"
    total_status = "自动映射" if result.plan.total_column is not None else "需关注"
    rows.append(_trace_row("总分", total_source, total_rule, total_status))
    return pd.DataFrame(rows)


def _trace_row(field: str, source: Any, rule: str, status: str) -> dict[str, str]:
    return {
        "输出字段": field,
        "来源": clean_text(source),
        "规则/证据": clean_text(rule),
        "状态": status,
    }


def _build_audit_issues(output: pd.DataFrame, warnings: list[str]) -> list[dict[str, Any]]:
    issues: list[dict[str, Any]] = []
    for warning in warnings:
        issues.append(
            {
                "类型": "解析告警",
                "数量": 1,
                "级别": "警告",
                "说明": warning,
                "建议": "查看字段映射和原始预览，确认解析规则是否符合预期。",
            }
        )

    if output.empty:
        issues.append(
            {
                "类型": "空结果",
                "数量": 1,
                "级别": "需人工审核",
                "说明": "解析没有生成可导出的记录。",
                "建议": "检查上传文件、表头行和表格类型。",
            }
        )
        return issues

    for column in output.columns:
        missing = int(output[column].map(lambda value: clean_text(value) == "").sum())
        if missing:
            issues.append(
                {
                    "类型": f"{column} 空值",
                    "数量": missing,
                    "级别": "需人工审核" if column in {"学号", "姓名", "学生姓名"} else "警告",
                    "说明": f"{column} 存在 {missing} 条空值。",
                    "建议": "回到原表补齐或确认是否应剔除。",
                }
            )

    id_column = "学号" if "学号" in output.columns else None
    if id_column:
        student_ids = output[id_column].map(clean_text)
        non_blank_ids = student_ids[student_ids != ""]
        duplicate_mask = non_blank_ids.duplicated(keep=False)
        duplicate_rows = int(duplicate_mask.sum())
        duplicate_values = int(non_blank_ids[duplicate_mask].nunique())
        if duplicate_rows:
            issues.append(
                {
                    "类型": "学号重复",
                    "数量": duplicate_rows,
                    "级别": "需人工审核",
                    "说明": f"发现 {duplicate_values} 个重复学号，涉及 {duplicate_rows} 条记录。",
                    "建议": "确认是否存在重修、跨班或重复录入。",
                }
            )

    score_columns = [column for column in output.columns if column not in {"学号", "姓名", "学生姓名", "课程名", "班级名"}]
    for column in score_columns:
        numeric = pd.to_numeric(output[column], errors="coerce")
        original_not_empty = output[column].map(lambda value: clean_text(value) != "")
        non_numeric = int((numeric.isna() & original_not_empty).sum())
        out_of_range = int(((numeric < 0) | (numeric > 100)).sum())
        if non_numeric:
            issues.append(
                {
                    "类型": f"{column} 非数值",
                    "数量": non_numeric,
                    "级别": "需人工审核",
                    "说明": f"{column} 有 {non_numeric} 条不是数字。",
                    "建议": "确认等级制成绩是否需要换算或人工保留。",
                }
            )
        if out_of_range:
            issues.append(
                {
                    "类型": f"{column} 超出范围",
                    "数量": out_of_range,
                    "级别": "需人工审核",
                    "说明": f"{column} 有 {out_of_range} 条不在 0-100 范围。",
                    "建议": "检查是否为比例、满分值不同或录入错误。",
                }
            )

    return issues


def _build_export_payload(
    export_ids: list[str],
    selected_columns: list[str],
    export_format: str,
    output_name: str | None = None,
) -> tuple[bytes, str, str]:
    records = st.session_state.records
    export_records = [records[record_id] for record_id in export_ids if records[record_id]["status"] == STATUS_DONE]
    if not export_records or not selected_columns:
        return b"", "empty.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

    if export_format == "csv":
        if len(export_records) == 1:
            df = _select_existing_columns(export_records[0]["output"], selected_columns)
            return (
                df.to_csv(index=False).encode("utf-8-sig"),
                _ensure_extension(output_name or f"{Path(export_records[0]['file_name']).stem}_cleaned.csv", "csv"),
                "text/csv",
            )
        buffer = BytesIO()
        with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as zipped:
            for record in export_records:
                df = _select_existing_columns(record["output"], selected_columns)
                zipped.writestr(f"{Path(record['file_name']).stem}_cleaned.csv", df.to_csv(index=False).encode("utf-8-sig"))
        return buffer.getvalue(), _ensure_extension(output_name or "grade_cleaned_csv.zip", "zip"), "application/zip"

    buffer = BytesIO()
    used_names: set[str] = set()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        for record in export_records:
            df = _select_existing_columns(record["output"], selected_columns)
            sheet_name = _safe_sheet_name(Path(record["file_name"]).stem, used_names)
            df.to_excel(writer, index=False, sheet_name=sheet_name)
    return (
        buffer.getvalue(),
        _ensure_extension(
            output_name
            or ("grade_cleaned.xlsx" if len(export_records) > 1 else f"{Path(export_records[0]['file_name']).stem}_cleaned.xlsx"),
            "xlsx",
        ),
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


def _select_existing_columns(output: pd.DataFrame, selected_columns: list[str]) -> pd.DataFrame:
    columns = [column for column in selected_columns if column in output.columns]
    return output[columns] if columns else output.iloc[:, :0]


def _ensure_extension(file_name: str, extension: str) -> str:
    file_name = (file_name or f"grade_cleaned.{extension}").strip()
    suffix = f".{extension.lower()}"
    if not file_name.lower().endswith(suffix):
        file_name = f"{file_name}{suffix}"
    return file_name


def _ordered_union_columns(outputs: list[pd.DataFrame]) -> list[str]:
    columns: list[str] = []
    for output in outputs:
        for column in output.columns:
            if column not in columns:
                columns.append(column)
    return columns


def _safe_sheet_name(name: str, used_names: set[str]) -> str:
    cleaned = re_sub_sheet_name(name)[:31] or "Sheet"
    candidate = cleaned
    index = 2
    while candidate in used_names:
        suffix = f"_{index}"
        candidate = f"{cleaned[: 31 - len(suffix)]}{suffix}"
        index += 1
    used_names.add(candidate)
    return candidate


def re_sub_sheet_name(name: str) -> str:
    for char in ["\\", "/", "*", "?", ":", "[", "]"]:
        name = name.replace(char, "_")
    return name


def _review_count(record: dict[str, Any]) -> int:
    return sum(int(issue["数量"]) for issue in record.get("audits", []) if issue.get("级别") == "需人工审核")


def _issue_count(record: dict[str, Any], keyword: str) -> int:
    return sum(int(issue["数量"]) for issue in record.get("audits", []) if keyword in issue.get("类型", "") or issue.get("级别") == "需人工审核")


def _file_id(uploaded: Any) -> str:
    digest = hashlib.sha1(uploaded.getvalue()).hexdigest()[:12]
    return f"{uploaded.name}:{uploaded.size}:{digest}"


def _format_size(size: int) -> str:
    if size < 1024:
        return f"{size} B"
    if size < 1024 * 1024:
        return f"{size / 1024:.1f} KB"
    return f"{size / 1024 / 1024:.1f} MB"


def _excel_column_name(index: int | None) -> str:
    if index is None:
        return ""
    number = index + 1
    letters = ""
    while number:
        number, remainder = divmod(number - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters


def _now_text() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def clean_text(value: Any) -> str:
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except TypeError:
        pass
    text = str(value).strip()
    return "" if text.lower() in {"nan", "none", "nat"} else text


def _inject_style() -> None:
    st.markdown(
        """
        <style>
        header[data-testid="stHeader"] { height: 0rem; background: transparent; pointer-events: none; }
        div[data-testid="stToolbar"] { display: none; }
        .block-container { padding-top: 1.4rem; max-width: 1500px; }
        .app-title { font-size: 2.15rem; font-weight: 760; color: #10203F; line-height: 1.1; }
        .app-subtitle { color: #64748B; margin-top: .45rem; font-size: 1rem; }
        div[data-testid="stFileUploader"] section[data-testid="stFileUploaderDropzone"] {
            min-height: 150px;
            border: 1.5px dashed #22C55E;
            border-radius: 18px;
            background: #F6FEF9;
            align-items: center;
            justify-content: center;
        }
        div[data-testid="stFileUploader"] section[data-testid="stFileUploaderDropzone"] > div {
            align-items: center;
            text-align: center;
        }
        div[data-testid="stFileUploader"] section[data-testid="stFileUploaderDropzone"] button {
            min-width: 120px;
            min-height: 42px;
            position: relative;
            z-index: 2;
        }
        div[data-testid="stFileUploader"] [data-testid="stFileChips"] {
            display: flex;
            flex-wrap: wrap;
            justify-content: flex-start;
            gap: .75rem;
            width: 100%;
        }
        div[data-testid="stFileUploader"] [data-testid="stFileChip"] {
            max-width: 430px;
            min-height: 76px;
            border: 1.5px solid #22C55E;
            border-radius: 18px;
            background: #F0FDF4;
            box-shadow: 0 8px 20px rgba(34, 197, 94, .08);
            padding: .75rem .9rem;
        }
        div[data-testid="stFileUploader"] [data-testid="stFileChip"] button,
        div[data-testid="stFileUploader"] [data-testid="stFileChipDeleteBtn"] button {
            border-radius: 999px;
        }
        div[data-testid="stFileUploader"] [data-testid="stFileChipName"] {
            color: #10203F;
            font-weight: 650;
        }
        div[data-testid="stFileUploader"] [data-testid="stFileChip"] small,
        div[data-testid="stFileUploader"] [data-testid="stFileChip"] div:nth-child(2) {
            color: #047857;
            font-weight: 600;
        }
        div[data-testid="stFileUploader"] [data-testid="stBaseButton-borderlessIcon"] {
            border: 1.5px solid #22C55E;
            border-radius: 999px;
            background: #ECFDF5;
            color: #047857;
        }
        div[data-testid="stDataFrame"] {
            border: 1px solid #BBF7D0;
            border-radius: 16px;
            overflow: hidden;
            box-shadow: 0 8px 24px rgba(15, 118, 110, .06);
            background: #FFFFFF;
        }
        div[data-testid="stDataFrame"] canvas {
            border-radius: 16px;
        }
        div[data-testid="stElementToolbar"] {
            display: none;
        }
        .upload-shell {
            border: 1.5px dashed #8BB4FF;
            border-radius: 8px;
            padding: 1rem;
            margin-top: 1rem;
            background: #FBFDFF;
        }
        .empty-state {
            margin-top: 1rem;
            border: 1px solid #E2E8F0;
            border-radius: 8px;
            padding: 1.2rem 1.4rem;
            background: #FFFFFF;
        }
        .empty-title { font-weight: 720; color: #0F172A; margin-bottom: .35rem; }
        .empty-title.second { margin-top: 1rem; }
        .empty-text { color: #475569; max-width: 760px; line-height: 1.7; }
        .file-status { font-size: .78rem; margin: -.35rem 0 .55rem 1.75rem; }
        .file-status.ok { color: #047857; }
        .file-status.bad { color: #B91C1C; }
        .file-status.wait { color: #64748B; }
        .result-status {
            text-align: center;
            color: #047857;
            background: #E8F7EF;
            border-radius: 6px;
            padding: .4rem .5rem;
            font-weight: 700;
        }
        div[data-testid="stMetric"] {
            border: 1px solid #E2E8F0;
            border-radius: 8px;
            padding: .85rem 1rem;
            background: #FFFFFF;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


if __name__ == "__main__":
    main()
