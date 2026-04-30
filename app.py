from __future__ import annotations

import csv
import json
import tempfile
from io import StringIO
from pathlib import Path
from typing import Any, Dict, List

import streamlit as st

from analyzer import InputFormatError, InputOptions, REPORT_SHEETS, SUPPORTED_SUFFIXES, analyze_files, normalize_file
from report_exporter import reports_to_excel_bytes


st.set_page_config(
    page_title="Telegram Chat Analyzer",
    page_icon="💬",
    layout="wide",
)

SAMPLE_DIR = Path("sample_data")


def parse_hints(raw: str) -> tuple[str, ...]:
    return tuple(part.strip() for part in raw.split(",") if part.strip())


def save_uploaded_files(uploaded_files) -> List[Path]:
    temp_dir = Path(tempfile.mkdtemp(prefix="telegram_analyzer_"))
    paths: List[Path] = []
    for file in uploaded_files:
        suffix = Path(file.name).suffix.lower()
        if suffix not in SUPPORTED_SUFFIXES:
            continue
        out_path = temp_dir / file.name
        out_path.write_bytes(file.getvalue())
        paths.append(out_path)
    return paths


def get_sample_paths() -> List[Path]:
    if not SAMPLE_DIR.exists():
        return []
    return sorted(SAMPLE_DIR.glob("*.json"))


def rows_to_csv(rows: List[Dict[str, Any]]) -> bytes:
    if not rows:
        return b""
    output = StringIO()
    fieldnames = list(rows[0].keys())
    writer = csv.DictWriter(output, fieldnames=fieldnames)
    writer.writeheader()
    writer.writerows(rows)
    return output.getvalue().encode("utf-8")


def rows_to_json(rows: List[Dict[str, Any]]) -> bytes:
    return json.dumps(rows, ensure_ascii=False, indent=2, default=str).encode("utf-8")


def render_header() -> None:
    st.markdown(
        """
        <style>
        .title {font-size: 2.35rem; font-weight: 900; margin-bottom: 0.15rem;}
        .subtitle {font-size: 1rem; color: #5f6368; margin-bottom: 1rem;}
        .small-note {color: #5f6368; font-size: 0.92rem;}
        </style>
        """,
        unsafe_allow_html=True,
    )
    st.markdown('<div class="title">💬 Telegram Chat Analyzer Website</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="subtitle">Upload Telegram JSON, CSV/XLSX tables, or plain text chats. The website normalizes the input, analyzes manager behavior, and creates an Excel report.</div>',
        unsafe_allow_html=True,
    )


def render_sidebar() -> InputOptions:
    with st.sidebar:
        st.header("Input settings")
        input_format_label = st.selectbox(
            "Input format",
            [
                "Auto detect",
                "Telegram Desktop JSON",
                "CSV table",
                "Excel table",
                "Plain text chat",
            ],
            help="Use Auto detect when your files have normal .json, .csv, .xlsx, or .txt extensions.",
        )
        format_map = {
            "Auto detect": "auto",
            "Telegram Desktop JSON": "telegram_json",
            "CSV table": "csv_table",
            "Excel table": "excel_table",
            "Plain text chat": "plain_text",
        }

        st.divider()
        st.subheader("Manager detection")
        manager_hints = st.text_input(
            "Manager names or keywords",
            value="Manager, Sales, Support",
            help="Comma-separated. Example: Dana, Ruslan, Sales Manager. Used when role column is missing.",
        )
        window_size = st.slider("Messages per mood window", min_value=3, max_value=12, value=6)

        st.divider()
        with st.expander("Field mapping for CSV / Excel .xlsx", expanded=False):
            st.caption("Leave empty for automatic detection.")
            sender_col = st.text_input("Sender column", value="")
            text_col = st.text_input("Text/message column", value="")
            timestamp_col = st.text_input("Timestamp/date column", value="")
            role_col = st.text_input("Role column", value="")
            message_id_col = st.text_input("Message ID column", value="")
            chat_title_col = st.text_input("Chat title column", value="")

        st.divider()
        st.subheader("Supported input examples")
        st.code("Telegram JSON: messages[].from, date, text\nCSV/XLSX: timestamp, sender, role, text\nTXT: 2026-04-21 10:00 - Name: Message", language="text")

    return InputOptions(
        input_format=format_map[input_format_label],
        sender_col=sender_col,
        text_col=text_col,
        timestamp_col=timestamp_col,
        role_col=role_col,
        message_id_col=message_id_col,
        chat_title_col=chat_title_col,
        manager_hints=parse_hints(manager_hints),
        window_size=window_size,
    )


def render_upload_area() -> List[Path]:
    st.subheader("1. Upload chats")
    uploaded_files = st.file_uploader(
        "Upload one or many chat files",
        type=["json", "csv", "xlsx", "txt", "md"],
        accept_multiple_files=True,
    )

    col1, col2 = st.columns([1, 1])
    with col1:
        use_samples = st.checkbox("Use included sample_data instead of upload", value=False)
    with col2:
        st.caption("The included samples are synthetic Telegram Desktop-style JSON chats.")

    if use_samples:
        paths = get_sample_paths()
        if not paths:
            st.error("sample_data folder was not found.")
        else:
            st.success(f"Loaded {len(paths)} sample files.")
        return paths

    if uploaded_files:
        paths = save_uploaded_files(uploaded_files)
        skipped = len(uploaded_files) - len(paths)
        if skipped:
            st.warning(f"Skipped {skipped} unsupported file(s).")
        return paths
    return []


def render_table(title: str, rows: List[Dict[str, Any]]) -> None:
    st.dataframe(rows, use_container_width=True, hide_index=True)
    st.download_button(
        f"Download {title} CSV",
        data=rows_to_csv(rows),
        file_name=f"{title.lower().replace(' ', '_')}.csv",
        mime="text/csv",
        key=f"csv_{title}",
    )


def render_personal_recommendations(rows: List[Dict[str, Any]]) -> None:
    st.dataframe(rows, use_container_width=True, hide_index=True)
    st.download_button(
        "Download Personal recommendations CSV",
        data=rows_to_csv(rows),
        file_name="personal_recommendations.csv",
        mime="text/csv",
        key="csv_personal_recommendations",
    )
    if not rows:
        return
    st.markdown("### Recommendation cards")
    for index, row in enumerate(rows, start=1):
        manager = row.get("manager_name", "Manager")
        priority = row.get("priority_level", "Priority")
        focus = row.get("main_focus", "Main focus")
        with st.expander(f"{index}. {manager} — {priority}: {focus}", expanded=index == 1):
            st.markdown(f"**Strength to keep:** {row.get('strength_to_keep', '')}")
            st.markdown(f"**Personal recommendation:** {row.get('personal_recommendation', '')}")
            st.markdown(f"**Coaching plan:** {row.get('coaching_plan', '')}")
            st.markdown(f"**Example phrase:** {row.get('example_phrase', '')}")


def render_report(reports: Dict[str, List[Dict[str, Any]]]) -> None:
    overview = reports.get("overview", [])
    critical = reports.get("critical_cases", [])
    messages = reports.get("messages", [])

    st.subheader("2. Report")
    if not overview:
        st.warning("No messages were analyzed.")
        return

    total_messages = sum(int(row.get("total_messages", 0) or 0) for row in overview)
    scores = [float(row.get("efficiency_score", 0) or 0) for row in overview]
    avg_score = round(sum(scores) / len(scores), 1) if scores else 0

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Chats", len(overview))
    c2.metric("Messages", total_messages)
    c3.metric("Critical cases", len(critical))
    c4.metric("Avg score", f"{avg_score}%")

    excel_bytes = reports_to_excel_bytes(reports)
    st.download_button(
        "⬇️ Download Excel report",
        data=excel_bytes,
        file_name="telegram_chat_analysis_report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )

    tab_names = ["Overview", "Manager summary", "Personal recommendations", "Mood flow", "Topics", "Critical cases", "Messages", "Input conversion"]
    tabs = st.tabs(tab_names)
    sheet_by_title = {
        "Overview": "overview",
        "Manager summary": "manager_summary",
        "Personal recommendations": "personal_recommendations",
        "Mood flow": "mood_flow",
        "Topics": "topics",
        "Critical cases": "critical_cases",
        "Messages": "messages",
    }
    for tab, title in zip(tabs[:7], tab_names[:7]):
        with tab:
            if title == "Personal recommendations":
                render_personal_recommendations(reports.get(sheet_by_title[title], []))
            else:
                render_table(title, reports.get(sheet_by_title[title], []))

    with tabs[7]:
        st.markdown("**Normalized input** means every input format is converted into one standard table:")
        st.code("source_file, chat_title, message_id, timestamp, sender, role, text", language="text")
        st.dataframe(messages[:200], use_container_width=True, hide_index=True)
        c1, c2 = st.columns(2)
        with c1:
            st.download_button(
                "Download normalized CSV",
                data=rows_to_csv(messages),
                file_name="normalized_messages.csv",
                mime="text/csv",
            )
        with c2:
            st.download_button(
                "Download normalized JSON",
                data=rows_to_json(messages),
                file_name="normalized_messages.json",
                mime="application/json",
            )


def render_input_preview(paths: List[Path], options: InputOptions) -> None:
    if not paths:
        return
    with st.expander("Preview normalized input", expanded=False):
        preview_rows: List[Dict[str, Any]] = []
        for path in paths[:3]:
            try:
                preview_rows.extend(normalize_file(path, options)[:20])
            except Exception as exc:
                st.error(f"{path.name}: {exc}")
        if preview_rows:
            st.dataframe(preview_rows, use_container_width=True, hide_index=True)


def main() -> None:
    render_header()
    options = render_sidebar()
    paths = render_upload_area()
    render_input_preview(paths, options)

    analyze_clicked = st.button("🔍 Analyze chats", type="primary", use_container_width=True, disabled=not paths)
    if not paths:
        st.info("Upload files or select sample_data to start.")
        return

    if analyze_clicked:
        try:
            with st.spinner("Analyzing chats and building report..."):
                reports = analyze_files(paths, options)
            st.session_state["reports"] = reports
            st.success("Analysis completed.")
        except InputFormatError as exc:
            st.error(f"Input format problem: {exc}")
        except Exception as exc:
            st.error(f"Analysis failed: {exc}")

    if "reports" in st.session_state:
        render_report(st.session_state["reports"])


if __name__ == "__main__":
    main()
