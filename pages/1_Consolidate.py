from io import BytesIO
from typing import Dict, List, Optional, Sequence, Tuple

import openpyxl
import streamlit as st

from filesplit import get_column_letter_by_header


def _load_uploaded_files(uploaded_files) -> Sequence[Tuple[str, bytes]]:
    """Extract byte content from uploaded files once."""
    buffers: List[Tuple[str, bytes]] = []
    for uploaded in uploaded_files:
        buffers.append((uploaded.name, uploaded.getvalue()))
    return tuple(buffers)


@st.cache_data(show_spinner=False)
def build_consolidated_workbook(
    uploaded_files: Sequence[Tuple[str, bytes]], header_label: str
) -> Tuple[Optional[bytes], Dict[str, List[str]], List[str], int]:
    """
    Merge the provided workbooks into a single workbook.

    Returns the combined workbook bytes, any sheets missing the target header,
    the ordered list of sheet names in the output, and the total merged row count.
    """
    sheet_data: Dict[str, Dict[str, object]] = {}
    missing_by_file: Dict[str, List[str]] = {}
    header_variants = [header_label, f"{header_label}s"]

    for filename, file_bytes in uploaded_files:
        workbook = openpyxl.load_workbook(BytesIO(file_bytes))
        for sheet_name in workbook.sheetnames:
            worksheet = workbook[sheet_name]
            col_letter, header_row = get_column_letter_by_header(worksheet, header_variants)
            if not col_letter:
                missing_by_file.setdefault(filename, []).append(sheet_name)
                continue

            header_values = [
                worksheet.cell(row=header_row, column=col).value for col in range(1, worksheet.max_column + 1)
            ]
            column_count = worksheet.max_column

            sheet_bucket = sheet_data.setdefault(
                sheet_name,
                {"header": header_values, "rows": [], "max_cols": column_count},
            )

            # Expand stored metadata if we encounter more columns in later files.
            if column_count > sheet_bucket["max_cols"]:
                diff = column_count - sheet_bucket["max_cols"]
                sheet_bucket["max_cols"] = column_count
                sheet_bucket["header"] = header_values
                for existing_row in sheet_bucket["rows"]:
                    existing_row.extend([None] * diff)

            max_cols = sheet_bucket["max_cols"]
            for row_index in range(header_row + 1, worksheet.max_row + 1):
                row_values = [
                    worksheet.cell(row=row_index, column=col).value if col <= column_count else None
                    for col in range(1, max_cols + 1)
                ]
                if all(value is None for value in row_values):
                    continue
                sheet_bucket["rows"].append(row_values)

        workbook.close()

    if not sheet_data:
        return None, missing_by_file, [], 0

    sorted_sheet_names = sorted(sheet_data.keys())
    output_wb = openpyxl.Workbook()
    first_sheet = True

    for sheet_name in sorted_sheet_names:
        payload = sheet_data[sheet_name]
        header = payload["header"]
        rows = payload["rows"]

        if first_sheet:
            ws_out = output_wb.active
            ws_out.title = sheet_name
            first_sheet = False
        else:
            ws_out = output_wb.create_sheet(title=sheet_name)

        ws_out.append(header)
        for row in rows:
            ws_out.append(row)

    buffer = BytesIO()
    output_wb.save(buffer)
    output_wb.close()

    total_rows = sum(len(payload["rows"]) for payload in sheet_data.values())
    return buffer.getvalue(), missing_by_file, sorted_sheet_names, total_rows


def main():
    st.title("Consolidate")
    st.write(
        "Combine multiple workbooks into a single file keyed by Team Lead or Mentor. "
        "Perfect for recreating a master workbook from individual exports."
    )

    st.sidebar.header("How to use")
    st.sidebar.write(
        "1. Choose whether to consolidate by Team Lead or Mentor and provide an optional output name.\n"
        "2. Upload one or more workbooks previously exported or structured identically.\n"
        "3. Press **Consolidate files** to merge the data.\n"
        "4. Download the single combined workbook."
    )

    default_filter = st.session_state.get("consolidate_filter_by_mentor", False)
    toggle_label = f"Consolidate by {'Mentor' if default_filter else 'Team Lead'}"
    filter_by_mentor = st.toggle(toggle_label, value=default_filter, key="consolidate_filter_by_mentor")
    target_header = "Mentor" if filter_by_mentor else "Team Lead"

    output_name_input = st.text_input(
        "Output workbook name",
        value="",
        placeholder="e.g., Consolidated September 2025",
        help="The generated file name will append .xlsx automatically if you omit it.",
        key="consolidate_output_name",
    )

    uploaded_files = st.file_uploader(
        "Upload one or more workbooks (.xlsx) to consolidate", type=["xlsx"], accept_multiple_files=True
    )

    file_buffers: Optional[Sequence[Tuple[str, bytes]]] = None
    result_key = None
    if uploaded_files:
        file_buffers = _load_uploaded_files(uploaded_files)
        result_key = (
            file_buffers,
            target_header,
            output_name_input.strip(),
        )

    stored_results = st.session_state.get("consolidation_results")
    should_show_results = stored_results is not None and stored_results.get("key") == result_key

    submitted = st.button("Consolidate files", use_container_width=True)

    if submitted:
        if not file_buffers:
            st.error("Upload at least one workbook before consolidating.")
            st.session_state.pop("consolidation_results", None)
            return

        with st.spinner("Consolidating workbooks..."):
            workbook_bytes, missing_sheets, sheet_names, row_count = build_consolidated_workbook(
                file_buffers, target_header
            )

        if not workbook_bytes:
            st.error(
                f"No {target_header.lower()} data was found across the uploaded workbooks. "
                "Confirm the selected column header exists."
            )
            st.session_state.pop("consolidation_results", None)
            return

        stored_results = {
            "workbook": workbook_bytes,
            "missing_sheets": missing_sheets,
            "sheet_names": sheet_names,
            "row_count": row_count,
            "target_header": target_header,
            "key": result_key,
        }
        st.session_state["consolidation_results"] = stored_results
    elif should_show_results:
        stored_results = st.session_state["consolidation_results"]
    else:
        if not uploaded_files:
            st.session_state.pop("consolidation_results", None)
            st.info("Start by uploading the workbooks you want to merge.")
        else:
            st.info("Click **Consolidate files** to merge the uploaded workbooks.")
            if stored_results and stored_results.get("key") != result_key:
                st.warning("Inputs changed. Click **Consolidate files** to refresh.")
        return

    missing_sheets = stored_results.get("missing_sheets", {})
    if missing_sheets:
        missing_messages = [
            f"- **{filename}**: {', '.join(sheets)}" for filename, sheets in missing_sheets.items()
        ]
        st.warning(
            "Some uploaded workbooks did not contain the selected column and were skipped:\n"
            + "\n".join(missing_messages)
        )

    sheet_names = stored_results.get("sheet_names", [])
    row_count = stored_results.get("row_count", 0)
    st.success(
        f"Merged {row_count} rows across {len(sheet_names)} sheet(s) "
        f"for {target_header.lower()} consolidation."
    )

    output_name = output_name_input.strip() or f"consolidated-{target_header.lower().replace(' ', '-')}"
    if not output_name.lower().endswith(".xlsx"):
        output_name += ".xlsx"

    st.download_button(
        label=f"Download {output_name}",
        data=stored_results["workbook"],
        file_name=output_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )


if __name__ == "__main__":
    main()
