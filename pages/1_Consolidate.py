from io import BytesIO
from typing import Dict, List, Optional, Tuple

import openpyxl
from openpyxl.utils import column_index_from_string
import streamlit as st

from filesplit import (
    get_column_letter_by_header,
    sanitize_prefix,
    _canonicalize_lead,
    _create_zip_from_workbooks,
)


def _load_uploaded_files(uploaded_files) -> List[Tuple[str, bytes]]:
    """Extract byte content from uploaded files once."""
    buffers: List[Tuple[str, bytes]] = []
    for uploaded in uploaded_files:
        buffers.append((uploaded.name, uploaded.getvalue()))
    return buffers


@st.cache_data(show_spinner=False)
def consolidate_entity_workbooks(
    uploaded_files: List[Tuple[str, bytes]], header_label: str
) -> Tuple[List[str], Dict[str, bytes], Dict[str, List[str]]]:
    """Return per-entity consolidated workbooks along with missing sheet details."""
    entity_data: Dict[str, Dict[str, Dict[str, List[List[Optional[object]]]]]] = {}
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

            col_index = column_index_from_string(col_letter)
            header_values = [
                worksheet.cell(row=header_row, column=col).value for col in range(1, worksheet.max_column + 1)
            ]

            for row_index in range(header_row + 1, worksheet.max_row + 1):
                entity_value = _canonicalize_lead(worksheet.cell(row=row_index, column=col_index).value)
                if not entity_value:
                    continue

                row_values = [
                    worksheet.cell(row=row_index, column=col).value for col in range(1, worksheet.max_column + 1)
                ]

                sheet_bucket = entity_data.setdefault(entity_value, {}).setdefault(
                    sheet_name, {"header": header_values, "rows": []}
                )
                sheet_bucket["rows"].append(row_values)

        workbook.close()

    sorted_entities = sorted(entity_data.keys())
    consolidated_workbooks: Dict[str, bytes] = {}

    for entity, sheets in entity_data.items():
        output_wb = openpyxl.Workbook()
        first_sheet = True
        for sheet_name, payload in sheets.items():
            if first_sheet:
                ws_out = output_wb.active
                ws_out.title = sheet_name
                first_sheet = False
            else:
                ws_out = output_wb.create_sheet(title=sheet_name)

            ws_out.append(payload["header"])
            for row in payload["rows"]:
                ws_out.append(row)

        buffer = BytesIO()
        output_wb.save(buffer)
        output_wb.close()
        consolidated_workbooks[entity] = buffer.getvalue()

    return sorted_entities, consolidated_workbooks, missing_by_file


def main():
    st.title("File Consolidator")
    st.write(
        "Combine multiple workbooks and export aggregated files for each Team Lead or Mentor. "
        "Great for merging prior exports back into a single, curated set."
    )

    st.sidebar.header("How to use")
    st.sidebar.write(
        "1. Choose whether to consolidate by Team Lead or Mentor and set an optional filename prefix.\n"
        "2. Upload one or more workbooks previously exported or structured identically.\n"
        "3. Press **Consolidate files** to merge the data.\n"
        "4. Review the generated files and download them individually or as a zip."
    )

    default_filter = st.session_state.get("consolidate_filter_by_mentor", False)
    toggle_label = f"Consolidate by {'Mentor' if default_filter else 'Team Lead'}"
    filter_by_mentor = st.toggle(toggle_label, value=default_filter, key="consolidate_filter_by_mentor")
    target_header = "Mentor" if filter_by_mentor else "Team Lead"

    prefix_input = st.text_input(
        "Optional filename prefix",
        value="",
        placeholder="e.g., Consolidated 2025_Team ",
        help="Prefix added ahead of each generated filename.",
        key="consolidate_prefix_input",
    )
    sanitized_prefix = sanitize_prefix(prefix_input)

    uploaded_files = st.file_uploader(
        "Upload one or more workbooks (.xlsx) to consolidate", type=["xlsx"], accept_multiple_files=True
    )

    file_buffers: Optional[List[Tuple[str, bytes]]] = None
    result_key = None
    if uploaded_files:
        file_buffers = _load_uploaded_files(uploaded_files)
        result_key = (
            tuple((name, len(content)) for name, content in file_buffers),
            target_header,
            sanitized_prefix,
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
            leads, workbooks, missing_sheets = consolidate_entity_workbooks(file_buffers, target_header)

        if not leads:
            st.error(f"No {target_header.lower()}s were found across the uploaded workbooks.")
            st.session_state.pop("consolidation_results", None)
            return

        stored_results = {
            "leads": leads,
            "workbooks": workbooks,
            "missing_sheets": missing_sheets,
            "target_header": target_header,
            "prefix": sanitized_prefix,
            "key": result_key,
        }
        st.session_state["consolidation_results"] = stored_results
    elif should_show_results:
        stored_results = st.session_state["consolidation_results"]
        leads = stored_results["leads"]
        workbooks = stored_results["workbooks"]
        missing_sheets = stored_results["missing_sheets"]
        target_header = stored_results["target_header"]
        sanitized_prefix = stored_results["prefix"]
    else:
        if not uploaded_files:
            st.session_state.pop("consolidation_results", None)
            st.info("Start by uploading the workbooks you want to merge.")
        else:
            st.info("Click **Consolidate files** to merge the uploaded workbooks.")
            if stored_results and stored_results.get("key") != result_key:
                st.warning("Inputs changed. Click **Consolidate files** to refresh.")
        return

    if stored_results.get("missing_sheets"):
        missing_messages = [
            f"- **{filename}**: {', '.join(sheets)}" for filename, sheets in stored_results["missing_sheets"].items()
        ]
        st.warning(
            "Some uploaded workbooks did not contain the selected column and were skipped:\n" + "\n".join(missing_messages)
        )

    st.success(f"Generated consolidated files for {len(stored_results['leads'])} {target_header.lower()}s.")
    st.write("Download the combined files below.")

    zip_bytes = _create_zip_from_workbooks(stored_results["workbooks"], sanitized_prefix)
    st.download_button(
        label="Download consolidated files as ZIP",
        data=zip_bytes,
        file_name=f"consolidated-{target_header.lower().replace(' ', '-')}.zip",
        mime="application/zip",
        use_container_width=True,
    )

    st.subheader("Individual downloads")
    columns = st.columns(2)
    for index, lead in enumerate(stored_results["leads"]):
        column = columns[index % len(columns)]
        column.download_button(
            label=f"Download {lead}",
            data=stored_results["workbooks"][lead],
            file_name=f"{sanitized_prefix}{lead}.xlsx" if sanitized_prefix else f"{lead}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=f"consolidated_{lead}",
            use_container_width=True,
        )


if __name__ == "__main__":
    main()
