import streamlit as st

from report_logic import (
    build_inputting_report,
    build_reconciliation_report,
    build_variance_report,
)
from ann_b3 import build_ann_b3_report

st.set_page_config(page_title="Stock Report Tool", page_icon="📊", layout="centered")

st.title("Stock Report Tool")
st.write(
    "Choose a report, upload your raw system export (.xlsx), and download the result. "
    "Large files (tens of thousands of rows) are fine — this may take up to a "
    "minute or two depending on file size, so please wait for it to finish."
)

MODE_OPTIONS = {
    "Inputting": "Pulls every article with stock on hand (SOH ≠ 0) from the raw export, ready for the manager to fill in physical counts.",
    "Reconciliation": "Collapses the raw report down to one row per article code, summing size-level values.",
    "Variance Report": "Builds the full 3-sheet workbook (Stock, Variance, Variance Report) with offsetting pairs removed and totals.",
    "Ann. B3": "Builds the per-store Ann B3 stocktake reconciliation workbook, grouped by category with Auditor/Manager/Reconciliation totals.",
}

mode = st.radio(
    "Report type",
    list(MODE_OPTIONS.keys()),
    captions=list(MODE_OPTIONS.values()),
)

uploaded = st.file_uploader("Upload your raw report (.xlsx)", type=["xlsx"])

if uploaded is not None:
    st.caption(f"Selected: {uploaded.name} ({uploaded.size / 1e6:.1f} MB)")

generate = st.button("Generate report", type="primary", disabled=uploaded is None)

if generate and uploaded is not None:
    file_bytes = uploaded.read()

    with st.spinner(f"Building your {mode.lower()} report… this can take a minute for large files."):
        try:
            diagnostic_text = None

            if mode == "Inputting":
                buf = build_inputting_report(file_bytes)
                filename = "soh_for_input.xlsx"
                extra_note = None
            elif mode == "Reconciliation":
                buf = build_reconciliation_report(file_bytes)
                filename = "reconciliation_by_code.xlsx"
                extra_note = None
            elif mode == "Variance Report":
                buf, pairs_removed = build_variance_report(file_bytes)
                filename = "variance_report.xlsx"
                extra_note = f"Removed {pairs_removed} offsetting pair(s) from the Variance Report sheet."
            else:  # Ann. B3
                buf, diagnostic_text = build_ann_b3_report(file_bytes)
                filename = "ann_b3_report.xlsx"
                extra_note = None

            st.success("Done!")
            if extra_note:
                st.info(extra_note)

            st.download_button(
                label="⬇ Download result",
                data=buf,
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

            if diagnostic_text:
                with st.expander("Show processing details (what was kept, dropped, and why)"):
                    st.code(diagnostic_text, language=None)

        except ValueError as e:
            st.error(
                f"Couldn't process this file: {e}\n\n"
                "Double-check that your file has the expected column headers "
                "(Code, SOH, Counted, Retail Price, etc.)."
            )
        except Exception as e:
            st.error(f"Something went wrong while processing your file: {e}")

