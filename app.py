#v2
import io
import csv
from datetime import datetime
from typing import List

import pandas as pd
import streamlit as st

from pdf_to_heats_xlsx import parse_pdf, build_workbook, Event, AlternateEntry


HEATS_HEADERS = [
    "#",
    "Gender",
    "Event",
    "Age Group",
    "Heat",
    "Cal",
] + [f"Lane {i}" for i in range(10)] + [f"Analyst {i}" for i in range(1, 5)]

ALT_HEADERS = [
    "#",
    "Gender",
    "Event",
    "Age Group",
    "Heat",
    "Alt Group",
    "Alt Rank",
    "Name",
    "Team",
    "Prelims",
]


def events_to_rows(events: List[Event]) -> List[List[str]]:
    rows: List[List[str]] = []
    for ev in events:
        for heat in ev.heats:
            row: List[str] = [
                str(ev.number),
                ev.gender,
                ev.event_code,
                ev.age_group,
                heat.label,
                "",  # Cal
            ]

            for lane in range(10):
                row.append(heat.lanes.get(lane, ""))

            # Analyst columns
            row.extend([""] * 4)
            rows.append(row)

    return rows


def alternates_to_rows(alternates: List[AlternateEntry]) -> List[List[str]]:
    rows: List[List[str]] = []
    for a in alternates:
        rows.append(
            [
                str(a.event_no),
                a.gender,
                a.event_code,
                a.age_group,
                a.heat_label,
                a.alt_group,
                str(a.rank),
                a.name,
                a.team,
                a.prelim,
            ]
        )
    return rows


def rows_to_delimited(headers: List[str], rows: List[List[str]], delimiter: str) -> str:
    buf = io.StringIO()
    writer = csv.writer(buf, delimiter=delimiter, lineterminator="\n")
    writer.writerow(headers)
    writer.writerows(rows)
    return buf.getvalue()


def dataframe_from_rows(headers: List[str], rows: List[List[str]]) -> pd.DataFrame:
    # Ensure consistent row width
    fixed_rows = [r + [""] * (len(headers) - len(r)) for r in rows]
    return pd.DataFrame(fixed_rows, columns=headers)


def copy_all_component(text_to_copy: str, button_label: str, key: str) -> None:
    """Render a small HTML/JS component that copies provided text to clipboard.

    Streamlit doesn't provide a native clipboard API, so we embed a tiny HTML snippet
    that writes TSV text to `navigator.clipboard`.

    To avoid HTML escaping issues, we pass the payload as base64.
    """

    import base64
    import streamlit.components.v1 as components

    payload_b64 = base64.b64encode(text_to_copy.encode("utf-8")).decode("ascii")

    html = f"""
    <div style=\"display:flex; gap:0.5rem; align-items:center; margin: 0.25rem 0 0.75rem 0;\">
      <button id=\"btn-{key}\" style=\"padding:0.4rem 0.8rem; border-radius:0.4rem; border:1px solid #ccc; background:#ffffff; cursor:pointer;\">
        {button_label}
      </button>
      <span id=\"status-{key}\" style=\"font-size:0.9rem; color:#555;\"></span>
    </div>

    <textarea id=\"payload-{key}\" style=\"display:none;\">{payload_b64}</textarea>

    <script>
      const btn = document.getElementById(\"btn-{key}\");
      const status = document.getElementById(\"status-{key}\");
      const b64 = document.getElementById(\"payload-{key}\").value;

      function base64ToUtf8(b64Str) {{
        const binStr = atob(b64Str);
        const bytes = Uint8Array.from(binStr, (m) => m.charCodeAt(0));
        return new TextDecoder("utf-8").decode(bytes);
      }}

      btn.addEventListener("click", async () => {{
        try {{
          const payload = base64ToUtf8(b64);
          await navigator.clipboard.writeText(payload);
          status.textContent = "Copied to clipboard";
          setTimeout(() => status.textContent = "", 2000);
        }} catch (err) {{
          status.textContent = "Copy failed (browser permission)";
        }}
      }});
    </script>
    """

    components.html(html, height=70)


def main() -> None:
    st.set_page_config(page_title="Swim Meet PDF → Heats XLSX", layout="wide")

    st.title("Swim Meet PDF → Heats XLSX")
    st.caption(
        "Upload a meet program PDF, generate the Heats workbook, preview the tables, download XLSX/CSV, or Copy All."
    )

    uploaded = st.file_uploader("Upload a PDF", type=["pdf"], accept_multiple_files=False)

    col1, col2, col3 = st.columns([1, 1, 2])
    with col1:
        preview_rows = st.number_input("Preview rows", min_value=10, max_value=500, value=50, step=10)
    with col2:
        include_alternates = st.checkbox("Include Alternates tab", value=True)
    with col3:
        base_name = "heats"
        if uploaded is not None and getattr(uploaded, "name", None):
            base_name = uploaded.name.rsplit(".", 1)[0]
        out_name = st.text_input("Output filename (without extension)", value=base_name)

    if uploaded is None:
        st.info("Upload a PDF to get started.")
        return

    parse_clicked = st.button("Parse PDF", type="primary")

    if not parse_clicked and "parsed" not in st.session_state:
        # Don’t auto-parse on every rerun; wait for the button.
        return

    if parse_clicked:
        try:
            file_like = io.BytesIO(uploaded.getvalue())
            title, events, alternates = parse_pdf(file_like)
            st.session_state["parsed"] = {
                "title": title,
                "events": events,
                "alternates": alternates,
                "uploaded_name": uploaded.name,
            }
        except Exception as e:
            st.error("Failed to parse PDF")
            st.exception(e)
            return

    parsed = st.session_state.get("parsed")
    if not parsed:
        return

    title = parsed["title"]
    events = parsed["events"]
    alternates = parsed["alternates"]

    st.subheader(title)
    st.write(
        {
            "events": len(events),
            "heats_rows": sum(len(ev.heats) for ev in events),
            "alternates": len(alternates),
            "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        }
    )

    heats_rows = events_to_rows(events)
    heats_df = dataframe_from_rows(HEATS_HEADERS, heats_rows)

    alt_rows = alternates_to_rows(alternates)
    alt_df = dataframe_from_rows(ALT_HEADERS, alt_rows)

    # Build XLSX in memory
    wb = build_workbook(events, alternates, title)
    xlsx_buf = io.BytesIO()
    wb.save(xlsx_buf)
    xlsx_bytes = xlsx_buf.getvalue()

    # Build CSV/TSV
    heats_csv = rows_to_delimited(HEATS_HEADERS, heats_rows, delimiter=",")
    heats_tsv = rows_to_delimited(HEATS_HEADERS, heats_rows, delimiter="\t")

    alt_csv = rows_to_delimited(ALT_HEADERS, alt_rows, delimiter=",")
    alt_tsv = rows_to_delimited(ALT_HEADERS, alt_rows, delimiter="\t")

    st.divider()

    d1, d2, d3, d4 = st.columns([1, 1, 1, 1])
    with d1:
        st.download_button(
            "Download XLSX",
            data=xlsx_bytes,
            file_name=f"{out_name}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    with d2:
        st.download_button(
            "Download Heats CSV",
            data=heats_csv.encode("utf-8"),
            file_name=f"{out_name}_heats.csv",
            mime="text/csv",
        )
    with d3:
        st.download_button(
            "Download Alternates CSV",
            data=alt_csv.encode("utf-8"),
            file_name=f"{out_name}_alternates.csv",
            mime="text/csv",
            disabled=not include_alternates,
        )

    st.divider()

    tabs = st.tabs(["Heats", "Alternates"] if include_alternates else ["Heats"])

    with tabs[0]:
        st.markdown("### Heats Preview")
        st.dataframe(heats_df.head(int(preview_rows)), use_container_width=True, hide_index=True)
        st.markdown("#### Copy all Heats (TSV)")
        copy_all_component(heats_tsv, "Copy all Heats", key="heats")
        with st.expander("Show TSV text (optional)"):
            st.text_area("Heats TSV", heats_tsv, height=200)

    if include_alternates:
        with tabs[1]:
            st.markdown("### Alternates Preview")
            st.dataframe(alt_df.head(int(preview_rows)), use_container_width=True, hide_index=True)
            st.markdown("#### Copy all Alternates (TSV)")
            copy_all_component(alt_tsv, "Copy all Alternates", key="alts")
            with st.expander("Show TSV text (optional)"):
                st.text_area("Alternates TSV", alt_tsv, height=200)


if __name__ == "__main__":
    main()
