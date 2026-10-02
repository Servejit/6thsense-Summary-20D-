import streamlit as st
import openpyxl
from openpyxl import load_workbook
from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter
from io import BytesIO


st.set_page_config(
    page_title="Excel Summary Generator",
    page_icon="📊",
    layout="wide"
)

st.title("📊 Excel Summary Generator")
st.write("Upload your Master Excel file to generate the Summary file.")


# ============================================================
# SAFE SHEET REFERENCE
# ============================================================

def ref(ws, col):
    name = ws.title.replace("'", "''")
    return f"'{name}'!{col}:{col}"


# ============================================================
# SUM I
# ============================================================

def sum_i_expression(row, sheets20):

    parts = []

    for ws in sheets20:
        parts.append(
            f'SUMIF({ref(ws,"A")},A{row},{ref(ws,"I")})'
        )

    return "=" + "+".join(parts)


# ============================================================
# C-B / AVG.4
# ============================================================

def cb_expression(ws, row):

    return (
        f'IFERROR(('
        f'XLOOKUP(A{row},{ref(ws,"A")},{ref(ws,"C")})-'
        f'XLOOKUP(A{row},{ref(ws,"A")},{ref(ws,"B")})'
        f')/'
        f'XLOOKUP(A{row},{ref(ws,"A")},{ref(ws,"B")})*100,'
        f'""'
        f')'
    )


# ============================================================
# D-B / AVG.4
# ============================================================

def db_expression(ws, row):

    return (
        f'IFERROR(('
        f'XLOOKUP(A{row},{ref(ws,"A")},{ref(ws,"D")})-'
        f'XLOOKUP(A{row},{ref(ws,"A")},{ref(ws,"B")})'
        f')/'
        f'XLOOKUP(A{row},{ref(ws,"A")},{ref(ws,"B")})*100,'
        f'""'
        f')'
    )


# ============================================================
# F-B / AVG.4
# ============================================================

def fb_expression(ws, row):

    return (
        f'IFERROR(('
        f'XLOOKUP(A{row},{ref(ws,"A")},{ref(ws,"F")})-'
        f'XLOOKUP(A{row},{ref(ws,"A")},{ref(ws,"B")})'
        f')/'
        f'XLOOKUP(A{row},{ref(ws,"A")},{ref(ws,"B")})*100,'
        f'""'
        f')'
    )


# ============================================================
# VOL.EXPAND
#
# Vol.Expand 1:
# Sheet 1 J - Average(Sheet 2,3,4 J)
#
# Vol.Expand 2:
# Sheet 2 J - Average(Sheet 3,4,5 J)
# ============================================================

def volume_expression(first_ws, average_sheets, row):

    first_volume = (
        f'SUMIF('
        f'{ref(first_ws,"A")},'
        f'A{row},'
        f'{ref(first_ws,"J")}'
        f')'
    )

    average_values = []

    for ws in average_sheets:

        average_values.append(
            f'SUMIF('
            f'{ref(ws,"A")},'
            f'A{row},'
            f'{ref(ws,"J")}'
            f')'
        )

    average_expression = (
        "(" +
        "+".join(average_values) +
        f")/{len(average_values)}"
    )

    return (
        f'=IFERROR('
        f'{first_volume}-({average_expression}),'
        f'""'
        f')'
    )


# ============================================================
# PROCESS FILE
# ============================================================

def process_excel(uploaded_file):

    input_bytes = uploaded_file.getvalue()

    wb = load_workbook(BytesIO(input_bytes))

    # --------------------------------------------------------
    # DELETE OLD SUMMARY
    # --------------------------------------------------------
    if "summary" in wb.sheetnames:
        del wb["summary"]

    sheets = wb.worksheets[:]

    if len(sheets) < 20:
        raise ValueError(
            f"At least 20 sheets are required. Found only {len(sheets)}."
        )

    first_sheet = sheets[0]
    sheets20 = sheets[:20]
    sheets10 = sheets[:10]
    sheets5 = sheets[:5]

    START_ROW = 2
    END_ROW = 211

    summary = wb.create_sheet("summary")

    # --------------------------------------------------------
    # HELPERS: reproduce Excel SUMIF/XLOOKUP behaviour in Python
    # --------------------------------------------------------
    def num(value):
        if value is None or value == "":
            return None
        if isinstance(value, bool):
            return float(value)
        if isinstance(value, (int, float)):
            return float(value)
        try:
            return float(str(value).replace(",", "").strip())
        except (ValueError, TypeError):
            return None

    def source_maps(ws):
        sums = {}
        first = {}
        for rr in range(1, ws.max_row + 1):
            key = ws.cell(rr, 1).value
            if key is None:
                continue
            if key not in first:
                first[key] = {
                    "B": num(ws.cell(rr, 2).value),
                    "C": num(ws.cell(rr, 3).value),
                    "D": num(ws.cell(rr, 4).value),
                    "F": num(ws.cell(rr, 6).value),
                    "I": num(ws.cell(rr, 9).value),
                    "J": num(ws.cell(rr, 10).value),
                }
            s = sums.setdefault(key, {"I": 0.0, "J": 0.0})
            iv = num(ws.cell(rr, 9).value)
            jv = num(ws.cell(rr, 10).value)
            if iv is not None:
                s["I"] += iv
            if jv is not None:
                s["J"] += jv
        return first, sums

    maps = {ws.title: source_maps(ws) for ws in sheets20}

    def lookup(ws, key, col):
        return maps[ws.title][0].get(key, {}).get(col)

    def sumif(ws, key, col):
        return maps[ws.title][1].get(key, {}).get(col, 0.0)

    def pct_delta(ws, key, target_col):
        target = lookup(ws, key, target_col)
        base = lookup(ws, key, "B")
        if target is None or base is None or base == 0:
            return None
        return (target - base) / base * 100.0

    def avg_first_four(values):
        valid = [v for v in values if v is not None]
        if not valid:
            return None
        return sum(valid[:4]) / len(valid[:4])

    def fmt2(value):
        if value is None:
            return ""
        return f"{value:.2f}"

    # --------------------------------------------------------
    # HEADINGS
    # --------------------------------------------------------
    summary["A1"] = first_sheet["A1"].value
    summary["B1"] = "Sum I"
    summary["C1"] = "16> C-B / Avg.4"
    summary["D1"] = "16< D-B / Avg.4"
    summary["E1"] = "Avg.4 O2H (1)"
    summary["F1"] = "Avg.4 O2H (2)"
    summary["G1"] = "Sum O2H.10"
    summary["H1"] = "Sum O2L.10"
    summary["I1"] = "Vol.Expand 1"
    summary["J1"] = "Vol.Expand 2"

    for col_num, ws in enumerate(sheets10, start=11):
        summary.cell(1, col_num).value = f"{ws.title} O2H"

    for col_num, ws in enumerate(sheets10, start=21):
        summary.cell(1, col_num).value = f"{ws.title} O2L"

    for col_num, ws in enumerate(sheets10, start=31):
        summary.cell(1, col_num).value = f"{ws.title} C2O"

    summary["AO1"] = '10 "-ve"'
    summary["AP1"] = '10 "+ve"'
    summary["AQ1"] = "%Chg.1"
    summary["AR1"] = "%Chg.2"
    summary["AS1"] = "%Chg.3"
    summary["AT1"] = "%Chg.4"

    # --------------------------------------------------------
    # CALCULATE AND WRITE REAL VALUES -- NOT FORMULAS
    # Covers every cell A2:AR211.
    # --------------------------------------------------------
    for r in range(START_ROW, END_ROW + 1):

        key = first_sheet.cell(r, 1).value
        summary.cell(r, 1).value = key

        # B = SUM I across first 20 sheets.
        b_value = sum(sumif(ws, key, "I") for ws in sheets20)
        summary.cell(r, 2).value = b_value

        cb_values = [pct_delta(ws, key, "C") for ws in sheets20]
        db_values = [pct_delta(ws, key, "D") for ws in sheets20]
        fb_values = [pct_delta(ws, key, "F") for ws in sheets20]

        # Excel LARGE(...,17) over 20 values = 4th smallest.
        cb_valid = sorted(v for v in cb_values if v is not None)
        # Excel SMALL(...,17) over 20 values = 4th largest.
        db_valid = sorted((v for v in db_values if v is not None), reverse=True)

        cb_17th_largest = (
            sorted(cb_valid, reverse=True)[16]
            if len(cb_valid) >= 17 else None
        )
        db_17th_smallest = (
            sorted(v for v in db_values if v is not None)[16]
            if len([v for v in db_values if v is not None]) >= 17 else None
        )

        cb_avg4 = (
            sum(cb_values[:4]) / len([v for v in cb_values[:4] if v is not None])
            if any(v is not None for v in cb_values[:4]) else None
        )
        db_avg4 = (
            sum(db_values[:4]) / len([v for v in db_values[:4] if v is not None])
            if any(v is not None for v in db_values[:4]) else None
        )

        # E/F = rolling four-row averages of the O2H 16> values.
        # (1) = current row through the next 3 rows: A2:A5
        # (2) = next row through the next 4 rows: A3:A6
        # This follows the requested sequence: A2:A5, then A3:A6,
        # then A4:A7, and so on.
        def row_o2h_avg(start_row, end_row):
            values = []
            for source_row in range(start_row, min(end_row, END_ROW) + 1):
                source_key = first_sheet.cell(source_row, 1).value
                if source_key is None:
                    continue
                value = pct_delta(sheets20[0], source_key, "C")
                if value is not None:
                    values.append(value)
            return sum(values) / len(values) if values else None

        o2h_avg4_1 = row_o2h_avg(r, r + 3) if r + 3 <= END_ROW else None
        o2h_avg4_2 = row_o2h_avg(r + 1, r + 4) if r + 4 <= END_ROW else None

        summary.cell(r, 5).value = o2h_avg4_1
        summary.cell(r, 6).value = o2h_avg4_2
        summary.cell(r, 3).value = (
            f"16>{fmt2(cb_17th_largest)}, Avg.4 ({fmt2(cb_avg4)})"
            if cb_17th_largest is not None
            else ""
        )
        summary.cell(r, 4).value = (
            f"16<{fmt2(db_17th_smallest)}, Avg.4 ({fmt2(db_avg4)})"
            if db_17th_smallest is not None
            else ""
        )

        # E/F = sums of the first 10 sheets.
        e_value = sum(v for v in cb_values[:10] if v is not None)
        f_value = sum(v for v in db_values[:10] if v is not None)
        summary.cell(r, 7).value = e_value
        summary.cell(r, 8).value = f_value

        # G/H = volume expansion.
        v1 = sumif(sheets5[0], key, "J")
        v1_avg_values = [sumif(ws, key, "J") for ws in sheets5[1:4]]
        v2 = sumif(sheets5[1], key, "J")
        v2_avg_values = [sumif(ws, key, "J") for ws in sheets5[2:5]]

        v1_avg = sum(v1_avg_values) / 3.0
        v2_avg = sum(v2_avg_values) / 3.0
        summary.cell(r, 9).value = v1 - v1_avg
        summary.cell(r, 10).value = v2 - v2_avg

        # I:R = individual O2H / C-B values.
        for col_num, value in enumerate(cb_values[:10], start=11):
            summary.cell(r, col_num).value = value

        # S:AB = individual O2L / D-B values.
        for col_num, value in enumerate(db_values[:10], start=21):
            summary.cell(r, col_num).value = value

        # AC:AL = individual C2O / F-B values.
        for col_num, value in enumerate(fb_values[:10], start=31):
            summary.cell(r, col_num).value = value

        # AM/AN reproduce the original Excel nested-IF logic:
        # count consecutive negative/positive C2O values from sheet 1
        # through sheet 10. This is NOT a total count if the sequence
        # is interrupted; the original formula stops at the first
        # non-matching value.
        negative_count = 0
        for value in fb_values[:10]:
            if value is not None and value < 0:
                negative_count += 1
            else:
                break

        positive_count = 0
        for value in fb_values[:10]:
            if value is not None and value > 0:
                positive_count += 1
            else:
                break

        # Position numbers show where |O2H| is smaller/larger than |O2L|,
        # matching the original I:R versus S:AB comparison.
        negative_positions = []
        positive_positions = []

        for idx in range(10):
            o2h = cb_values[idx]
            o2l = db_values[idx]

            if o2h is not None and o2l is not None:
                if abs(o2h) < abs(o2l):
                    negative_positions.append(str(idx + 1))
                elif abs(o2h) > abs(o2l):
                    positive_positions.append(str(idx + 1))

        # The + marker compares the Avg.4 values embedded in C and D.
        c_inside = abs(cb_avg4) if cb_avg4 is not None else 999999999
        d_inside = abs(db_avg4) if db_avg4 is not None else 0

        am_plus = "+" if c_inside < d_inside else ""
        an_plus = "+" if c_inside > d_inside else ""

        am_text = str(negative_count) + am_plus
        if negative_positions:
            am_text += "," + ",".join(negative_positions)

        an_text = str(positive_count) + an_plus
        if positive_positions:
            an_text += "," + ",".join(positive_positions)

        summary.cell(r, 41).value = am_text
        summary.cell(r, 42).value = an_text

        # AO:AR = first four sheets' Column I values.
        for col_num, ws in enumerate(sheets[:4], start=43):
            value = sumif(ws, key, "I")
            summary.cell(r, col_num).value = value

    # --------------------------------------------------------
    # FORMATTING
    # --------------------------------------------------------
    summary.freeze_panes = "A2"

    for col_num in range(1, 47):
        summary.column_dimensions[get_column_letter(col_num)].width = 16

    summary.column_dimensions["A"].width = 22
    summary.column_dimensions["B"].width = 18
    summary.column_dimensions["C"].width = 28
    summary.column_dimensions["D"].width = 28
    summary.column_dimensions["E"].width = 18
    summary.column_dimensions["F"].width = 18
    summary.column_dimensions["G"].width = 18
    summary.column_dimensions["H"].width = 18
    summary.column_dimensions["I"].width = 18
    summary.column_dimensions["J"].width = 18
    summary.column_dimensions["AO"].width = 30
    summary.column_dimensions["AP"].width = 30

    for r in range(START_ROW, END_ROW + 1):
        for c in range(5, 9):
            summary.cell(r, c).number_format = "0.00"
        for c in range(9, 11):
            summary.cell(r, c).number_format = "+0.00;-0.00;0.00"
        for c in range(11, 41):
            summary.cell(r, c).number_format = "0.00"
        for c in range(43, 47):
            summary.cell(r, c).number_format = "0.00"

    for row in summary.iter_rows(min_row=1, max_row=END_ROW, min_col=1, max_col=46):
        for cell in row:
            cell.alignment = Alignment(vertical="center")

    summary.auto_filter.ref = f"A1:AT{END_ROW}"

    # Keep workbook calculation settings enabled for any formulas that may
    # exist in the original/source sheets. The Summary sheet itself contains
    # actual cached values, so viewers do not depend on formula recalculation.
    wb.calculation.fullCalcOnLoad = True
    wb.calculation.forceFullCalc = True
    wb.calculation.calcMode = "auto"

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output


# ============================================================
# STREAMLIT UI
# ============================================================

uploaded_file = st.file_uploader(
    "Upload Master Excel File",
    type=["xlsx"]
)


if uploaded_file is not None:

    st.success(
        f"File uploaded: {uploaded_file.name}"
    )

    if st.button(
        "🚀 Generate Summary",
        type="primary"
    ):

        try:

            with st.spinner(
                "Generating summary..."
            ):

                output_file = process_excel(
                    uploaded_file
                )

            st.success(
                "✅ Summary file generated successfully."
            )

            st.download_button(
                label="⬇️ Download summary_output.xlsx",
                data=output_file,
                file_name="summary_output.xlsx",
                mime=(
                    "application/vnd.openxmlformats-officedocument."
                    "spreadsheetml.sheet"
                )
            )

        except Exception as e:

            st.error(
                f"Error: {e}"
            )
