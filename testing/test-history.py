"""ComAp current values report.

Creates an Excel report for Logic link Company Ltd with:
- Styled summary sheet
- All current values from the main controller
- Differential values versus previous run (from SQLite snapshots)
Then sends the Excel file as an email attachment.
"""
import json
import logging
import smtplib
import sqlite3
import ssl
from datetime import datetime, timedelta
from email.message import EmailMessage
from pathlib import Path

from comap import api
from dotenv import dotenv_values
from openpyxl.drawing.image import Image as XLImage
from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, OneCellAnchor
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils.units import pixels_to_EMU

BASE_DIR = Path(__file__).resolve().parents[1]
LOGO_PATH = BASE_DIR / "testing" / "logic-link-logo.png"
secrets = dotenv_values(BASE_DIR / ".env.secret")
shared = dotenv_values(BASE_DIR / ".env.shared")

logging.basicConfig(level=logging.ERROR)


def parse_number(value) -> float:
    text = str(value).replace(",", "").strip()
    return float(text)


def try_parse_number(value):
    try:
        return parse_number(value)
    except (ValueError, TypeError):
        return None


def to_value_map(values: list[dict]) -> dict:
    return {item.get("name"): item.get("value") for item in values if item.get("name")}


def safe_percent(part: float, whole: float) -> float:
    if whole <= 0:
        return 0.0
    return (part / whole) * 100.0


def excel_safe(value):
    if value is None:
        return ""
    if isinstance(value, (str, int, float, bool)):
        return value
    return json.dumps(value, ensure_ascii=True)


def ensure_db(db_path: Path) -> sqlite3.Connection:
    db_path.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(db_path)
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS snapshot_runs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            run_time TEXT NOT NULL
        )
        """
    )
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS snapshot_values (
            run_id INTEGER NOT NULL,
            name TEXT NOT NULL,
            value_text TEXT,
            value_num REAL,
            unit TEXT,
            guid TEXT,
            FOREIGN KEY(run_id) REFERENCES snapshot_runs(id)
        )
        """
    )
    conn.commit()
    return conn


def get_latest_snapshot(conn: sqlite3.Connection):
    row = conn.execute(
        "SELECT id, run_time FROM snapshot_runs ORDER BY id DESC LIMIT 1"
    ).fetchone()
    return row


def parse_run_time(text: str):
    for fmt in ("%Y-%m-%d %H:%M:%S", "%m/%d/%Y %H:%M"):
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            continue
    return None


def get_snapshot_near_target(conn: sqlite3.Connection, target_dt: datetime):
    rows = conn.execute(
        "SELECT id, run_time FROM snapshot_runs ORDER BY id ASC"
    ).fetchall()
    best_row = None
    best_delta = None
    for row in rows:
        parsed = parse_run_time(row[1])
        if parsed is None:
            continue
        delta = abs((parsed - target_dt).total_seconds())
        if best_delta is None or delta < best_delta:
            best_delta = delta
            best_row = row
    return best_row


def get_snapshot_value_map(conn: sqlite3.Connection, run_id: int) -> dict:
    rows = conn.execute(
        """
        SELECT name, value_text, value_num, unit, guid
        FROM snapshot_values
        WHERE run_id = ?
        """,
        (run_id,),
    ).fetchall()

    result = {}
    for name, value_text, value_num, unit, guid in rows:
        result[name] = {
            "value_text": value_text,
            "value_num": value_num,
            "unit": unit,
            "guid": guid,
        }
    return result


def save_snapshot(conn: sqlite3.Connection, run_time: str, all_values: list[dict]) -> None:
    cursor = conn.cursor()
    cursor.execute("INSERT INTO snapshot_runs(run_time) VALUES (?)", (run_time,))
    run_id = cursor.lastrowid

    for item in all_values:
        name = str(item.get("name", ""))
        value_raw = item.get("value")
        value_text = excel_safe(value_raw)
        value_num = try_parse_number(value_raw)
        unit = excel_safe(item.get("unit", ""))
        guid = excel_safe(item.get("guid", ""))

        cursor.execute(
            """
            INSERT INTO snapshot_values(run_id, name, value_text, value_num, unit, guid)
            VALUES (?, ?, ?, ?, ?, ?)
            """,
            (run_id, name, value_text, value_num, unit, guid),
        )

    conn.commit()


def build_differentials(current_values: list[dict], previous_map: dict) -> list[dict]:
    rows = []
    for item in current_values:
        name = excel_safe(item.get("name", ""))
        current_raw = item.get("value")
        current_text = excel_safe(current_raw)
        current_num = try_parse_number(current_raw)
        unit = excel_safe(item.get("unit", ""))
        guid = excel_safe(item.get("guid", ""))

        prev = previous_map.get(name)
        previous_text = prev["value_text"] if prev else ""
        previous_num = prev["value_num"] if prev else None

        delta = ""
        if current_num is not None and previous_num is not None:
            delta = current_num - previous_num

        rows.append(
            {
                "name": name,
                "current": current_text,
                "previous": previous_text,
                "delta": delta,
                "unit": unit,
                "guid": guid,
            }
        )

    return rows


def build_excel_report(
    path: Path,
    company_name: str,
    report_title: str,
    generated_at: str,
    previous_run_time,
    report_rows: list[tuple[str, object, str]],
    warnings: list[str],
) -> None:
    workbook = Workbook()

    # Sheet 1: Styled summary
    summary_sheet = workbook.active
    summary_sheet.title = "Summary"
    summary_sheet.sheet_view.showGridLines = False

    summary_sheet.column_dimensions["A"].width = 12
    summary_sheet.column_dimensions["B"].width = 38
    summary_sheet.column_dimensions["C"].width = 36
    summary_sheet.column_dimensions["D"].width = 18
    summary_sheet.row_dimensions[1].height = 24
    summary_sheet.row_dimensions[2].height = 24
    summary_sheet.row_dimensions[3].height = 24

    thin = Side(style="thin", color="000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    header_fill = PatternFill(fill_type="solid", fgColor="A6A6A6")
    stripe_fill = PatternFill(fill_type="solid", fgColor="D9E1F2")
    white_fill = PatternFill(fill_type="solid", fgColor="FFFFFF")

    summary_sheet.merge_cells("C2:D3")
    summary_sheet["C2"] = f"{company_name.upper()}"
    summary_sheet["C2"].font = Font(bold=True, size=18, color="7F6000")
    summary_sheet["C2"].alignment = Alignment(horizontal="center", vertical="center")
    if LOGO_PATH.exists():
        try:
            logo = XLImage(str(LOGO_PATH))
            logo.width = 430
            logo.height = 88

            # Center the logo across columns A:D with a small top margin.
            col_widths_px = [
                int(summary_sheet.column_dimensions[col].width * 7) for col in ("A", "B", "C", "D")
            ]
            total_width_px = sum(col_widths_px)
            start_x_px = max((total_width_px - int(logo.width)) // 2, 0)

            col_idx = 0
            x_in_col_px = start_x_px
            for i, width_px in enumerate(col_widths_px):
                if x_in_col_px < width_px:
                    col_idx = i
                    break
                x_in_col_px -= width_px

            logo.anchor = OneCellAnchor(
                _from=AnchorMarker(
                    col=col_idx,
                    colOff=pixels_to_EMU(x_in_col_px),
                    row=0,
                    rowOff=pixels_to_EMU(4),
                ),
                ext=None,
            )
            logo.anchor.ext.cx = pixels_to_EMU(int(logo.width))
            logo.anchor.ext.cy = pixels_to_EMU(int(logo.height))
            summary_sheet.add_image(logo)
            summary_sheet["C2"] = ""
        except Exception:
            # Keep report generation working even if image loading fails.
            pass

    summary_sheet["B5"] = "Generated At"
    summary_sheet["B5"].font = Font(size=12)
    summary_sheet["D5"] = generated_at
    summary_sheet["D5"].font = Font(size=12)
    summary_sheet["D5"].alignment = Alignment(horizontal="center")

    summary_sheet.merge_cells("B6:D6")
    summary_sheet["B6"] = report_title
    summary_sheet["B6"].font = Font(size=14, bold=True)
    summary_sheet["B6"].alignment = Alignment(horizontal="center", vertical="center")

    header_row = 8
    headers = ["Item", "Description", "Values", "Unit"]
    for col, value in enumerate(headers, start=1):
        cell = summary_sheet.cell(row=header_row, column=col, value=value)
        cell.font = Font(bold=True, color="FFFFFF", size=12)
        cell.fill = header_fill
        cell.border = border
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for idx, (desc, value, unit) in enumerate(report_rows, start=1):
        row = header_row + idx
        fill = stripe_fill if idx % 2 == 1 else white_fill

        summary_sheet.cell(row=row, column=1, value=idx)
        summary_sheet.cell(row=row, column=2, value=desc)
        summary_sheet.cell(row=row, column=3, value=excel_safe(value))
        summary_sheet.cell(row=row, column=4, value=unit)

        for col in range(1, 5):
            cell = summary_sheet.cell(row=row, column=col)
            cell.fill = fill
            cell.border = border
            if col == 1:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif col == 3:
                cell.alignment = Alignment(horizontal="right", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center")

    if warnings:
        warn_row = header_row + len(report_rows) + 2
        summary_sheet.cell(row=warn_row, column=1, value="Warnings").font = Font(bold=True, color="C00000")
        for i, warning in enumerate(warnings, start=1):
            summary_sheet.cell(row=warn_row + i, column=1, value=f"- {warning}")

    summary_sheet.cell(row=4, column=1, value=f"Compared To: {previous_run_time if previous_run_time else 'No previous snapshot'}")
    summary_sheet.merge_cells("A4:D4")

    workbook.save(path)


def send_email_with_attachment(subject: str, body: str, attachment_path: Path) -> None:
    message = EmailMessage()
    message["From"] = secrets["MAIL_FROM"]
    message["To"] = secrets["MAIL_TO"]
    message["Subject"] = subject
    message.set_content(body)

    with open(attachment_path, "rb") as file:
        data = file.read()
    message.add_attachment(
        data,
        maintype="application",
        subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=attachment_path.name,
    )

    host = secrets["MAILTRAP_HOST"]
    port = int(secrets["MAILTRAP_PORT"])
    username = secrets["MAILTRAP_USERNAME"]
    password = secrets["MAILTRAP_PASSWORD"]

    if port == 465:
        context = ssl.create_default_context()
        with smtplib.SMTP_SSL(host, port, context=context) as server:
            server.login(username, password)
            server.send_message(message)
        return

    with smtplib.SMTP(host, port) as server:
        server.ehlo()
        server.starttls(context=ssl.create_default_context())
        server.ehlo()
        server.login(username, password)
        server.send_message(message)


# Use the ComAp Cloud Identity API to get the Bearer token
identity = api.Identity(secrets["COMAP_KEY"])
token = identity.authenticate(secrets["CLIENT_ID"], secrets["SECRET"])

if token is not None:
    wsv = api.WSV(secrets["LOGIN_ID"], secrets["COMAP_KEY"], token["access_token"])

    company_name = shared.get("COMPANY_NAME", "Logic link Company Ltd")
    report_title = shared.get("REPORT_TITLE", "Pun Hlaing Energy consumption per day")

    filename = shared.get("FILENAME", "Energy_Report.xlsx")
    report_path = Path(filename)
    if not report_path.is_absolute():
        report_path = BASE_DIR / report_path

    # Use a fixed absolute path to avoid creating a new DB per run
    db_path = Path("/opt/comap-report/data/comap_snapshots.db")

    now_dt = datetime.now()
    generated_at = now_dt.strftime("%m/%d/%Y %H:%M")
    run_time_storage = now_dt.strftime("%Y-%m-%d %H:%M:%S")
    warnings = []

    main_values = wsv.values(shared["GENSET_ID_Main"])
    main_value_map = to_value_map(main_values)

    solar_mwh_raw = main_value_map.get("Solar MWh")
    m_kwh_i_raw = main_value_map.get("M kWh I")
    f1_mwh_raw = main_value_map.get("F-1 MWH")
    f2_mwh_raw = main_value_map.get("F-2 MWH")
    f3_mwh_raw = main_value_map.get("F-3 MWh")

    try:
        solar_mwh = parse_number(solar_mwh_raw) if solar_mwh_raw is not None else 0.0
    except ValueError:
        solar_mwh = 0.0
        warnings.append("Solar MWh is not numeric; treated as 0")

    try:
        m_kwh_i = parse_number(m_kwh_i_raw) if m_kwh_i_raw is not None else 0.0
    except ValueError:
        m_kwh_i = 0.0
        warnings.append("M kWh I is not numeric; treated as 0")

    if solar_mwh_raw is None:
        warnings.append("Solar MWh not found; treated as 0")
    if m_kwh_i_raw is None:
        warnings.append("M kWh I not found; treated as 0")
    if f1_mwh_raw is None:
        warnings.append("F-1 MWh not found; treated as 0")
    if f2_mwh_raw is None:
        warnings.append("F-2 MWh not found; treated as 0")
    if f3_mwh_raw is None:
        warnings.append("F-3 MWh not found; treated as 0")

    try:
        f1_mwh = parse_number(f1_mwh_raw) if f1_mwh_raw is not None else 0.0
    except ValueError:
        f1_mwh = 0.0
        warnings.append("F-1 MWh is not numeric; treated as 0")

    try:
        f2_mwh = parse_number(f2_mwh_raw) if f2_mwh_raw is not None else 0.0
    except ValueError:
        f2_mwh = 0.0
        warnings.append("F-2 MWh is not numeric; treated as 0")

    try:
        f3_mwh = parse_number(f3_mwh_raw) if f3_mwh_raw is not None else 0.0
    except ValueError:
        f3_mwh = 0.0
        warnings.append("F-3 MWh is not numeric; treated as 0")

    solar_kwh = solar_mwh * 1000.0

    genset_total_kwh = 0.0
    for key in ["GENSET_ID_1", "GENSET_ID_2", "GENSET_ID_3", "GENSET_ID_4"]:
        raw_values = wsv.values(shared[key])
        genset_value_map = to_value_map(raw_values)
        raw_value = genset_value_map.get("Genset kWh")

        if raw_value is None:
            warnings.append(f"{key} Genset kWh not found; treated as 0")
            continue
        try:
            genset_total_kwh += parse_number(raw_value)
        except ValueError:
            warnings.append(f"{key} Genset kWh is not numeric; treated as 0")

    sum_total_kwh = solar_kwh + m_kwh_i + genset_total_kwh

    conn = ensure_db(db_path)
    target_dt = now_dt - timedelta(days=1)
    baseline = get_snapshot_near_target(conn, target_dt)
    previous_run_time = baseline[1] if baseline else None
    previous_map = get_snapshot_value_map(conn, baseline[0]) if baseline else {}

    previous_solar_mwh = previous_map.get("Solar MWh", {}).get("value_num")
    previous_grid_kwh = previous_map.get("M kWh I", {}).get("value_num")
    previous_f1_mwh = previous_map.get("F-1 MWh", {}).get("value_num")
    previous_f2_mwh = previous_map.get("F-2 MWh", {}).get("value_num")
    previous_f3_mwh = previous_map.get("F-3 MWh", {}).get("value_num")
    previous_genset_total_kwh = previous_map.get("Genset Total kWh (Calculated)", {}).get("value_num")

    if (
        previous_solar_mwh is None
        or previous_grid_kwh is None
        or previous_f1_mwh is None
        or previous_f2_mwh is None
        or previous_f3_mwh is None
        or previous_genset_total_kwh is None
    ):
        warnings.append("24-hour baseline not found. Usage shown as current values on first run.")
        usage_solar_mwh = solar_mwh
        usage_grid_kwh = m_kwh_i
        usage_f1_mwh = f1_mwh
        usage_f2_mwh = f2_mwh
        usage_f3_mwh = f3_mwh
        usage_genset_kwh = genset_total_kwh
    else:
        usage_solar_mwh = solar_mwh - previous_solar_mwh
        usage_grid_kwh = m_kwh_i - previous_grid_kwh
        usage_f1_mwh = f1_mwh - previous_f1_mwh
        usage_f2_mwh = f2_mwh - previous_f2_mwh
        usage_f3_mwh = f3_mwh - previous_f3_mwh
        usage_genset_kwh = genset_total_kwh - previous_genset_total_kwh

    usage_sum_total_kwh = (usage_solar_mwh * 1000.0) + usage_grid_kwh + usage_genset_kwh
    usage_solar_pct = safe_percent(usage_solar_mwh * 1000.0, usage_sum_total_kwh)
    usage_grid_pct = safe_percent(usage_grid_kwh, usage_sum_total_kwh)
    usage_genset_pct = safe_percent(usage_genset_kwh, usage_sum_total_kwh)

    report_rows = [
        ("Solar Energy consumption", round(usage_solar_mwh, 3), "MWh"),
        ("Grid Energy consumption", round(usage_grid_kwh, 3), "kWh"),
        ("Genset Total kWh", round(usage_genset_kwh, 3), "kWh"),
        ("Sum of Total kWh", round(usage_sum_total_kwh, 3), "kWh"),
        ("Avg% Solar kWh", f"{usage_solar_pct:.2f}%", "%"),
        ("Avg% Grid kWh", f"{usage_grid_pct:.2f}%", "%"),
        ("Avg% Genset Total", f"{usage_genset_pct:.2f}%", "%"),
        ("F-1 MWh", round(usage_f1_mwh, 3), "MWh"),
        ("F-2 MWh", round(usage_f2_mwh, 3), "MWh"),
        ("F-3 MWh", round(usage_f3_mwh, 3), "MWh"),
    ]

    metrics_for_snapshot = list(main_values)
    metrics_for_snapshot.extend(
        [
            {"name": "Genset Total kWh (Calculated)", "value": genset_total_kwh, "unit": "kWh", "guid": ""},
            {"name": "Sum of Total kWh (Calculated)", "value": sum_total_kwh, "unit": "kWh", "guid": ""},
        ]
    )
    save_snapshot(conn, run_time_storage, metrics_for_snapshot)
    conn.close()

    build_excel_report(
        report_path,
        company_name,
        report_title,
        generated_at,
        previous_run_time,
        report_rows,
        warnings,
    )

    print("Excel report generated:", str(report_path))
    email_body = (
        "Dear Sir,\n"
        "Please find attached the daily energy consumption report generated via the "
        "ComAp WebSupervisor API through the LPI local server. The report covers one "
        "main supply, one solar source, three feeders, and four gensets.\n"
        "Best regards,"
    )
    send_email_with_attachment("Daily Energy Consumption Report", email_body, report_path)
    print("Email sent with Excel attachment.")
