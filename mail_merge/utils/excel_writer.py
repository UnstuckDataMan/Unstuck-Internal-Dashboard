"""
Excel output writer for merge results and send schedules.

Produces professionally formatted workbooks with:
  - Coloured, grouped headers
  - Dropdown data validation on tracking fields
  - Conditional formatting for status columns
  - Auto-filter and frozen header row
"""
import calendar as cal_mod
import re
import zipfile
from io import BytesIO
from typing import List, Dict

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import Rule
from openpyxl.styles.differential import DifferentialStyle


# ──────────────────────────────────────────────────────────────────────────── #
# Colour tokens
# ──────────────────────────────────────────────────────────────────────────── #
C = {
    # Section header fills
    'hdr_prospect':  '1E3A5F',   # dark navy   – prospect cols
    'hdr_sender':    '1565C0',   # blue        – sender / recipient
    'hdr_template':  '0D47A1',   # darker blue – generated copy
    'hdr_chaser':    '283593',   # indigo      – chaser
    'hdr_tracking':  '1B5E20',   # dark green  – tracking
    'hdr_schedule':  '1E3A5F',   # navy        – schedule sheet
    # Row fills
    'row_even':  'EBF5FB',
    'row_odd':   'FFFFFF',
    # Status colours (conditional formatting)
    'green_bg':  'C8E6C9',
    'red_bg':    'FFCDD2',
    'amber_bg':  'FFF9C4',
    'blue_bg':   'E3F2FD',
    'purple_bg': 'F3E5F5',
    'orange_bg': 'FFE0B2',
}

THIN_BORDER_COLOR = 'D0D7DE'


def _thin_border():
    s = Side(style='thin', color=THIN_BORDER_COLOR)
    return Border(left=s, right=s, top=s, bottom=s)


def _hdr_cell(cell, fill_hex: str, font_size: int = 9):
    cell.font = Font(name='Calibri', bold=True, color='FFFFFF', size=font_size)
    cell.fill = PatternFill('solid', fgColor=fill_hex)
    cell.alignment = Alignment(horizontal='center', vertical='center',
                                wrap_text=True)
    cell.border = _thin_border()


def _data_cell(cell, fill_hex: str, wrap: bool = False, bold: bool = False):
    cell.font = Font(name='Calibri', size=9, bold=bold)
    cell.fill = PatternFill('solid', fgColor=fill_hex)
    cell.alignment = Alignment(vertical='top', wrap_text=wrap)
    cell.border = _thin_border()


def _add_dv(ws, formula: str, col_letter: str, max_row: int):
    dv = DataValidation(type='list', formula1=formula, showDropDown=False)
    dv.sqref = f"{col_letter}2:{col_letter}{max_row}"
    ws.add_data_validation(dv)


def _cf_rule(formula: str, fill_hex: str) -> Rule:
    """Build a formula-based CF rule with a properly registered dxf fill.

    FormulaRule/CellIsRule both fail to reliably register the PatternFill in
    the workbook's dxf styles block. Using Rule + DifferentialStyle directly
    is the only approach that works across all openpyxl versions.
    Setting both fgColor and bgColor is required by Excel for solid fills
    in differential formatting contexts.
    """
    fill = PatternFill(patternType='solid', fgColor=fill_hex, bgColor=fill_hex)
    dxf  = DifferentialStyle(fill=fill)
    return Rule(type='expression', dxf=dxf, formula=[formula])


def _add_cf_equal(ws, data_range: str, match_value: str, fill_hex: str):
    start_cell = data_range.split(':')[0]
    col = ''.join(c for c in start_cell if c.isalpha())
    row = ''.join(c for c in start_cell if c.isdigit())
    ws.conditional_formatting.add(
        data_range,
        _cf_rule(f'${col}{row}="{match_value}"', fill_hex),
    )


# ──────────────────────────────────────────────────────────────────────────── #
# EXCEL 365 NATIVE CHECKBOX INJECTION
# ──────────────────────────────────────────────────────────────────────────── #

def _patch_content_types(data: bytes) -> bytes:
    """Add richData and metadata Override entries to [Content_Types].xml."""
    text = data.decode('utf-8')
    additions = (
        '<Override PartName="/xl/metadata.xml"'
        ' ContentType="application/vnd.ms-excel.sheetMetadata+xml"/>'
        '<Override PartName="/xl/richData/rdRichValueTypes.xml"'
        ' ContentType="application/vnd.ms-office.spreadsheetml.rdRichValueTypes+xml"/>'
        '<Override PartName="/xl/richData/rdrichvalue.xml"'
        ' ContentType="application/vnd.ms-office.spreadsheetml.rdRichValue+xml"/>'
    )
    return text.replace('</Types>', additions + '</Types>').encode('utf-8')


def _patch_workbook_rels(data: bytes) -> bytes:
    """Add relationships for richData and metadata files to workbook.xml.rels."""
    text = data.decode('utf-8')
    additions = (
        '<Relationship Id="rIdMeta"'
        ' Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata"'
        ' Target="metadata.xml"/>'
        '<Relationship Id="rIdRVT"'
        ' Type="http://schemas.microsoft.com/office/2022/10/relationships/richValueRelTypes"'
        ' Target="richData/rdRichValueTypes.xml"/>'
        '<Relationship Id="rIdRVD"'
        ' Type="http://schemas.microsoft.com/office/2022/10/relationships/rdRichValue"'
        ' Target="richData/rdrichvalue.xml"/>'
    )
    return text.replace('</Relationships>', additions + '</Relationships>').encode('utf-8')


def _patch_ws_vm(data: bytes, cell_refs: List[str]) -> bytes:
    """Stamp vm="N" on each checkbox cell's opening tag in the worksheet XML.

    Cells with a value are never self-closing, so matching [^>]* up to > is
    safe.  The vm index is 1-based (1 = first <bk> in metadata valueMetadata).
    """
    text = data.decode('utf-8')
    for i, ref in enumerate(cell_refs):
        vm_idx = i + 1
        text = re.sub(
            rf'(<c\s+r="{re.escape(ref)}"[^>]*)>',
            rf'\g<1> vm="{vm_idx}">',
            text,
            count=1,
        )
    return text.encode('utf-8')


def _inject_excel_checkboxes(path: str, cell_refs: List[str]) -> None:
    """Post-process a saved xlsx to give column-A cells native Excel 365 checkboxes.

    Each cell in *cell_refs* must already hold a Python False value so that
    openpyxl writes  t="b" <v>0</v>.  This function adds the three XML parts
    (rdRichValueTypes, rdrichvalue, metadata) that Excel 365 needs to render
    the cell as an interactive checkbox widget, and stamps each cell tag with
    the matching vm="N" attribute.  The file is rewritten in-place.

    If anything goes wrong the original file is left on disk unchanged.
    """
    if not cell_refs:
        return

    n = len(cell_refs)

    # ── rdRichValueTypes.xml — defines the $Checkbox rich-value type ───────
    rd_types_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<rvTypesInfo'
        ' xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata2"'
        ' xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"'
        ' mc:Ignorable="x"'
        ' xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        '<global><keyFlags>'
        '<key name="ExcelValueFlags"><flag name="ShowCard" value="1"/></key>'
        '</keyFlags></global>'
        '<types>'
        '<type name="$Checkbox">'
        '<schema type="mixed"/>'
        '<keyFlags>'
        '<key name="ExcelValueFlags"><flag name="HideCard" value="1"/></key>'
        '</keyFlags>'
        '</type>'
        '</types>'
        '</rvTypesInfo>'
    )

    # ── rdrichvalue.xml — one rv entry per cell (all unchecked = 0) ────────
    rv_entries = ''.join(
        '<rv t="0"><v>0</v><v>$Checkbox</v></rv>' for _ in range(n)
    )
    rd_richvalue_xml = (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<rvData'
        f' xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata"'
        f' count="{n}">{rv_entries}</rvData>'
    )

    # ── metadata.xml — maps each cell's vm index to its rv entry ──────────
    # vm is 1-based; bk[0] → cell vm="1", bk[1] → cell vm="2", …
    # rc v="I" is 0-based index into rdrichvalue (each cell has its own entry)
    bk_entries = ''.join(f'<bk><rc t="0" v="{i}"/></bk>' for i in range(n))
    metadata_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<metadata'
        ' xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        '<metadataTypes count="1">'
        '<metadataType name="XLMDEC" minSupportedVersion="120000"'
        ' copy="1" pasteAll="1" pasteValues="1" merge="1" splitFirst="1"'
        ' rowColShift="1" clearFormats="1" clearValues="1" adjust="1" cellMeta="1"/>'
        '</metadataTypes>'
        f'<valueMetadata count="{n}">{bk_entries}</valueMetadata>'
        '</metadata>'
    )

    try:
        with open(path, 'rb') as fh:
            original = fh.read()

        buf = BytesIO()
        with zipfile.ZipFile(BytesIO(original), 'r') as zin, \
             zipfile.ZipFile(buf, 'w', compression=zipfile.ZIP_DEFLATED) as zout:

            for item in zin.infolist():
                raw = zin.read(item.filename)
                if item.filename == '[Content_Types].xml':
                    raw = _patch_content_types(raw)
                elif item.filename == 'xl/_rels/workbook.xml.rels':
                    raw = _patch_workbook_rels(raw)
                elif item.filename == 'xl/worksheets/sheet1.xml':
                    raw = _patch_ws_vm(raw, cell_refs)
                zout.writestr(item, raw)

            # Inject new XML parts
            zout.writestr('xl/richData/rdRichValueTypes.xml',
                          rd_types_xml.encode('utf-8'))
            zout.writestr('xl/richData/rdrichvalue.xml',
                          rd_richvalue_xml.encode('utf-8'))
            zout.writestr('xl/metadata.xml',
                          metadata_xml.encode('utf-8'))

        with open(path, 'wb') as fh:
            fh.write(buf.getvalue())

    except Exception:
        # Injection failed — original file is still valid (cells show TRUE/FALSE)
        pass


# ──────────────────────────────────────────────────────────────────────────── #
# MERGE OUTPUT
# ──────────────────────────────────────────────────────────────────────────── #

def write_merge_output(
    output_path: str,
    prospect_headers: List[str],
    merged_rows: List[Dict],
    has_chaser: bool,
    email_column: str = '',
    has_schedule: bool = False,
):
    wb = Workbook()
    ws = wb.active
    ws.title = 'Outreach List'
    ws.freeze_panes = 'A2'

    # ── Column layout ──────────────────────────────────────────────────────
    # Section 1: sent checkbox — always first so the SDR can tick rows off easily
    sec_status = ['Send Status']
    # Section 2: schedule dates
    sec_schedule = ['Send Date', 'Send Time'] if has_schedule else []
    # Section 3: sender / recipient
    sec_routing = ['Sender Account', 'Recipient Email']
    # Section 4: generated copy
    sec_template = ['Subject Line', 'Email Body']
    if has_chaser:
        sec_template += ['Chaser Subject', 'Chaser Body']
    sec_template += ['A/B Variant']
    # Section 5: tracking
    sec_tracking = ['Response', 'Lead Status', 'Notes']

    all_cols = sec_status + sec_schedule + sec_routing + sec_template + sec_tracking

    # Map header → colour token
    color_map: Dict[str, str] = {}
    for h in sec_status:
        color_map[h] = C['hdr_tracking']
    for h in sec_schedule:
        color_map[h] = C['hdr_tracking']
    for h in sec_routing:
        color_map[h] = C['hdr_sender']
    for h in ['Subject Line', 'Email Body', 'A/B Variant']:
        color_map[h] = C['hdr_template']
    for h in ['Chaser Subject', 'Chaser Body']:
        color_map[h] = C['hdr_chaser']
    for h in sec_tracking:
        color_map[h] = C['hdr_tracking']

    # Column widths
    col_widths = {
        'Send Status': 10,
        'Sender Account': 30, 'Recipient Email': 34,
        'Subject Line': 44,   'Email Body': 64,
        'Chaser Subject': 44, 'Chaser Body': 64,
        'A/B Variant': 12,
        'Response': 22,
        'Lead Status': 18,    'Send Date': 14,
        'Send Time': 12,      'Notes': 32,
    }

    # ── Write header row ───────────────────────────────────────────────────
    for ci, header in enumerate(all_cols, 1):
        cell = ws.cell(row=1, column=ci, value=header)
        _hdr_cell(cell, color_map.get(header, C['hdr_prospect']))
        ws.column_dimensions[get_column_letter(ci)].width = col_widths.get(header, 20)

    ws.row_dimensions[1].height = 32

    # ── Data validation ────────────────────────────────────────────────────
    col_map = {h: i + 1 for i, h in enumerate(all_cols)}
    total_data_rows = len(merged_rows)
    last_row = total_data_rows + 1

    def col_let(name):
        return get_column_letter(col_map[name])

    _add_dv(ws, '"No Response,Positive Reply,Negative Reply,Unsubscribed,Auto-Reply"',
            col_let('Response'), last_row)
    _add_dv(ws, '"Not a Lead,MQL,SQL,Meeting Booked,Closed Won,Closed Lost"',
            col_let('Lead Status'), last_row)

    # ── Write data rows ────────────────────────────────────────────────────
    prev_sender = None
    for ri, row_data in enumerate(merged_rows, 2):
        current_sender = row_data.get('__sender_account__', '')
        is_first_of_sender = (current_sender != prev_sender)
        prev_sender = current_sender

        # Yellow stripe on the first row of each new sender block so the SDR
        # can see at a glance when the sending account changes.
        if is_first_of_sender:
            fill_hex = 'FFFDE7'
        else:
            fill_hex = 'FFFFFF'

        row_values = {
            'Send Status': False,
            'Sender Account': row_data.get('__sender_account__', ''),
            'Recipient Email': row_data.get('__recipient_email__', ''),
            'Subject Line': row_data.get('__subject_line__', ''),
            'Email Body': row_data.get('__email_body__', ''),
            'Chaser Subject': row_data.get('__chaser_subject__', ''),
            'Chaser Body': row_data.get('__chaser_body__', ''),
            'A/B Variant': row_data.get('__template_variant__', ''),
            'Response': 'No Response',
            'Lead Status': 'Not a Lead',
            'Send Date': row_data.get('__send_date__', ''),
            'Send Time': row_data.get('__send_time__', ''),
            'Notes': '',
        }

        wrap_cols = {'Email Body', 'Chaser Body', 'Notes'}

        for ci, header in enumerate(all_cols, 1):
            cell = ws.cell(row=ri, column=ci, value=row_values.get(header, ''))
            if header == 'Send Status':
                # Native checkbox cell — keep it centred, standard font size
                cell.font = Font(name='Calibri', size=9)
                cell.fill = PatternFill('solid', fgColor=fill_hex)
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.border = _thin_border()
            else:
                _data_cell(cell, fill_hex, wrap=(header in wrap_cols))

        ws.row_dimensions[ri].height = 72 if total_data_rows <= 200 else 36

    # ── Conditional formatting ─────────────────────────────────────────────
    def cf_range(name):
        return f"{col_let(name)}2:{col_let(name)}{last_row}"

    # Whole-row pastel green when checkbox is ticked — added first so cell-level
    # rules (Response, Lead Status) take priority over the row highlight.
    ws.conditional_formatting.add(
        f"A2:{get_column_letter(len(all_cols))}{last_row}",
        _cf_rule('$A2=TRUE', 'E8F5E9'),
    )

    _add_cf_equal(ws, cf_range('Response'), 'Positive Reply', C['green_bg'])
    _add_cf_equal(ws, cf_range('Response'), 'Negative Reply', C['red_bg'])
    _add_cf_equal(ws, cf_range('Response'), 'Unsubscribed',   C['amber_bg'])

    _add_cf_equal(ws, cf_range('Lead Status'), 'MQL',          C['amber_bg'])
    _add_cf_equal(ws, cf_range('Lead Status'), 'SQL',          C['orange_bg'])
    _add_cf_equal(ws, cf_range('Lead Status'), 'Meeting Booked', C['purple_bg'])
    _add_cf_equal(ws, cf_range('Lead Status'), 'Closed Won',   C['green_bg'])
    _add_cf_equal(ws, cf_range('Lead Status'), 'Closed Lost',  C['red_bg'])

    # ── Auto-filter ────────────────────────────────────────────────────────
    ws.auto_filter.ref = f"A1:{get_column_letter(len(all_cols))}1"

    # ── Summary sheet ──────────────────────────────────────────────────────
    ws2 = wb.create_sheet('Summary')
    _write_merge_summary(ws2, merged_rows, has_chaser)

    wb.save(output_path)

    # Inject Excel 365 native checkboxes into every Send Status cell (column A)
    _inject_excel_checkboxes(
        output_path,
        [f'A{r}' for r in range(2, last_row + 1)],
    )


def _write_merge_summary(ws, merged_rows: List[Dict], has_chaser: bool):
    ws.column_dimensions['A'].width = 34
    ws.column_dimensions['B'].width = 22

    t = ws.cell(row=1, column=1, value='Campaign Merge Summary')
    t.font = Font(name='Calibri', bold=True, size=14, color='1E3A5F')

    senders = {r.get('__sender_account__', '') for r in merged_rows}
    variants = {r.get('__template_variant__', '') for r in merged_rows}

    stats = [
        ('Total Prospects Merged',  len(merged_rows)),
        ('Unique Sender Accounts',  len(senders)),
        ('Template Variants',        len(variants)),
        ('Chaser Templates Included', 'Yes' if has_chaser else 'No'),
    ]

    for ri, (label, value) in enumerate(stats, 3):
        lc = ws.cell(row=ri, column=1, value=label)
        lc.font = Font(name='Calibri', bold=True, size=10)
        lc.fill = PatternFill('solid', fgColor='EBF5FB')
        vc = ws.cell(row=ri, column=2, value=value)
        vc.font = Font(name='Calibri', size=10)

    # Sender breakdown
    sender_counts: Dict[str, int] = {}
    for r in merged_rows:
        s = r.get('__sender_account__', 'Unknown')
        sender_counts[s] = sender_counts.get(s, 0) + 1

    row = len(stats) + 5
    ws.cell(row=row, column=1, value='Sender Distribution').font = Font(
        bold=True, size=11, color='1E3A5F')
    row += 1

    for sender, count in sorted(sender_counts.items()):
        ws.cell(row=row, column=1, value=sender)
        ws.cell(row=row, column=2, value=count)
        row += 1


# ──────────────────────────────────────────────────────────────────────────── #
# SCHEDULE OUTPUT
# ──────────────────────────────────────────────────────────────────────────── #

def write_schedule_output(
    output_path: str,
    schedule: List[Dict],
    year: int,
    month: int,
    prospect_count: int,
):
    wb = Workbook()
    ws = wb.active
    ws.title = 'Send Schedule'
    ws.freeze_panes = 'A2'

    sch_cols = [
        '#', 'Date', 'Day', 'Send Time',
        'Sender Account', 'Prospect ID', 'Status', 'Notes',
    ]
    col_widths = {
        '#': 6, 'Date': 14, 'Day': 12, 'Send Time': 12,
        'Sender Account': 34,
        'Prospect ID': 13, 'Status': 14, 'Notes': 32,
    }

    # Headers
    for ci, h in enumerate(sch_cols, 1):
        cell = ws.cell(row=1, column=ci, value=h)
        _hdr_cell(cell, C['hdr_schedule'])
        ws.column_dimensions[get_column_letter(ci)].width = col_widths.get(h, 15)
    ws.row_dimensions[1].height = 30

    # Data validation – Status column (col 7)
    total_rows = len(schedule)
    _add_dv(ws, '"Scheduled,Sent,Failed,Skipped"', 'G', total_rows + 1)

    # Write rows, alternating fill colour per date block
    date_fills: Dict[str, str] = {}
    fill_options = ['EBF5FB', 'FDFEFE']
    fill_toggle = [0]

    def get_date_fill(d: str) -> str:
        if d not in date_fills:
            date_fills[d] = fill_options[fill_toggle[0] % 2]
            fill_toggle[0] += 1
        return date_fills[d]

    for ri, send in enumerate(schedule, 2):
        fill_hex = get_date_fill(send['date'])
        vals = {
            '#': ri - 1,
            'Date': send['date'],
            'Day': send['day_of_week'],
            'Send Time': send['send_time'],
            'Sender Account': send['sender'],
            'Prospect ID': send['prospect_id'],
            'Status': 'Scheduled',
            'Notes': '',
        }
        for ci, h in enumerate(sch_cols, 1):
            cell = ws.cell(row=ri, column=ci, value=vals.get(h, ''))
            _data_cell(cell, fill_hex)
        ws.row_dimensions[ri].height = 16

    # Conditional formatting on Status column
    status_range = f"G2:G{total_rows + 1}"
    _add_cf_equal(ws, status_range, 'Sent',      C['green_bg'])
    _add_cf_equal(ws, status_range, 'Failed',    C['red_bg'])
    _add_cf_equal(ws, status_range, 'Skipped',   C['amber_bg'])

    ws.auto_filter.ref = f"A1:{get_column_letter(len(sch_cols))}1"

    # Summary sheet
    ws2 = wb.create_sheet('Campaign Summary')
    _write_schedule_summary(ws2, schedule, year, month, prospect_count)

    wb.save(output_path)


def _write_schedule_summary(
    ws, schedule: List[Dict], year: int, month: int, prospect_count: int
):
    from utils.scheduler import get_working_days
    ws.column_dimensions['A'].width = 36
    ws.column_dimensions['B'].width = 20
    ws.column_dimensions['C'].width = 20
    ws.column_dimensions['D'].width = 20

    month_name = cal_mod.month_name[month]

    title = ws.cell(row=1, column=1, value=f'{month_name} {year} — Campaign Summary')
    title.font = Font(name='Calibri', bold=True, size=14, color='1E3A5F')

    working_days = get_working_days(year, month)
    unique_senders = {s['sender'] for s in schedule}

    stats = [
        ('Campaign Month',            f'{month_name} {year}'),
        ('Total Prospects',            prospect_count),
        ('Total Sends incl. Chasers',  prospect_count * 2),
        ('Working Days in Month',      len(working_days)),
        ('Scheduled Send Slots',       len(schedule)),
        ('Active Sender Accounts',     len(unique_senders)),
        ('Send Window (local time)',    '08:30 – 15:30'),
        ('Max Per Sender Per Day',      15),
        ('Min Global Gap Between Sends', '5 minutes'),
    ]

    for ri, (label, value) in enumerate(stats, 3):
        lc = ws.cell(row=ri, column=1, value=label)
        lc.font = Font(name='Calibri', bold=True, size=10)
        lc.fill = PatternFill('solid', fgColor='EBF5FB')
        vc = ws.cell(row=ri, column=2, value=value)
        vc.font = Font(name='Calibri', size=10)

    # Daily breakdown table
    row = len(stats) + 5
    ws.cell(row=row, column=1,
            value='Daily Breakdown').font = Font(bold=True, size=12, color='1E3A5F')
    row += 1

    day_hdrs = ['Date', 'Day', 'Sends Scheduled', 'Active Senders']
    for ci, h in enumerate(day_hdrs, 1):
        cell = ws.cell(row=row, column=ci, value=h)
        _hdr_cell(cell, C['hdr_schedule'], font_size=9)
    row += 1

    day_data: Dict[str, Dict] = {}
    for send in schedule:
        d = send['date']
        if d not in day_data:
            day_data[d] = {'day': send['day_of_week'], 'count': 0, 'senders': set()}
        day_data[d]['count'] += 1
        day_data[d]['senders'].add(send['sender'])

    fill_toggle = True
    for d, info in sorted(day_data.items()):
        fill_hex = 'EBF5FB' if fill_toggle else 'FFFFFF'
        fill_toggle = not fill_toggle
        vals = [d, info['day'], info['count'], len(info['senders'])]
        for ci, v in enumerate(vals, 1):
            cell = ws.cell(row=row, column=ci, value=v)
            _data_cell(cell, fill_hex)
        row += 1
