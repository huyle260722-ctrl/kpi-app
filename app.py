from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime
import io

app = Flask(__name__, static_folder='static', static_url_path='')
CORS(app)

# ── Serve frontend ──────────────────────────────────────────────
@app.route('/')
def index():
    return app.send_static_file('index.html')

# ── Excel generation endpoint ───────────────────────────────────
@app.route('/api/export', methods=['POST'])
def export_excel():
    payload = request.get_json()
    year_month  = payload.get('month', '')   # "2026-06"
    report_data = payload.get('data', {})    # { "2026-06-01": [rows] }

    if not year_month or not report_data:
        return jsonify({'error': 'Thiếu dữ liệu'}), 400

    try:
        wb = build_workbook(report_data, year_month)
        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)
        y, m = year_month.split('-')
        filename = f'Bao_Cao_Thang_{m}_{y}.xlsx'
        return send_file(
            buf,
            as_attachment=True,
            download_name=filename,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
    except Exception as e:
        return jsonify({'error': str(e)}), 500


# ── Excel builder ───────────────────────────────────────────────
def get_week(day):
    if day <= 7:  return 1
    if day <= 14: return 2
    if day <= 21: return 3
    return 4

def build_workbook(report_data, year_month):
    y, m = year_month.split('-')

    C_TITLE   = '5B95F9'
    C_HDR     = 'E8F0FE'
    C_KPI_DAY = 'CFE2F3'
    C_WEEK    = 'C9DAF8'
    C_WHITE   = 'FFFFFF'

    def fill(c):   return PatternFill('solid', fgColor=c)
    def font(bold=False, size=11, color='000000'):
        return Font(name='Calibri', bold=bold, size=size, color=color)
    def align(h='center', v='center', wrap=True):
        return Alignment(horizontal=h, vertical=v, wrap_text=wrap)

    thin  = Side(style='thin')
    thick = Side(style='thick')
    B_thin  = Border(top=thin,  bottom=thin,  left=thin,  right=thin)
    B_thick = Border(top=thick, bottom=thick, left=thick, right=thick)

    def B_day(is_first, is_last, is_only):
        t = thick if is_first or is_only else thin
        b = thick if is_last  or is_only else thin
        return Border(top=t, bottom=b, left=thick, right=thin)

    B_right_thick = Border(top=thin, bottom=thin, left=thin, right=thick)
    B_hdr_right   = Border(top=thin, bottom=thin, left=thin, right=thick)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = f'{m}{y}'

    col_w = [26, 38.57, 26.71, 28.86, 24.57, 14.57, 23.57, 28.43, 29.43, 23.71, 28.86, 19.86]
    for i, w in enumerate(col_w, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    ws.row_dimensions[1].height = 14.25
    ws.row_dimensions[2].height = 31.5
    ws.row_dimensions[3].height = 70.5

    # Title row
    ws.merge_cells('A1:L2')
    c = ws['A1']
    c.value = 'THÔNG TIN USER-TÊN NHÓM (用户信息)'
    c.font  = font(bold=True, size=13, color='FFFFFF')
    c.fill  = fill(C_TITLE)
    c.alignment = align()
    c.border = B_thick

    # Header row
    headers = [
        'Ngày\n(日)', 'TÊN USER\n(用户名)', 'CỦA THÀNH VIÊN (成员的）',
        'KPI Ngày\n(日KPI)', 'Cấp hiện tại\n(当前等级)', 'Giới tính\n(性别)',
        'Nơi sống hiện tại\n(目前居住地)',
        'Nhắn tin thêm trên nền tảng khác\n(是否在其他平台聊天)',
        'KPI TỔNG NGÀY\n(每天KPI总计)', 'KPI TỔNG TUẦN\n(每周KPI总计)',
        'TUẦN\n(周)', 'Ghi chú\n(备注)'
    ]
    for col_i, hdr in enumerate(headers, 1):
        c = ws.cell(row=3, column=col_i, value=hdr)
        c.font = font(bold=True, size=10)
        c.fill = fill(C_HDR)
        c.alignment = align()
        c.border = B_thin

    dates = sorted(report_data.keys())
    week_dates  = {1:[], 2:[], 3:[], 4:[]}
    week_totals = {1:0,  2:0,  3:0,  4:0}
    for d in dates:
        wk = get_week(int(d.split('-')[2]))
        week_dates[wk].append(d)
        week_totals[wk] += sum(r['kpi'] for r in report_data[d])

    current_row    = 4
    week_first_row = {}
    week_last_row  = {}
    day_row_ranges = {}

    for d in dates:
        rows = report_data[d]
        day  = int(d.split('-')[2])
        wk   = get_week(day)
        dt   = datetime(int(y), int(m), day)
        n    = len(rows)
        first_r = current_row

        if wk not in week_first_row:
            week_first_row[wk] = current_row

        for i, r in enumerate(rows):
            ws.row_dimensions[current_row].height = 45
            is_first = (i == 0)
            is_last  = (i == n - 1)
            is_only  = (n == 1)

            # A: Ngày
            c = ws.cell(row=current_row, column=1)
            if is_first:
                c.value = dt
                c.number_format = 'D/M'
            c.font = font(bold=True)
            c.fill = fill(C_WHITE)
            c.alignment = align()
            c.border = B_day(is_first, is_last, is_only)

            # B: userName
            c = ws.cell(row=current_row, column=2, value=r['userName'])
            c.font = font(); c.fill = fill(C_WHITE)
            c.alignment = align(); c.border = B_thin

            # C: memberName
            c = ws.cell(row=current_row, column=3, value=r['memberName'])
            c.font = font(); c.fill = fill(C_WHITE)
            c.alignment = align(); c.border = B_thin

            # D: KPI ngày
            c = ws.cell(row=current_row, column=4, value=r['kpi'])
            c.font = font(); c.fill = fill(C_WHITE)
            c.alignment = align(); c.border = B_thin
            c.number_format = '#,##0'

            # E-G: empty
            for col_i in [5, 6, 7]:
                c = ws.cell(row=current_row, column=col_i)
                c.fill = fill(C_WHITE); c.alignment = align(); c.border = B_thin

            # H: x
            c = ws.cell(row=current_row, column=8, value='x')
            c.font = font(); c.fill = fill(C_WHITE)
            c.alignment = align(); c.border = B_thin

            # I: KPI tổng ngày (blue, merge later)
            c = ws.cell(row=current_row, column=9)
            c.fill = fill(C_KPI_DAY); c.font = font(bold=True)
            c.alignment = align(); c.border = B_thin
            c.number_format = '#,##0'

            # J: KPI tổng tuần (merge later)
            c = ws.cell(row=current_row, column=10)
            c.fill = fill(C_WHITE); c.font = font(bold=True)
            c.alignment = align(); c.border = B_right_thick
            c.number_format = '#,##0'

            # K: Tuần label (merge later)
            c = ws.cell(row=current_row, column=11)
            c.fill = fill(C_WEEK); c.font = font(bold=True)
            c.alignment = align(); c.border = B_thin

            # L: Ghi chú
            c = ws.cell(row=current_row, column=12)
            c.fill = fill(C_HDR); c.font = font(bold=True)
            c.alignment = align(); c.border = B_hdr_right

            current_row += 1

        last_r = current_row - 1
        day_row_ranges[d] = (first_r, last_r)
        week_last_row[wk] = last_r

    # Merge A per day + set KPI tổng ngày
    for d, (r1, r2) in day_row_ranges.items():
        day_total = sum(r['kpi'] for r in report_data[d])
        ws.cell(row=r1, column=9).value = day_total
        if r2 > r1:
            ws.merge_cells(f'A{r1}:A{r2}')
            ws.merge_cells(f'I{r1}:I{r2}')

    # Merge J, K per week + set values
    week_labels = {
        1: 'TUẦN 1: (第1周):   ', 2: 'TUẦN 2: (第2周):   ',
        3: 'TUẦN 3: (第三周):   ', 4: 'TUẦN 4: (第四周):   '
    }
    for wk in [1, 2, 3, 4]:
        if wk not in week_first_row: continue
        r1, r2 = week_first_row[wk], week_last_row[wk]
        ws.cell(row=r1, column=10).value = week_totals[wk]
        ws.cell(row=r1, column=11).value = week_labels[wk]
        if r2 > r1:
            ws.merge_cells(f'J{r1}:J{r2}')
            ws.merge_cells(f'K{r1}:K{r2}')

    # Footer
    grand = sum(week_totals.values())
    fr1, fr2 = current_row, current_row + 1
    ws.row_dimensions[fr1].height = 40
    ws.row_dimensions[fr2].height = 40

    def footer_row(row_num, label, value):
        ws.merge_cells(f'A{row_num}:H{row_num}')
        c = ws.cell(row=row_num, column=1, value=label)
        c.font = font(bold=True, size=12); c.fill = fill(C_WEEK)
        c.alignment = align(); c.border = B_thin
        ws.merge_cells(f'I{row_num}:L{row_num}')
        c = ws.cell(row=row_num, column=9, value=value)
        c.font = font(bold=True, size=14); c.fill = fill(C_WHITE)
        c.alignment = align(); c.border = B_thin
        c.number_format = '#,##0'

    footer_row(fr1, 'x', grand)
    footer_row(fr2, 'TỔNG KPI ĐỦ SỐ (KPI全部 - 达)', grand)

    # Sheet 2: Data
    ws2 = wb.create_sheet('Data')
    ws2.append(['Date', 'UserName', 'MemberName', 'KPI', 'Week'])
    for d in dates:
        for r in report_data[d]:
            ws2.append([d, r['userName'], r['memberName'], r['kpi'],
                        f"Tuần {get_week(int(d.split('-')[2]))}"])
    for col, w in zip(['A','B','C','D','E'], [14,25,20,12,12]):
        ws2.column_dimensions[col].width = w

    return wb


if __name__ == '__main__':
    app.run(debug=True, port=5000)
