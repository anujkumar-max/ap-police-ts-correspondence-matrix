import http.server
import socketserver
import json
import os
import sys
import urllib.parse
import urllib.request
import csv
import io
from datetime import datetime, date

PORT = 8050
GOOGLE_SHEET_ID = "1xz_SMnLauFRVSTPaoRhltL3xTiuanIZqA45x96AI25Y"
GOOGLE_CSV_URL = f"https://docs.google.com/spreadsheets/d/{GOOGLE_SHEET_ID}/gviz/tq?tqx=out:csv"
EXCEL_PATH = r"c:\Users\prasa\Desktop\TS-Correspondence Matrix\AP TS-Correspondence MATRIX.xlsx"
STATIC_DIR = os.path.dirname(os.path.abspath(__file__))

_cached_data = None
_last_fetch_time = 0

def parse_indian_date(val):
    if not val:
        return None
    if isinstance(val, (datetime, date)):
        return val if isinstance(val, date) else val.date()
    val_str = str(val).strip()
    if not val_str or val_str.lower() in ['n/a', 'none', '']:
        return None
    
    clean_val = val_str.split(' ')[0].split('T')[0]
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%m/%d/%Y", "%d.%m.%Y"):
        try:
            return datetime.strptime(clean_val, fmt).date()
        except Exception:
            pass
    return None

def categorize_stage(stage_raw):
    stage = str(stage_raw).strip() if stage_raw is not None else ""
    stage_lower = stage.lower()

    # 1. On Hold / Cancelled / Closed / Inactive / N/A / Blank
    if not stage or stage_lower in ['n/a', 'none', 'null', '', 'on hold', 'cancelled', 'correspondence closed', 'inactive']:
        return 'On Hold / Closed'

    # 2. Any stage mentioning "pending" (case-insensitive)
    if 'pending' in stage_lower:
        return 'Pending'

    # 3. Completed / Dispatched / Acknowledgement
    if stage_lower in ['completed', 'file dispatched', 'acknowledgement filed']:
        return 'Completed'

    # 4. Remaining all are In-Progress
    return 'In-Progress'

def calculate_analytics_from_records(records):
    today = datetime.now().date()
    now_str = datetime.now().strftime("%Y-%m-%d %I:%M:%S %p")

    for r in records:
        d_recv = parse_indian_date(r.get('Date Received to office'))
        d_orig = parse_indian_date(r.get('Original Date of Letter/Mail'))
        d_close = parse_indian_date(r.get('Final Closure Date')) or parse_indian_date(r.get('Dispatched Date'))

        effective_start = d_recv or d_orig

        stage_raw = r.get('Current Stage') or ''
        r['status_group'] = categorize_stage(stage_raw)

        tat_days = 0
        age_days = 0

        if r['status_group'] == 'Completed':
            if effective_start and d_close:
                tat_days = max(0, (d_close - effective_start).days)
            elif effective_start:
                tat_days = max(0, (today - effective_start).days)
            r['calculated_tat_days'] = tat_days
            r['calculated_aging_days'] = tat_days
            r['display_time_metric'] = f"{tat_days}d (TAT)" if effective_start else "—"
        elif r['status_group'] in ['Pending', 'In-Progress']:
            if effective_start:
                age_days = max(0, (today - effective_start).days)
                r['calculated_aging_days'] = age_days
                r['display_time_metric'] = f"{age_days}d (Age)"
            else:
                r['calculated_aging_days'] = 0
                r['display_time_metric'] = "—"
            r['calculated_tat_days'] = 0
        else: # On Hold / Closed
            if effective_start and d_close:
                tat_days = max(0, (d_close - effective_start).days)
                r['display_time_metric'] = f"{tat_days}d (TAT)"
            elif effective_start:
                age_days = max(0, (today - effective_start).days)
                r['display_time_metric'] = f"{age_days}d (Age)"
            else:
                r['display_time_metric'] = "—"
            r['calculated_tat_days'] = 0
            r['calculated_aging_days'] = 0

        priority = (r.get('Priority') or '').strip()
        if priority not in ['Critical', 'High', 'Normal']:
            priority = 'Normal'
        r['Priority'] = priority

    total_records = len(records)
    completed_count = sum(1 for r in records if r['status_group'] == 'Completed')
    inprogress_count = sum(1 for r in records if r['status_group'] == 'In-Progress')
    pending_count = sum(1 for r in records if r['status_group'] == 'Pending')
    onhold_closed_count = sum(1 for r in records if r['status_group'] == 'On Hold / Closed')
    critical_count = sum(1 for r in records if r['Priority'] == 'Critical')
    high_count = sum(1 for r in records if r['Priority'] == 'High')

    tat_list = [r['calculated_tat_days'] for r in records if r['status_group'] == 'Completed' and r['calculated_tat_days'] > 0]
    avg_tat = round(sum(tat_list) / len(tat_list), 1) if tat_list else 0

    aging_buckets = {"< 3 Days": 0, "4 - 7 Days": 0, "8 - 15 Days": 0, "15+ Days": 0}
    for r in records:
        if r['status_group'] in ['Pending', 'In-Progress'] and r['display_time_metric'] != '—':
            age = r['calculated_aging_days']
            if age <= 3:
                aging_buckets["< 3 Days"] += 1
            elif age <= 7:
                aging_buckets["4 - 7 Days"] += 1
            elif age <= 15:
                aging_buckets["8 - 15 Days"] += 1
            else:
                aging_buckets["15+ Days"] += 1

    officers_map = {}
    for r in records:
        off = r.get('Concerned Officer') or 'Unassigned'
        if off in ['', 'N/A', 'None']:
            off = 'Unassigned'
        if off not in officers_map:
            officers_map[off] = {'total': 0, 'completed': 0, 'inprogress': 0, 'pending': 0, 'onhold': 0, 'critical': 0, 'staff': set()}
        officers_map[off]['total'] += 1
        if r['status_group'] == 'Completed':
            officers_map[off]['completed'] += 1
        elif r['status_group'] == 'In-Progress':
            officers_map[off]['inprogress'] += 1
        elif r['status_group'] == 'Pending':
            officers_map[off]['pending'] += 1
        else:
            officers_map[off]['onhold'] += 1

        if r['Priority'] in ['Critical', 'High']:
            officers_map[off]['critical'] += 1
        st = r.get('Concerned Staff', '')
        if st and st not in ['', 'N/A', 'None']:
            officers_map[off]['staff'].add(st)

    officer_analytics = []
    for off, data in sorted(officers_map.items(), key=lambda x: x[1]['total'], reverse=True):
        officer_analytics.append({
            'officer': off,
            'total': data['total'],
            'completed': data['completed'],
            'inprogress': data['inprogress'],
            'pending': data['pending'],
            'onhold': data['onhold'],
            'critical': data['critical'],
            'staff': sorted(list(data['staff']))
        })

    staff_map = {}
    for r in records:
        st = r.get('Concerned Staff') or 'Unassigned'
        if st in ['', 'N/A', 'None']:
            st = 'Unassigned'
        if st not in staff_map:
            staff_map[st] = {'total': 0, 'completed': 0, 'inprogress': 0, 'pending': 0, 'onhold': 0}
        staff_map[st]['total'] += 1
        if r['status_group'] == 'Completed':
            staff_map[st]['completed'] += 1
        elif r['status_group'] == 'In-Progress':
            staff_map[st]['inprogress'] += 1
        elif r['status_group'] == 'Pending':
            staff_map[st]['pending'] += 1
        else:
            staff_map[st]['onhold'] += 1

    staff_analytics = []
    for st, data in sorted(staff_map.items(), key=lambda x: x[1]['total'], reverse=True):
        staff_analytics.append({
            'staff': st,
            'total': data['total'],
            'completed': data['completed'],
            'inprogress': data['inprogress'],
            'pending': data['pending'],
            'onhold': data['onhold']
        })

    projects_map = {}
    for r in records:
        p = r.get('Project') or 'Other / Misc'
        if p in ['', 'N/A', 'None']:
            p = 'Other / Misc'
        if p not in projects_map:
            projects_map[p] = {'total': 0, 'completed': 0, 'inprogress': 0, 'pending': 0, 'onhold': 0, 'critical': 0}
        projects_map[p]['total'] += 1
        if r['status_group'] == 'Completed':
            projects_map[p]['completed'] += 1
        elif r['status_group'] == 'In-Progress':
            projects_map[p]['inprogress'] += 1
        elif r['status_group'] == 'Pending':
            projects_map[p]['pending'] += 1
        else:
            projects_map[p]['onhold'] += 1
        if r['Priority'] == 'Critical':
            projects_map[p]['critical'] += 1

    project_analytics = []
    for p, data in sorted(projects_map.items(), key=lambda x: x[1]['total'], reverse=True):
        project_analytics.append({
            'project': p,
            'total': data['total'],
            'completed': data['completed'],
            'inprogress': data['inprogress'],
            'pending': data['pending'],
            'onhold': data['onhold'],
            'critical': data['critical']
        })

    channels_map = {}
    for r in records:
        ch = r.get('Received Through') or 'Other'
        if ch in ['', 'N/A', 'None']:
            ch = 'Other'
        channels_map[ch] = channels_map.get(ch, 0) + 1

    channel_analytics = [{'channel': k, 'count': v} for k, v in sorted(channels_map.items(), key=lambda x: x[1], reverse=True)]

    return {
        'status': 'success',
        'data_source': 'Google Sheets Live Cloud',
        'sheet_id': GOOGLE_SHEET_ID,
        'last_updated': now_str,
        'server_time': datetime.now().strftime("%I:%M:%S %p"),
        'kpis': {
            'total': total_records,
            'completed': completed_count,
            'inprogress': inprogress_count,
            'pending': pending_count,
            'onhold_closed': onhold_closed_count,
            'critical': critical_count,
            'high': high_count,
            'avg_tat_days': avg_tat
        },
        'aging_buckets': aging_buckets,
        'officers': officer_analytics,
        'staff': staff_analytics,
        'projects': project_analytics,
        'channels': channel_analytics,
        'records': records
    }

def fetch_google_sheet_data():
    global _cached_data, _last_fetch_time
    now_ts = datetime.now().timestamp()

    if _cached_data and (now_ts - _last_fetch_time < 5):
        _cached_data['server_time'] = datetime.now().strftime("%I:%M:%S %p")
        return _cached_data

    # Try Google Sheets
    try:
        req = urllib.request.Request(
            GOOGLE_CSV_URL,
            headers={'User-Agent': 'Mozilla/5.0'}
        )
        with urllib.request.urlopen(req, timeout=8) as response:
            csv_text = response.read().decode('utf-8')

        reader = csv.reader(io.StringIO(csv_text))
        rows = list(reader)

        if len(rows) >= 2:
            headers = [h.strip() for h in rows[0]]
            records = []
            for r_idx in range(1, len(rows)):
                row = rows[r_idx]
                if not row or not row[0] or str(row[0]).strip() == "":
                    continue
                row_dict = {}
                has_data = False
                for c_idx, h in enumerate(headers):
                    val = row[c_idx].strip() if c_idx < len(row) else ""
                    row_dict[h] = val
                    if c_idx > 0 and val not in ['', 'N/A', 'None']:
                        has_data = True
                if has_data:
                    records.append(row_dict)

            _cached_data = calculate_analytics_from_records(records)
            _last_fetch_time = now_ts
            return _cached_data

    except Exception:
        pass

    # Fallback to local Excel file if Google Sheet is restricted
    if os.path.exists(EXCEL_PATH):
        try:
            import openpyxl
            wb = openpyxl.load_workbook(EXCEL_PATH, data_only=True)
            ws = wb['Correspondence Matrix (Tappals)']
            headers = [ws.cell(1, c).value for c in range(1, 22)]
            records = []
            for r in range(2, ws.max_row + 1):
                row_id = ws.cell(r, 1).value
                if not row_id or str(row_id).strip() == "":
                    continue
                row_dict = {}
                has_data = False
                for c in range(1, 22):
                    h = headers[c - 1]
                    val = ws.cell(r, c).value
                    if isinstance(val, (datetime, date)):
                        row_dict[h] = val.strftime("%Y-%m-%d")
                        has_data = True
                    elif val is not None:
                        val_str = str(val).strip()
                        row_dict[h] = val_str
                        if c > 1 and val_str not in ['', 'N/A', 'None']:
                            has_data = True
                    else:
                        row_dict[h] = ""
                if has_data:
                    records.append(row_dict)

            _cached_data = calculate_analytics_from_records(records)
            _cached_data['data_source'] = 'Local Excel / Fallback Snapshot'
            _last_fetch_time = now_ts
            return _cached_data
        except Exception as e:
            return {'status': 'error', 'message': str(e)}

    return {"status": "error", "message": "Unable to fetch correspondence data."}

class DashboardRequestHandler(http.server.SimpleHTTPRequestHandler):
    def do_GET(self):
        parsed = urllib.parse.urlparse(self.path)
        path = parsed.path

        if path == '/api/data':
            data = fetch_google_sheet_data()
            self.send_response(200)
            self.send_header('Content-Type', 'application/json; charset=utf-8')
            self.send_header('Access-Control-Allow-Origin', '*')
            self.send_header('Cache-Control', 'no-cache, no-store, must-revalidate')
            self.end_headers()
            self.wfile.write(json.dumps(data, ensure_ascii=False).encode('utf-8'))
            return

        elif path == '/' or path == '/index.html':
            index_path = os.path.join(STATIC_DIR, 'index.html')
            if os.path.exists(index_path):
                self.send_response(200)
                self.send_header('Content-Type', 'text/html; charset=utf-8')
                self.end_headers()
                with open(index_path, 'rb') as f:
                    self.wfile.write(f.read())
                return

        return super().do_GET()

def start_server():
    os.chdir(STATIC_DIR)
    socketserver.TCPServer.allow_reuse_address = True
    with socketserver.TCPServer(("", PORT), DashboardRequestHandler) as httpd:
        print(f"================================================================")
        print(f" AP Police TS Correspondence Executive Analytics Dashboard Server ")
        print(f" Running live on: http://localhost:{PORT}")
        print(f" Stage Grouping: Pending, In-Progress, Completed, On Hold/Closed")
        print(f"================================================================")
        httpd.serve_forever()

if __name__ == '__main__':
    start_server()
