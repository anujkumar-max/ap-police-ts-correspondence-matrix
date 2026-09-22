import http.server
import socketserver
import json
import os
import sys
import urllib.parse
import urllib.request
import csv
import io
import re
from datetime import datetime, date

PORT = 8050
GOOGLE_SHEET_ID = "1xz_SMnLauFRVSTPaoRhltL3xTiuanIZqA45x96AI25Y"
GOOGLE_SHEET_URL = f"https://docs.google.com/spreadsheets/d/{GOOGLE_SHEET_ID}/edit?usp=sharing"
GOOGLE_CSV_URL = f"https://docs.google.com/spreadsheets/d/{GOOGLE_SHEET_ID}/gviz/tq?tqx=out:csv"
EXCEL_PATH = r"c:\Users\prasa\Desktop\TS-Correspondence Matrix\AP TS-Correspondence MATRIX.xlsx"
STATIC_DIR = os.path.dirname(os.path.abspath(__file__))

_cached_data = None
_last_fetch_time = 0

def clean_text(val):
    if val is None:
        return 'N/A'
    s = str(val).strip()
    if not s or s.lower() in ['n/a', 'none', 'null', '']:
        return 'N/A'
    # Sanitize unicode dashes and special characters
    s = s.replace('\ufffd', '-').replace('\u2013', '-').replace('\u2014', '-')
    return s

def parse_pro_id(val):
    cleaned = clean_text(val)
    if cleaned in ['N/A', 'ALL', 'Unassigned']:
        return {'id': 'N/A', 'name': 'Unassigned', 'rank': '', 'label': 'Unassigned'}
    
    # Regex match for PRO-xxx - Name (Rank)
    m = re.match(r'^(PRO-\d+)\s*[-:]\s*(.+?)(?:\s*\((.+?)\))?$', cleaned)
    if m:
        pro_id = m.group(1).strip()
        name = m.group(2).strip()
        rank = (m.group(3) or '').strip()
        label = f"{pro_id} - {name}" + (f" ({rank})" if rank else "")
        return {'id': pro_id, 'name': name, 'rank': rank, 'label': label}
    
    return {'id': 'PRO-GEN', 'name': cleaned, 'rank': '', 'label': cleaned}

def parse_prj_id(val):
    cleaned = clean_text(val)
    if cleaned in ['N/A', 'Miscellaneous']:
        return {'id': 'MISC', 'name': cleaned, 'label': cleaned}
    
    # Regex match for PRJ-xxx - Project Name
    m = re.match(r'^(PRJ-\d+)\s*[-:]\s*(.+)$', cleaned)
    if m:
        prj_id = m.group(1).strip()
        name = m.group(2).strip()
        label = f"{prj_id} - {name}"
        return {'id': prj_id, 'name': name, 'label': label}
    
    return {'id': 'PRJ-MISC', 'name': cleaned, 'label': cleaned}

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
    today_display = datetime.now().strftime("%d-%m-%Y | %I:%M %p")

    for r in records:
        d_recv = parse_indian_date(r.get('Date Received to office'))
        d_orig = parse_indian_date(r.get('Original Date of Letter/Mail'))
        d_close = parse_indian_date(r.get('Final Closure Date')) or parse_indian_date(r.get('Dispatched Date'))

        effective_start = d_recv or d_orig

        stage_raw = r.get('Current Stage') or ''
        r['status_group'] = categorize_stage(stage_raw)

        # Parse Pro-IDs & Project IDs
        off_info = parse_pro_id(r.get('Concerned Officer'))
        prj_info = parse_prj_id(r.get('Project'))

        r['officer_id'] = off_info['id']
        r['officer_name'] = off_info['name']
        r['officer_rank'] = off_info['rank']
        r['officer_label'] = off_info['label']

        r['project_id'] = prj_info['id']
        r['project_name'] = prj_info['name']
        r['project_label'] = prj_info['label']

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
        else: # On Hold / Closed / Stage N/A
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

    # Group by Officers with Complete Categorization (Pending, In-Progress, Stage N/A, Completed)
    officers_map = {}
    for r in records:
        off_label = r['officer_label']
        if off_label not in officers_map:
            officers_map[off_label] = {
                'officer': off_label,
                'pro_id': r['officer_id'],
                'name': r['officer_name'],
                'rank': r['officer_rank'],
                'total': 0, 'completed': 0, 'inprogress': 0, 'pending': 0, 'onhold': 0, 'critical': 0,
                'projects': set(),
                'pending_files': [],
                'inprogress_files': [],
                'na_stage_files': [],
                'completed_files': []
            }
        om = officers_map[off_label]
        om['total'] += 1

        file_entry = {
            'id': r.get('ID'),
            'project': r.get('project_label'),
            'prj_id': r.get('project_id'),
            'prj_name': r.get('project_name'),
            'subject': r.get('Subject / Work Description'),
            'stage': r.get('Current Stage') if r.get('Current Stage') else 'N/A',
            'age': r.get('display_time_metric'),
            'priority': r.get('Priority'),
            'recv_date': r.get('Date Received to office')
        }

        if r['status_group'] == 'Completed':
            om['completed'] += 1
            om['completed_files'].append(file_entry)
        elif r['status_group'] == 'In-Progress':
            om['inprogress'] += 1
            om['inprogress_files'].append(file_entry)
        elif r['status_group'] == 'Pending':
            om['pending'] += 1
            om['pending_files'].append(file_entry)
        else:
            om['onhold'] += 1
            om['na_stage_files'].append(file_entry)

        if r['Priority'] in ['Critical', 'High']:
            om['critical'] += 1
        
        if r['project_label'] not in ['N/A', '']:
            om['projects'].add(r['project_label'])

    officer_analytics = []
    for off_label, data in sorted(officers_map.items(), key=lambda x: (x[1]['pending'] + x[1]['inprogress'] + x[1]['onhold']), reverse=True):
        active_total = data['pending'] + data['inprogress'] + data['onhold']
        officer_analytics.append({
            'officer': off_label,
            'pro_id': data['pro_id'],
            'name': data['name'],
            'rank': data['rank'],
            'total': data['total'],
            'active_total': active_total,
            'completed': data['completed'],
            'inprogress': data['inprogress'],
            'pending': data['pending'],
            'onhold': data['onhold'],
            'critical': data['critical'],
            'projects': sorted(list(data['projects'])),
            'pending_files': data['pending_files'],
            'inprogress_files': data['inprogress_files'],
            'na_stage_files': data['na_stage_files'],
            'completed_files': data['completed_files']
        })

    # Group by Projects
    projects_map = {}
    for r in records:
        prj_label = r['project_label']
        if prj_label not in projects_map:
            projects_map[prj_label] = {
                'project': prj_label,
                'prj_id': r['project_id'],
                'name': r['project_name'],
                'total': 0, 'completed': 0, 'inprogress': 0, 'pending': 0, 'onhold': 0, 'critical': 0,
                'officers': set()
            }
        pm = projects_map[prj_label]
        pm['total'] += 1
        if r['status_group'] == 'Completed':
            pm['completed'] += 1
        elif r['status_group'] == 'In-Progress':
            pm['inprogress'] += 1
        elif r['status_group'] == 'Pending':
            pm['pending'] += 1
        else:
            pm['onhold'] += 1

        if r['Priority'] == 'Critical':
            pm['critical'] += 1
        
        if r['officer_label'] not in ['N/A', 'Unassigned']:
            pm['officers'].add(r['officer_label'])

    project_analytics = []
    for prj_label, data in sorted(projects_map.items(), key=lambda x: x[1]['total'], reverse=True):
        project_analytics.append({
            'project': prj_label,
            'prj_id': data['prj_id'],
            'name': data['name'],
            'total': data['total'],
            'completed': data['completed'],
            'inprogress': data['inprogress'],
            'pending': data['pending'],
            'onhold': data['onhold'],
            'critical': data['critical'],
            'officers': sorted(list(data['officers']))
        })

    # Channels
    channels_map = {}
    for r in records:
        ch = clean_text(r.get('Received Through'))
        if ch == 'N/A':
            ch = 'Other'
        channels_map[ch] = channels_map.get(ch, 0) + 1

    channel_analytics = [{'channel': k, 'count': v} for k, v in sorted(channels_map.items(), key=lambda x: x[1], reverse=True)]

    # Formulate Detailed WhatsApp Shareable Text
    active_officers = [o for o in officer_analytics if o['active_total'] > 0]
    wa_lines = [
        "🏛️ *AP POLICE TECHNICAL SERVICES (PCS&S)*",
        "📋 *OFFICER-WISE ACTIVE WORKLOAD & PENDING ABSTRACT*",
        f"📅 Date: {today_display}",
        "",
        "📊 *Overall Abstract:*",
        f"• Total Inflow: {total_records} Files",
        f"• 🟡 Pending Action: {pending_count} Files",
        f"• 🔵 Active In-Progress: {inprogress_count} Files",
        f"• ⚪ Stage N/A (Notice Required): {onhold_closed_count} Files",
        f"• 🟢 Completed / Dispatched: {completed_count} Files",
        "",
        "───────────────────────────────",
        "👤 *OFFICER-WISE WORKLOAD BREAKDOWN:*",
        "───────────────────────────────"
    ]

    for idx, o in enumerate(active_officers, 1):
        wa_lines.append("")
        wa_lines.append(f"{idx}. *{o['officer']}*: Total Active: {o['active_total']} ({o['pending']} Pending, {o['inprogress']} In-Prog, {o['onhold']} Stage N/A)")
        
        if o['pending_files']:
            wa_lines.append("   🟡 *Pending Action:*")
            for f in o['pending_files']:
                wa_lines.append(f"   - {f['id']} [{f['project']}]: ({f['age']})")
        
        if o['inprogress_files']:
            wa_lines.append("   🔵 *In-Progress:*")
            for f in o['inprogress_files']:
                wa_lines.append(f"   - {f['id']} [{f['project']}]: ({f['age']})")
        
        if o['na_stage_files']:
            wa_lines.append("   ⚪ *Stage N/A (Notice Required):*")
            for f in o['na_stage_files']:
                wa_lines.append(f"   - {f['id']} [{f['project']}]: ({f['age']})")

    wa_lines.append("")
    wa_lines.append("───────────────────────────────")
    wa_lines.append("🔗 *Update Status in Google Sheet:*")
    wa_lines.append(GOOGLE_SHEET_URL)
    wa_lines.append("───────────────────────────────")
    wa_lines.append("_Generated from AP Police TS Executive Command Portal_")
    whatsapp_text = "\n".join(wa_lines)

    # Formulate Option 2: Short Numbers-Only WhatsApp Text
    wa_short_lines = [
        "🏛️ *AP POLICE TECHNICAL SERVICES (PCS&S)*",
        "⚡ *DAILY BRIEF - CORRESPONDENCE STATUS*",
        f"📅 Date: {today_display}",
        "",
        f"📊 *Total: {total_records}* | 🟡 *Pend: {pending_count}* | 🔵 *In-Prog: {inprogress_count}* | ⚪ *N/A: {onhold_closed_count}* | 🟢 *Done: {completed_count}*",
        "",
        "───────────────────────────────",
        "👮 *OFFICER WORKLOAD [Active = Pend / InProg / N/A]:*",
        "───────────────────────────────"
    ]

    for idx, o in enumerate(active_officers, 1):
        wa_short_lines.append(f"{idx}. {o['officer']}: *{o['active_total']}* ({o['pending']} / {o['inprogress']} / {o['onhold']})")

    wa_short_lines.append("")
    wa_short_lines.append("───────────────────────────────")
    wa_short_lines.append("🔗 *Update Status in Google Sheet:*")
    wa_short_lines.append(GOOGLE_SHEET_URL)
    wa_short_lines.append("───────────────────────────────")
    wa_short_lines.append("_Generated from AP Police TS Executive Command Portal_")
    whatsapp_short_text = "\n".join(wa_short_lines)

    return {
        'status': 'success',
        'data_source': 'Google Sheets Live Cloud',
        'sheet_id': GOOGLE_SHEET_ID,
        'sheet_url': GOOGLE_SHEET_URL,
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
        'projects': project_analytics,
        'channels': channel_analytics,
        'records': records,
        'pending_abstract': {
            'total_pending': pending_count,
            'total_inprogress': inprogress_count,
            'total_na_stage': onhold_closed_count,
            'officers_active': active_officers,
            'whatsapp_text': whatsapp_text,
            'whatsapp_short_text': whatsapp_short_text
        }
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
            headers = [ws.cell(1, c).value for c in range(1, ws.max_column + 1) if ws.cell(1, c).value]
            records = []
            for r in range(2, ws.max_row + 1):
                row_id = ws.cell(r, 1).value
                if not row_id or str(row_id).strip() == "":
                    continue
                row_dict = {}
                has_data = False
                for c in range(1, len(headers) + 1):
                    h = str(headers[c - 1]).strip()
                    val = ws.cell(r, c).value
                    if isinstance(val, (datetime, date)):
                        row_dict[h] = val.strftime("%d/%m/%Y")
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
        print(f" Officer-Wise Pending & Active Workload Abstract Active")
        print(f"================================================================")
        httpd.serve_forever()

if __name__ == '__main__':
    start_server()
