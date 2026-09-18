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
STATIC_DIR = os.path.dirname(os.path.abspath(__file__))

_cached_data = None
_last_fetch_time = 0

def fetch_google_sheet_data():
    global _cached_data, _last_fetch_time
    now_ts = datetime.now().timestamp()

    # Cache for 5 seconds to prevent rate-limiting
    if _cached_data and (now_ts - _last_fetch_time < 5):
        _cached_data['server_time'] = datetime.now().strftime("%I:%M:%S %p")
        return _cached_data

    try:
        req = urllib.request.Request(
            GOOGLE_CSV_URL,
            headers={'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'}
        )
        with urllib.request.urlopen(req, timeout=8) as response:
            csv_text = response.read().decode('utf-8')

        reader = csv.reader(io.StringIO(csv_text))
        rows = list(reader)

        if len(rows) < 2:
            return {"status": "error", "message": "Google Sheet is empty or headers missing."}

        headers = [h.strip() for h in rows[0]]
        records = []
        now = datetime.now()

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
                date_recv_str = row_dict.get('Date Received to office', '')
                date_close_str = row_dict.get('Final Closure Date', '')
                
                aging_days = 0
                tat_days = 0

                if date_recv_str and date_recv_str != 'N/A':
                    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y", "%d-%m-%Y"):
                        try:
                            d_recv = datetime.strptime(date_recv_str.split(' ')[0], fmt)
                            if date_close_str and date_close_str != 'N/A':
                                for cfmt in ("%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y", "%d-%m-%Y"):
                                    try:
                                        d_close = datetime.strptime(date_close_str.split(' ')[0], cfmt)
                                        tat_days = max(0, (d_close - d_recv).days)
                                        aging_days = tat_days
                                        break
                                    except Exception:
                                        pass
                            else:
                                aging_days = max(0, (now - d_recv).days)
                            break
                        except Exception:
                            pass

                row_dict['calculated_aging_days'] = aging_days
                row_dict['calculated_tat_days'] = tat_days

                stage = (row_dict.get('Current Stage') or '').strip()
                if stage in ['Completed', 'Correspondence Closed', 'File Dispatched']:
                    row_dict['status_group'] = 'Completed'
                elif stage in ['in-progress', 'Draft Work Allocated', 'Draft Preparation In Progress', 'Draft Prepared - Sent For Approval', 'File Sent through eOffice', 'Work Completed-Pending for approval', 'Pending For Approval', 'Letter sent & Reply Required - Awaiting Replies']:
                    row_dict['status_group'] = 'In-Progress'
                else:
                    row_dict['status_group'] = 'Pending'

                priority = (row_dict.get('Priority') or '').strip()
                if priority not in ['Critical', 'High', 'Normal']:
                    priority = 'Normal'
                row_dict['Priority'] = priority

                records.append(row_dict)

        total_records = len(records)
        completed_count = sum(1 for r in records if r['status_group'] == 'Completed')
        inprogress_count = sum(1 for r in records if r['status_group'] == 'In-Progress')
        pending_count = sum(1 for r in records if r['status_group'] == 'Pending')
        critical_count = sum(1 for r in records if r['Priority'] == 'Critical')
        high_count = sum(1 for r in records if r['Priority'] == 'High')

        tat_list = [r['calculated_tat_days'] for r in records if r['status_group'] == 'Completed' and r['calculated_tat_days'] > 0]
        avg_tat = round(sum(tat_list) / len(tat_list), 1) if tat_list else 0

        aging_buckets = {"< 3 Days": 0, "4 - 7 Days": 0, "8 - 15 Days": 0, "15+ Days": 0}
        for r in records:
            if r['status_group'] != 'Completed':
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
                officers_map[off] = {'total': 0, 'completed': 0, 'inprogress': 0, 'pending': 0, 'critical': 0, 'staff': set()}
            officers_map[off]['total'] += 1
            if r['status_group'] == 'Completed':
                officers_map[off]['completed'] += 1
            elif r['status_group'] == 'In-Progress':
                officers_map[off]['inprogress'] += 1
            else:
                officers_map[off]['pending'] += 1
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
                'critical': data['critical'],
                'staff': sorted(list(data['staff']))
            })

        staff_map = {}
        for r in records:
            st = r.get('Concerned Staff') or 'Unassigned'
            if st in ['', 'N/A', 'None']:
                st = 'Unassigned'
            if st not in staff_map:
                staff_map[st] = {'total': 0, 'completed': 0, 'inprogress': 0, 'pending': 0}
            staff_map[st]['total'] += 1
            if r['status_group'] == 'Completed':
                staff_map[st]['completed'] += 1
            elif r['status_group'] == 'In-Progress':
                staff_map[st]['inprogress'] += 1
            else:
                staff_map[st]['pending'] += 1

        staff_analytics = []
        for st, data in sorted(staff_map.items(), key=lambda x: x[1]['total'], reverse=True):
            staff_analytics.append({
                'staff': st,
                'total': data['total'],
                'completed': data['completed'],
                'inprogress': data['inprogress'],
                'pending': data['pending']
            })

        projects_map = {}
        for r in records:
            p = r.get('Project') or 'Other / Misc'
            if p in ['', 'N/A', 'None']:
                p = 'Other / Misc'
            if p not in projects_map:
                projects_map[p] = {'total': 0, 'completed': 0, 'inprogress': 0, 'pending': 0, 'critical': 0}
            projects_map[p]['total'] += 1
            if r['status_group'] == 'Completed':
                projects_map[p]['completed'] += 1
            elif r['status_group'] == 'In-Progress':
                projects_map[p]['inprogress'] += 1
            else:
                projects_map[p]['pending'] += 1
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
                'critical': data['critical']
            })

        channels_map = {}
        for r in records:
            ch = r.get('Received Through') or 'Other'
            if ch in ['', 'N/A', 'None']:
                ch = 'Other'
            channels_map[ch] = channels_map.get(ch, 0) + 1

        channel_analytics = [{'channel': k, 'count': v} for k, v in sorted(channels_map.items(), key=lambda x: x[1], reverse=True)]

        _cached_data = {
            'status': 'success',
            'data_source': 'Google Sheets Live Cloud',
            'sheet_id': GOOGLE_SHEET_ID,
            'last_updated': datetime.now().strftime("%Y-%m-%d %I:%M:%S %p"),
            'server_time': datetime.now().strftime("%I:%M:%S %p"),
            'kpis': {
                'total': total_records,
                'completed': completed_count,
                'inprogress': inprogress_count,
                'pending': pending_count,
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
        _last_fetch_time = now_ts
        return _cached_data

    except urllib.error.HTTPError as e:
        if e.code == 401:
            return {
                'status': 'unauthorized',
                'message': 'Google Sheet access is Restricted. Please set General Access to "Anyone with the link can view".'
            }
        return {'status': 'error', 'message': f'HTTP Error {e.code}: {e.reason}'}
    except Exception as e:
        return {'status': 'error', 'message': str(e)}

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
        print(f" AP Police TS Correspondence Live Google Sheets Dashboard Server ")
        print(f" Running live on: http://localhost:{PORT}")
        print(f" Google Sheet ID: {GOOGLE_SHEET_ID}")
        print(f" Real-Time Sync: Active (Every 5-10s)")
        print(f"================================================================")
        httpd.serve_forever()

if __name__ == '__main__':
    start_server()
