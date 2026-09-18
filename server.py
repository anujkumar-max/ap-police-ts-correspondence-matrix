import http.server
import socketserver
import json
import os
import sys
import urllib.parse
from datetime import datetime, date
import openpyxl

PORT = 8050
EXCEL_PATH = r"c:\Users\prasa\Desktop\TS-Correspondence Matrix\AP TS-Correspondence MATRIX.xlsx"
STATIC_DIR = os.path.dirname(os.path.abspath(__file__))

_cached_data = None
_cached_mtime = None

def get_correspondence_data(force_reload=False):
    global _cached_data, _cached_mtime

    if not os.path.exists(EXCEL_PATH):
        return {"status": "error", "message": "Excel file not found at " + EXCEL_PATH}

    current_mtime = os.path.getmtime(EXCEL_PATH)

    # Return cached data if file hasn't changed on disk
    if not force_reload and _cached_data is not None and _cached_mtime == current_mtime:
        # Update server poll timestamp
        _cached_data['server_time'] = datetime.now().strftime("%I:%M:%S %p")
        return _cached_data

    try:
        wb = openpyxl.load_workbook(EXCEL_PATH, data_only=True)
        ws_name = 'Correspondence Matrix (Tappals)' if 'Correspondence Matrix (Tappals)' in wb.sheetnames else wb.sheetnames[0]
        ws = wb[ws_name]

        headers = []
        for c in range(1, 22):
            val = ws.cell(row=1, column=c).value
            headers.append(str(val).strip() if val else f"Col_{c}")

        records = []
        now = datetime.now()

        for r in range(2, ws.max_row + 1):
            row_id = ws.cell(row=r, column=1).value
            if not row_id or str(row_id).strip() == "" or str(row_id).strip().lower() == 'none':
                continue

            row_dict = {}
            has_substantive_data = False

            for c in range(1, 22):
                val = ws.cell(row=r, column=c).value
                col_name = headers[c - 1]

                if isinstance(val, (datetime, date)):
                    row_dict[col_name] = val.strftime("%Y-%m-%d")
                    has_substantive_data = True
                elif val is not None:
                    val_str = str(val).strip()
                    row_dict[col_name] = val_str
                    if c > 1 and val_str not in ['', 'N/A', 'None']:
                        has_substantive_data = True
                else:
                    row_dict[col_name] = ""

            if has_substantive_data:
                date_recv_str = row_dict.get('Date Received to office', '')
                date_close_str = row_dict.get('Final Closure Date', '')
                
                aging_days = 0
                tat_days = 0

                if date_recv_str and date_recv_str != 'N/A':
                    try:
                        d_recv = datetime.strptime(date_recv_str, "%Y-%m-%d")
                        if date_close_str and date_close_str != 'N/A':
                            d_close = datetime.strptime(date_close_str, "%Y-%m-%d")
                            tat_days = max(0, (d_close - d_recv).days)
                            aging_days = tat_days
                        else:
                            aging_days = max(0, (now - d_recv).days)
                    except Exception:
                        pass

                row_dict['calculated_aging_days'] = aging_days
                row_dict['calculated_tat_days'] = tat_days

                stage = (row_dict.get('Current Stage') or '').strip()
                if stage in ['Completed', 'Correspondence Closed', 'File Dispatched']:
                    row_dict['status_group'] = 'Completed'
                elif stage in ['in-progress', 'Draft Work Allocated', 'Draft Preparation In Progress', 'Draft Prepared - Sent For Approval', 'File Sent through eOffice', 'Work Completed-Pending for approval', 'Pending For Approval', 'Letter sent & Reply Required - Awaiting Replies']:
                    row_dict['status_group'] = 'In-Progress'
                elif stage in ['Pending', 'Decision Pending']:
                    row_dict['status_group'] = 'Pending'
                else:
                    row_dict['status_group'] = 'Pending' if not stage or stage == 'N/A' else stage

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

        dept_map = {}
        for r in records:
            d = r.get('Received From Department /Wing') or 'Unknown'
            if d in ['', 'N/A', 'None']:
                d = 'Direct / Internal'
            dept_map[d] = dept_map.get(d, 0) + 1
        
        dept_analytics = [{'department': k, 'count': v} for k, v in sorted(dept_map.items(), key=lambda x: x[1], reverse=True)[:8]]

        _cached_data = {
            'status': 'success',
            'file_modified_time': datetime.fromtimestamp(current_mtime).strftime("%Y-%m-%d %I:%M:%S %p"),
            'server_time': datetime.now().strftime("%I:%M:%S %p"),
            'file_name': os.path.basename(EXCEL_PATH),
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
            'departments': dept_analytics,
            'records': records
        }
        _cached_mtime = current_mtime

        return _cached_data

    except Exception as e:
        import traceback
        traceback.print_exc()
        return {'status': 'error', 'message': str(e)}

class DashboardRequestHandler(http.server.SimpleHTTPRequestHandler):
    def do_GET(self):
        parsed = urllib.parse.urlparse(self.path)
        path = parsed.path

        if path == '/api/data':
            data = get_correspondence_data()
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
        print(f" Network access: http://0.0.0.0:{PORT}")
        print(f" Tracking Excel: {EXCEL_PATH}")
        print(f" Auto-Sync: Active (Instant Disk Detection)")
        print(f"================================================================")
        httpd.serve_forever()

if __name__ == '__main__':
    start_server()
