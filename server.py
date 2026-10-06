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

VIP_DEPTS = ['dgp', 'high court', 'highcourt', 'rtgs', 'ite&c', 'ite & c', 'mha', 'ncrb']
VIP_STAGES = ['file with dgp', 'file with govt', 'pending for dgp approval', 'pending with dgp', 'pending for dgp']
VIP_DESIGS = [
    'dgp', 'director general of police',
    'secretary', 'secreatry',
    'principal secretary', 'chief secretary', 'cs to the govt', 'cs to govt',
    'registrar', 'registar',
    'adg', 'addl.director general', 'additional director general',
    'igp', 'inspector general of police', 'deputy inspector general',
    'deputy director'
]

def clean_text(val):
    if val is None:
        return 'N/A'
    s = str(val).strip()
    if not s or s.lower() in ['n/a', 'none', 'null', '']:
        return 'N/A'
    s = s.replace('\ufffd', '-').replace('\u2013', '-').replace('\u2014', '-')
    return s

# Official AP Police TS Officer Seniority Hierarchy
OFFICER_HIERARCHY_IDS = [
    'PRO-045', # K. Sreelakshmi (SP)
    'PRO-074', # G. Veeraraghava Reddy (Addl. SP)
    'PRO-078', # V. Vishnu Swaroop (DSP)
    'PRO-079', # M. Hema Latha (DSP)
    'PRO-081', # P. Bhavana (DSP)
    'PRO-082', # P. Sindhu Priya (DSP)
    'PRO-159', # N. Sarojini (AAO)
    'PRO-092', # V. Sudharshana Reddy (CI)
    'PRO-094', # K. Satish (CI)
    'PRO-095', # K. Vijaya Kumar (CI)
    'PRO-106', # K. Sreekanth (CI)
    'PRO-123', # M. Manohar Rao (CI)
    'PRO-091', # M. Mohan (CI)
    'PRO-153', # B. Vijay Kumar Reddy (SI)
    'PRO-137', # G. Ravi Kiran (SI)
    'PRO-124', # G. Jyothi (SI)
    'PRO-140', # MD. Sadhik (SI)
    'PRO-136', # T. Anoj Kumar (SI)
    'PRO-154', # CJ. Bharath (SI)
    'PRO-152', # Ch. Aditya Srinivas (SI)
    'PRO-125', # J. Kalpana (SI)
    'PRO-142', # K. Venkata Rao (SI)
    'PRO-150', # D. Rama Koti Naik (SI)
    'PRO-151', # P.V. Naidu (PMO) (SI)
    'PRO-155', # S.S. Siva Rama Sastry (SI)
    'PRO-157', # IR Koteswara Rao (SI)
    'PRO-158'  # B. Rani (ASI)
]
RANK_HIERARCHY_ORDER = ['SP', 'Addl. SP', 'DSP', 'AAO', 'CI', 'SI', 'ASI', 'Common Task', 'Unassigned', 'Other']

def get_officer_hierarchy_weight(pro_id, rank):
    pid = (pro_id or '').strip().upper()
    if pid in OFFICER_HIERARCHY_IDS:
        return OFFICER_HIERARCHY_IDS.index(pid)
    rnk = rank or 'Other'
    if rnk in RANK_HIERARCHY_ORDER:
        return 100 + RANK_HIERARCHY_ORDER.index(rnk)
    return 199

def parse_pro_id(val):
    cleaned = clean_text(val)
    if cleaned.upper() == 'ALL':
        return {'id': 'ALL', 'name': 'ALL (Common Task)', 'rank': 'Common Task', 'label': 'ALL'}
    if cleaned in ['N/A', 'Unassigned', '']:
        return {'id': 'N/A', 'name': 'Unassigned', 'rank': 'Unassigned', 'label': 'Unassigned'}
    
    m_with_rank = re.match(r'^(PRO-\d+)\s*[-:]\s*(.+)\s*\(([^()]+)\)$', cleaned)
    if m_with_rank:
        pro_id = m_with_rank.group(1).strip()
        name = m_with_rank.group(2).trim() if hasattr(m_with_rank.group(2), 'trim') else m_with_rank.group(2).strip()
        rank = m_with_rank.group(3).strip()
        label = f"{pro_id} - {name} ({rank})"
        return {'id': pro_id, 'name': name, 'rank': rank, 'label': label}

    m_no_rank = re.match(r'^(PRO-\d+)\s*[-:]\s*(.+)$', cleaned)
    if m_no_rank:
        pro_id = m_no_rank.group(1).strip()
        name = m_no_rank.group(2).strip()
        return {'id': pro_id, 'name': name, 'rank': 'Other', 'label': f"{pro_id} - {name}"}
    
    return {'id': 'PRO-GEN', 'name': cleaned, 'rank': 'Other', 'label': cleaned}

def parse_prj_id(val):
    cleaned = clean_text(val)
    if cleaned in ['N/A', ''] or cleaned.lower() == 'miscellaneous':
        return {'id': 'PRJ-037', 'name': 'Miscellaneous', 'label': 'PRJ-037 - Miscellaneous'}
    if cleaned.upper() == 'ALL':
        return {'id': 'PRJ-038', 'name': 'ALL', 'label': 'PRJ-038 - ALL'}
    
    m = re.match(r'^(PRJ-\d+)\s*[-:]\s*(.+)$', cleaned)
    if m:
        prj_id = m.group(1).strip()
        name = m.group(2).strip()
        label = f"{prj_id} - {name}"
        return {'id': prj_id, 'name': name, 'label': label}
    
    return {'id': 'PRJ-037', 'name': cleaned, 'label': f"PRJ-037 - {cleaned}"}

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
    stage_lower = re.sub(r'\s+', ' ', stage.lower())

    # Excluded categories (NOT required to show in brief / active KPIs)
    if (stage_lower in ['correspondence closed', 'cancelled', 'canceled', 'returned to concerned department'] or
        'correspondence closed' in stage_lower or
        'returned to concerned department' in stage_lower):
        return 'Closed / Excluded'

    # Category 2: "Completed" is separate category
    if stage_lower == 'completed':
        return 'Completed'

    # Category 3: "N/A" is separate category because stage is not updated
    if stage_lower in ['n/a', 'na', 'not available']:
        return 'Stage N/A'

    # Category 1: In-Progress categories (Draft Prepared, Sent for Approval, File Sent through eOffice,
    # File with IC/DSP/DIG/IGP/DGP/Govt. of AP, Returned for Corrections, Pending for Approval,
    # Pending for DGP Approval, Approved, In Progress, Work Completed - Pending for Approval,
    # File Dispatched, Mail Sent, Letter Sent, Proforma Sent, On Hold, Overdue, etc.)
    return 'In-Progress'

def calculate_analytics_from_records(records):
    today = datetime.now().date()
    now_str = datetime.now().strftime("%Y-%m-%d %I:%M:%S %p")
    today_display = datetime.now().strftime("%d-%m-%Y | %I:%M %p")

    for r in records:
        recv_raw = r.get('Date & Time Received in Office') or r.get('Date & Time Received to office') or r.get('Date Received in Office') or r.get('Date Received to office') or r.get('Date Received') or ''
        r['Date Received to office'] = recv_raw
        r['Date & Time Received in Office'] = recv_raw
        d_recv = parse_indian_date(recv_raw)
        
        orig_raw = r.get('Original Date of Letter/Mail') or r.get('Original Date') or ''
        r['Original Date of Letter/Mail'] = orig_raw
        d_orig = parse_indian_date(orig_raw)

        close_raw = r.get('Final Closure Date') or ''
        disp_raw = r.get('File Dispatched Date') or r.get('Dispatched Date') or ''
        r['Dispatched Date'] = disp_raw
        r['File Dispatched Date'] = disp_raw
        d_close = parse_indian_date(close_raw) or parse_indian_date(disp_raw)

        effective_start = d_recv or d_orig

        stage_raw = r.get('Current Stage') or ''
        r['status_group'] = categorize_stage(stage_raw)

        off_info = parse_pro_id(r.get('Concerned Officer'))
        prj_info = parse_prj_id(r.get('Project'))

        r['officer_id'] = off_info['id']
        r['officer_name'] = off_info['name']
        r['officer_rank'] = off_info['rank']
        r['officer_label'] = off_info['label']

        r['project_id'] = prj_info['id']
        r['project_name'] = prj_info['name']
        r['project_label'] = prj_info['label']

        r['Source'] = clean_text(r.get('Source / Instructing Officer') or r.get('Source') or '')
        r['Source / Instructing Officer'] = r['Source']
        r['Received Through'] = clean_text(r.get('Received Through') or r.get('Received Through ') or '')

        dept_raw = clean_text(r.get('Initiating Department / Wing') or r.get('Initiating Department /Wing') or r.get('Received From Department /Wing') or r.get('Received From Department') or r.get('Received From Department / Wing') or '')
        r['Received From Department /Wing'] = dept_raw
        r['Initiating Department / Wing'] = dept_raw

        desig_raw = clean_text(r.get('Initiating Officer Designation') or r.get('Initiating Officer designation') or r.get('Received From Officer designation ') or r.get('Received From Officer designation') or r.get('Received From Officer Designation') or '')
        r['Received From Officer designation '] = desig_raw
        r['Initiating Officer Designation'] = desig_raw

        nature_raw = clean_text(r.get('Nature of Request') or r.get('Nature of request') or r.get('Nature') or '')
        r['nature_of_request'] = nature_raw if nature_raw != 'N/A' else 'General Action'
        r['Nature of Request'] = r['nature_of_request']

        due_date_raw = clean_text(r.get('Due date/Event date') or r.get('Due Date / Event Date') or r.get('Due date / Event date') or r.get('Due Date') or '')
        r['due_date_raw'] = due_date_raw if due_date_raw != 'N/A' else ''
        r['Due date/Event date'] = r['due_date_raw']
        d_due = parse_indian_date(r['due_date_raw'])
        r['due_date_iso'] = d_due.strftime("%Y-%m-%d") if d_due else ''

        followup_raw = clean_text(r.get('Further Fallow up Required with Concerned Department ') or r.get('Further Fallow up Required with Concerned Department') or r.get('Further Follow up Required with Concerned Department ') or r.get('Further Follow up Required with Concerned Department') or r.get('Further Follow-up Required') or '')
        r['Further Fallow up Required with Concerned Department '] = followup_raw
        r['Further Follow-up Required'] = followup_raw

        last_followup_raw = clean_text(r.get('Last Follow-up Date') or r.get('Last Followup Date') or '')
        r['Last Follow-up Date'] = last_followup_raw

        delay_days_raw = clean_text(r.get('Delay days') or r.get('Delay Days') or '')
        r['Delay days'] = delay_days_raw

        decision_raw = clean_text(r.get('Final Decision ') or r.get('Final Decision') or '')
        r['Final Decision '] = decision_raw

        if d_due:
            delta_days = (d_due - today).days
            if delta_days < 0:
                timing_status = 'Past'
                timing_label = f"{-delta_days}d ago"
            elif delta_days == 0:
                timing_status = 'Today'
                timing_label = "Today ⭐"
            elif delta_days == 1:
                timing_status = 'Tomorrow'
                timing_label = "Tomorrow ⏳"
            else:
                timing_status = 'Upcoming'
                timing_label = f"In {delta_days}d"
        else:
            delta_days = 9999
            timing_status = 'No Date'
            timing_label = '—'

        r['event_timing_status'] = timing_status
        r['event_timing_label'] = timing_label
        r['event_delta_days'] = delta_days

        remarks_raw = clean_text(r.get('Remarks/other Information') or r.get('Remarks/other Information ') or r.get('Remarks') or r.get('Remarks / Information') or '')
        r['remarks'] = remarks_raw if remarks_raw != 'N/A' else ''
        r['Remarks'] = r['remarks']
        r['Remarks/other Information '] = r['remarks']
        
        url_match = re.search(r'(https?://[^\s]+)', r['remarks'])
        r['remarks_link'] = url_match.group(1) if url_match else ''

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
        elif r['status_group'] in ['In-Progress', 'Stage N/A']:
            if effective_start:
                age_days = max(0, (today - effective_start).days)
                r['calculated_aging_days'] = age_days
                r['display_time_metric'] = f"{age_days}d (Age)"
            else:
                r['calculated_aging_days'] = 0
                r['display_time_metric'] = "—"
            r['calculated_tat_days'] = 0
        else:
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

        # ================= 4 DEDICATED VIP PILLARS DETECTION =================
        vip_reasons = []
        is_vip_dept = any(d in dept_raw.lower() for d in VIP_DEPTS)
        is_vip_stage = any(s in stage_raw.lower() for s in VIP_STAGES)
        is_vip_followup = followup_raw.lower() in ['yes', 'y', 'true']
        is_vip_desig = any(dg in desig_raw.lower() for dg in VIP_DESIGS)

        if is_vip_dept:
            vip_reasons.append(f"🏛️ VIP Dept: {dept_raw}")
        if is_vip_stage:
            vip_reasons.append(f"📜 Govt/DGP Stage: {stage_raw}")
        if is_vip_followup:
            vip_reasons.append("🚨 Urgent Follow-Up Required (Yes)")
        if is_vip_desig:
            vip_reasons.append(f"🎖️ VIP Officer: {desig_raw}")

        is_vip = (is_vip_dept or is_vip_stage or is_vip_followup or is_vip_desig)
        r['is_vip'] = is_vip
        r['is_vip_dept'] = is_vip_dept
        r['is_vip_stage'] = is_vip_stage
        r['is_vip_followup'] = is_vip_followup
        r['is_vip_desig'] = is_vip_desig
        r['vip_reasons'] = vip_reasons

        orig_priority = (r.get('Priority') or '').strip()
        if orig_priority not in ['Critical', 'High', 'Normal']:
            orig_priority = ''

        if is_vip:
            if (orig_priority == 'Critical' or 
                is_vip_followup or 
                'dgp' in dept_raw.lower() or 
                'govt' in stage_raw.lower() or 
                'dgp' in stage_raw.lower() or
                'highcourt' in dept_raw.lower() or 'high court' in dept_raw.lower()):
                resolved_priority = 'Critical'
            else:
                resolved_priority = 'High'
        else:
            resolved_priority = orig_priority if orig_priority in ['Critical', 'High'] else 'Normal'

        r['Priority'] = resolved_priority

    # Extract Scheduled Events / Meetings / Workshops
    event_keywords = ['meeting', 'conference', 'vc', 'video', 'workshop', 'seminar', 'training', 'webinar', 'review', 'session', 'course']
    scheduled_events = []
    for r in records:
        nature_lower = (r.get('nature_of_request') or '').lower()
        is_event = any(kw in nature_lower for kw in event_keywords)
        
        if is_event:
            scheduled_events.append({
                'id': r.get('ID'),
                'project': r.get('project_label'),
                'prj_id': r.get('project_id'),
                'prj_name': r.get('project_name'),
                'officer': r.get('officer_label'),
                'officer_id': r.get('officer_id'),
                'officer_name': r.get('officer_name'),
                'officer_rank': r.get('officer_rank'),
                'subject': r.get('Subject / Work Description'),
                'due_date_raw': r.get('due_date_raw'),
                'due_date_iso': r.get('due_date_iso'),
                'nature_of_request': r.get('nature_of_request'),
                'remarks': r.get('remarks'),
                'remarks_link': r.get('remarks_link'),
                'status_group': r.get('status_group'),
                'current_stage': r.get('Current Stage'),
                'priority': r.get('Priority'),
                'is_vip': r.get('is_vip'),
                'vip_reasons': r.get('vip_reasons'),
                'event_timing_status': r.get('event_timing_status'),
                'event_timing_label': r.get('event_timing_label'),
                'event_delta_days': r.get('event_delta_days')
            })

    scheduled_events.sort(key=lambda x: (
        0 if 0 <= x['event_delta_days'] < 9999 else (1 if x['event_delta_days'] == 9999 else 2),
        x['event_delta_days'] if 0 <= x['event_delta_days'] < 9999 else (-x['event_delta_days'] if x['event_delta_days'] < 0 else 0)
    ))

    upcoming_events_count = sum(1 for e in scheduled_events if e['status_group'] != 'Closed / Excluded' and (e['event_delta_days'] >= 0 or e['status_group'] != 'Completed'))

    # Overall KPIs (Only the 3 reportable categories: In-Progress, Completed, Stage N/A)
    reportable_records = [r for r in records if r['status_group'] != 'Closed / Excluded']
    total_records = len(reportable_records)
    completed_count = sum(1 for r in reportable_records if r['status_group'] == 'Completed')
    inprogress_count = sum(1 for r in reportable_records if r['status_group'] == 'In-Progress')
    pending_count = 0
    onhold_closed_count = sum(1 for r in reportable_records if r['status_group'] == 'Stage N/A')
    excluded_count = sum(1 for r in records if r['status_group'] == 'Closed / Excluded')
    critical_count = sum(1 for r in reportable_records if r['Priority'] == 'Critical')
    high_count = sum(1 for r in reportable_records if r['Priority'] == 'High')
    vip_total_count = sum(1 for r in reportable_records if r.get('is_vip'))

    # Dedicated VIP Analytics
    vip_records = [r for r in reportable_records if r.get('is_vip')]
    vip_dept_records = [r for r in reportable_records if r.get('is_vip_dept')]
    vip_stage_records = [r for r in reportable_records if r.get('is_vip_stage')]
    vip_followup_records = [r for r in reportable_records if r.get('is_vip_followup')]
    vip_desig_records = [r for r in reportable_records if r.get('is_vip_desig')]
    vip_active_records = [r for r in vip_records if r.get('status_group') != 'Completed']
    vip_completed_records = [r for r in vip_records if r.get('status_group') == 'Completed']

    tat_list = [r['calculated_tat_days'] for r in reportable_records if r['status_group'] == 'Completed' and r['calculated_tat_days'] > 0]
    avg_tat = round(sum(tat_list) / len(tat_list), 1) if tat_list else 0

    aging_buckets = {"< 3 Days": 0, "4 - 7 Days": 0, "8 - 15 Days": 0, "15+ Days": 0}
    for r in reportable_records:
        if r['status_group'] in ['In-Progress', 'Stage N/A'] and r['display_time_metric'] != '—':
            age = r['calculated_aging_days']
            if age <= 3:
                aging_buckets["< 3 Days"] += 1
            elif age <= 7:
                aging_buckets["4 - 7 Days"] += 1
            elif age <= 15:
                aging_buckets["8 - 15 Days"] += 1
            else:
                aging_buckets["15+ Days"] += 1

    # Group by Officers (Exclude Closed / Excluded from officer brief & counts)
    officers_map = {}
    for r in reportable_records:
        off_label = r['officer_label']
        if off_label not in officers_map:
            officers_map[off_label] = {
                'officer': off_label,
                'pro_id': r['officer_id'],
                'name': r['officer_name'],
                'rank': r['officer_rank'],
                'total': 0, 'completed': 0, 'inprogress': 0, 'pending': 0, 'onhold': 0, 'critical': 0, 'vip_count': 0,
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
            'is_vip': r.get('is_vip'),
            'vip_reasons': r.get('vip_reasons'),
            'recv_date': r.get('Date Received to office'),
            'due_date': r.get('due_date_raw'),
            'nature': r.get('nature_of_request'),
            'remarks': r.get('remarks')
        }

        if r['status_group'] == 'Completed':
            om['completed'] += 1
            om['completed_files'].append(file_entry)
        elif r['status_group'] == 'In-Progress':
            om['inprogress'] += 1
            om['inprogress_files'].append(file_entry)
        else:
            om['onhold'] += 1
            om['na_stage_files'].append(file_entry)

        if r['Priority'] in ['Critical', 'High']:
            om['critical'] += 1
        
        if r.get('is_vip'):
            om['vip_count'] += 1
        
        if r['project_label'] not in ['N/A', '']:
            om['projects'].add(r['project_label'])

    officer_analytics = []
    for off_label, data in sorted(
        officers_map.items(),
        key=lambda x: (
            get_officer_hierarchy_weight(x[1]['pro_id'], x[1]['rank']),
            -(x[1]['inprogress'] + x[1]['onhold']),
            -x[1]['total']
        )
    ):
        active_total = data['inprogress'] + data['onhold']
        officer_analytics.append({
            'officer': off_label,
            'pro_id': data['pro_id'],
            'name': data['name'],
            'rank': data['rank'],
            'total': data['total'],
            'active_total': active_total,
            'completed': data['completed'],
            'inprogress': data['inprogress'],
            'pending': 0,
            'onhold': data['onhold'],
            'critical': data['critical'],
            'vip_count': data['vip_count'],
            'projects': sorted(list(data['projects'])),
            'pending_files': data['pending_files'],
            'inprogress_files': data['inprogress_files'],
            'na_stage_files': data['na_stage_files'],
            'completed_files': data['completed_files']
        })

    # Designation-Wise Summary
    rank_order = RANK_HIERARCHY_ORDER
    rank_summary = {}
    for o in officer_analytics:
        rnk = o['rank'] if o['rank'] and o['rank'] != 'Other' else ('Unassigned' if o['pro_id'] == 'N/A' else ('Common Task' if o['pro_id'] == 'ALL' else 'Other'))
        if rnk not in rank_summary:
            rank_summary[rnk] = {
                'rank': rnk,
                'officers_count': 0,
                'total': 0,
                'active_total': 0,
                'pending': 0,
                'inprogress': 0,
                'onhold': 0,
                'completed': 0,
                'critical': 0,
                'vip_count': 0,
                'officers': []
            }
        rs = rank_summary[rnk]
        rs['officers_count'] += 1
        rs['total'] += o['total']
        rs['active_total'] += o['active_total']
        rs['pending'] += o['pending']
        rs['inprogress'] += o['inprogress']
        rs['onhold'] += o['onhold']
        rs['completed'] += o['completed']
        rs['critical'] += o['critical']
        rs['vip_count'] += o['vip_count']
        rs['officers'].append(o)

    designation_analytics = []
    for rnk in rank_order:
        if rnk in rank_summary:
            designation_analytics.append(rank_summary[rnk])
    for rnk, data in rank_summary.items():
        if rnk not in rank_order:
            designation_analytics.append(data)

    # Group by Projects (Exclude Closed / Excluded)
    projects_map = {}
    for r in reportable_records:
        prj_label = r['project_label']
        if prj_label not in projects_map:
            projects_map[prj_label] = {
                'project': prj_label,
                'prj_id': r['project_id'],
                'name': r['project_name'],
                'total': 0, 'completed': 0, 'inprogress': 0, 'pending': 0, 'onhold': 0, 'critical': 0, 'vip_count': 0,
                'officers': set()
            }
        pm = projects_map[prj_label]
        pm['total'] += 1
        if r['status_group'] == 'Completed':
            pm['completed'] += 1
        elif r['status_group'] == 'In-Progress':
            pm['inprogress'] += 1
        else:
            pm['onhold'] += 1

        if r['Priority'] in ['Critical', 'High']:
            pm['critical'] += 1
        
        if r.get('is_vip'):
            pm['vip_count'] += 1

        if r['officer_label'] not in ['N/A', '']:
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
            'pending': 0,
            'onhold': data['onhold'],
            'critical': data['critical'],
            'vip_count': data['vip_count'],
            'officers': sorted(list(data['officers']))
        })

    # Group by Channels
    channels_map = {}
    for r in reportable_records:
        ch = (r.get('Received Through') or r.get('Received Through ') or '').strip()
        if not ch or ch in ['N/A', 'None']:
            ch = 'Other'
        channels_map[ch] = channels_map.get(ch, 0) + 1

    channel_analytics = [{'channel': k, 'count': v} for k, v in sorted(channels_map.items(), key=lambda x: x[1], reverse=True)]

    # Formulate Short Numbers-Only WhatsApp Text (Only 3 Categories: In-Progress, Completed, N/A)
    reportable_officers = [o for o in officer_analytics if o['total'] > 0]
    wa_short_lines = [
        "🏛️ *AP POLICE TECHNICAL SERVICES (PCS&S)*",
        "⚡ *DAILY BRIEF - CORRESPONDENCE STATUS*",
        f"📅 Date: {today_display}",
        "",
        f"📊 *Total: {total_records}* | 🔵 *In-Progress: {inprogress_count}* | 🟢 *Completed: {completed_count}* | ⚪ *N/A: {onhold_closed_count}*",
        "",
        "───────────────────────────────",
        "👮 *OFFICER WORKLOAD Current stage:* [Total = In-Prog / Completed / N/A]",
        "───────────────────────────────",
        ""
    ]

    individual_reportable_officers = [o for o in reportable_officers if o['pro_id'].upper() != 'ALL' and o['officer'].upper() != 'ALL']
    for idx, o in enumerate(individual_reportable_officers, 1):
        wa_short_lines.append(f"{idx}. {o['officer']}: *{o['total']}* ({o['inprogress']} / {o['completed']} / {o['onhold']})")

    common_reportable_files = [r for r in reportable_records if (r.get('officer_id', '').upper() == 'ALL' or r.get('officer_label', '').upper() == 'ALL')]
    if common_reportable_files:
        wa_short_lines.append("")
        wa_short_lines.append("───────────────────────────────")
        wa_short_lines.append("📢 *COMMON TASKS (Assigned to ALL Officers):*")
        wa_short_lines.append("───────────────────────────────")
        for idx, f in enumerate(common_reportable_files, 1):
            vip_tag = " ⭐[VIP]" if f.get('is_vip') else ""
            wa_short_lines.append(f"{idx}. *{f['ID']}*{vip_tag} [{f.get('project_label', 'PRJ-038 - ALL')}] — *{f.get('Current Stage') or 'In-Progress'}*")
            wa_short_lines.append(f"   • {(f.get('Subject / Work Description') or '')[:90]}")

    wa_short_lines.append("")
    wa_short_lines.append("───────────────────────────────")
    wa_short_lines.append("📝 *Note:* Kindly update the current stages in the Google Sheet if there are any changes.")
    wa_short_lines.append("🔗 *Update Status in Google Sheet:*")
    wa_short_lines.append(GOOGLE_SHEET_URL)
    wa_short_lines.append("───────────────────────────────")
    wa_short_lines.append("_Generated from AP Police TS Executive Command Portal_")
    whatsapp_short_text = "\n".join(wa_short_lines)
    whatsapp_text = whatsapp_short_text

    # Formulate Format 3: Exclusive Events & Meeting Schedules WhatsApp Text
    events_lines = [
        "🏛️ *AP POLICE TECHNICAL SERVICES (PCS&S)*",
        "📅 *UPCOMING MEETINGS, WORKSHOPS & EVENT SCHEDULE*",
        f"📆 Date: {today_display}",
        "",
        "───────────────────────────────",
        "📌 *SCHEDULED EVENTS & MEETINGS:*",
        "───────────────────────────────"
    ]

    active_events = [e for e in scheduled_events if e['status_group'] != 'Closed / Excluded' and (e['event_delta_days'] >= 0 or e['status_group'] != 'Completed')]
    if not active_events:
        events_lines.append("")
        events_lines.append("• No upcoming scheduled meetings/workshops at present.")
    else:
        for idx, ev in enumerate(active_events, 1):
            date_str = ev['due_date_raw'] if ev['due_date_raw'] else 'Date TBA'
            timing = f" ({ev['event_timing_label']})" if ev['event_timing_label'] != '—' else ''
            nature = ev['nature_of_request']
            icon = "💻" if any(k in nature.lower() for k in ['vc', 'video', 'conference']) else ("🎓" if any(k in nature.lower() for k in ['workshop', 'seminar', 'training']) else "🏢")
            is_postponed = (ev.get('current_stage') or '').lower().strip() == 'postponed' or 'postponed' in (ev.get('remarks') or '').lower()
            status_tag = " ⏸️ *[POSTPONED]*" if is_postponed else ""
            
            events_lines.append("")
            events_lines.append(f"{idx}. {icon} *{nature}*{status_tag}")
            events_lines.append(f"   • *File ID:* {ev['id']} [{ev['project']}]")
            events_lines.append(f"   • *Date / Time:* 📅 {date_str}{timing}")
            events_lines.append(f"   • *Officer:* {ev['officer']}")
            if is_postponed:
                events_lines.append(f"   • *Status:* ⏸️ *Postponed (Awaiting Rescheduled Date)*")
            events_lines.append(f"   • *Subject:* {ev['subject']}")
            if ev['remarks']:
                events_lines.append(f"   • *Venue / VC Link:* {ev['remarks']}")

    events_lines.append("")
    events_lines.append("───────────────────────────────")
    events_lines.append("🔗 *Update Status in Google Sheet:*")
    events_lines.append(GOOGLE_SHEET_URL)
    events_lines.append("───────────────────────────────")
    events_lines.append("_Generated from AP Police TS Executive Command Portal_")
    whatsapp_events_text = "\n".join(events_lines)

    # Formulate Format 4: VIP Short Brief WhatsApp Text
    vip_followup_active = [r for r in vip_active_records if r.get('is_vip_followup')]
    vip_wa_lines = [
        "🏛️ *AP POLICE TECHNICAL SERVICES (PCS&S)*",
        "⭐ *VIP & HIGH PRIORITY CORRESPONDENCE SHORT BRIEF*",
        f"📅 Date: {today_display}",
        "",
        f"📊 *VIP Overview:* Total: *{len(vip_records)}* | 🔵 In-Prog / N/A: *{len(vip_active_records)}* | 🟢 Completed: *{len(vip_completed_records)}* | ⚠️ Follow-up: *{len(vip_followup_active)}*",
        "",
        "───────────────────────────────",
        "📌 *ACTIVE VIP FILES:*",
        "───────────────────────────────"
    ]

    if not vip_active_records:
        vip_wa_lines.append("")
        vip_wa_lines.append("✅ No active VIP files pending resolution.")
    else:
        for idx, r in enumerate(vip_active_records, 1):
            followup_tag = " ⚠️" if r.get('is_vip_followup') else ""
            vip_wa_lines.append(f"{idx}. *{r['ID']}*{followup_tag} [{r['project_label']}] — {r['Current Stage']} ({r['display_time_metric']}) — {r['officer_label']}")

    vip_wa_lines.append("")
    vip_wa_lines.append("───────────────────────────────")
    vip_wa_lines.append("📝 *Note:* Kindly update the current stages in the Google Sheet if there are any changes.")
    vip_wa_lines.append("🔗 *Update Status in Google Sheet:*")
    vip_wa_lines.append(GOOGLE_SHEET_URL)
    vip_wa_lines.append("───────────────────────────────")
    vip_wa_lines.append("_Generated from AP Police TS Executive Command Portal_")
    whatsapp_vip_text = "\n".join(vip_wa_lines)

    vip_analytics_payload = {
        'total_count': len(vip_records),
        'dept_count': len(vip_dept_records),
        'stage_count': len(vip_stage_records),
        'followup_count': len(vip_followup_records),
        'desig_count': len(vip_desig_records),
        'active_count': len(vip_active_records),
        'completed_count': len(vip_completed_records),
        'records': vip_records,
        'whatsapp_vip_text': whatsapp_vip_text
    }

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
            'excluded_closed': excluded_count,
            'critical': critical_count,
            'high': high_count,
            'vip_count': vip_total_count,
            'avg_tat_days': avg_tat
        },
        'aging_buckets': aging_buckets,
        'officers': officer_analytics,
        'designations': designation_analytics,
        'projects': project_analytics,
        'channels': channel_analytics,
        'scheduled_events': scheduled_events,
        'upcoming_events_count': upcoming_events_count,
        'vip_analytics': vip_analytics_payload,
        'records': records,
        'pending_abstract': {
            'total_pending': pending_count,
            'total_inprogress': inprogress_count,
            'total_na_stage': onhold_closed_count,
            'vip_count': vip_total_count,
            'officers_active': reportable_officers,
            'designations': designation_analytics,
            'whatsapp_text': whatsapp_text,
            'whatsapp_short_text': whatsapp_short_text,
            'whatsapp_events_text': whatsapp_events_text,
            'whatsapp_vip_text': whatsapp_vip_text
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
                    if not h:
                        continue
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
        print(f" Dedicated VIP Command Center Active (4 Pillars)")
        print(f"================================================================")
        httpd.serve_forever()

if __name__ == '__main__':
    start_server()
