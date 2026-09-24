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
    'secretary to the govt of ap', 'secretary to the govt. of ap', 'secretary to the govt.of ap', 'secreatry to govt of india', 'secretary to govt',
    'principal secretary', 'chief secretary', 'cs to the govt', 'cs to govt',
    'registar highcourt', 'registrar it high court', 'registrar high court', 'registrar',
    'adg', 'addl.director general of police', 'additional director general of police', 'addl.director general',
    'igp', 'inspector general of police', 'deputy inspector general of police', 'deputy inspector general',
    'deputy director'
]

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
        return {'id': 'N/A', 'name': 'Unassigned', 'rank': 'Unassigned', 'label': 'Unassigned'}
    
    # Regex match for PRO-xxx - Name (Rank)
    m = re.match(r'^(PRO-\d+)\s*[-:]\s*(.+?)(?:\s*\((.+?)\))?$', cleaned)
    if m:
        pro_id = m.group(1).strip()
        name = m.group(2).strip()
        rank = (m.group(3) or 'Other').strip()
        label = f"{pro_id} - {name}" + (f" ({rank})" if rank else "")
        return {'id': pro_id, 'name': name, 'rank': rank, 'label': label}
    
    return {'id': 'PRO-GEN', 'name': cleaned, 'rank': 'Other', 'label': cleaned}

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

    # 1. Explicitly Completed / Dispatched / Acknowledgement
    if stage_lower in ['completed', 'file dispatched', 'acknowledgement filed']:
        return 'Completed'

    # 2. Explicitly On Hold / Cancelled / Closed / Inactive
    if stage_lower in ['on hold', 'cancelled', 'correspondence closed', 'inactive']:
        return 'On Hold / Closed'

    # 3. Pending (Any stage mentioning "pending", or blank / N/A which represents intake awaiting allocation/action)
    if 'pending' in stage_lower or stage_lower in ['n/a', 'none', 'null', '']:
        return 'Pending'

    # 4. Remaining all are In-Progress
    return 'In-Progress'

def calculate_analytics_from_records(records):
    today = datetime.now().date()
    now_str = datetime.now().strftime("%Y-%m-%d %I:%M:%S %p")
    today_display = datetime.now().strftime("%d-%m-%Y | %I:%M %p")

    for r in records:
        # Resolve Date fields
        recv_raw = r.get('Date & Time Received to office') or r.get('Date Received to office') or r.get('Date Received') or ''
        r['Date Received to office'] = recv_raw # normalize standard key
        d_recv = parse_indian_date(recv_raw)
        
        orig_raw = r.get('Original Date of Letter/Mail') or r.get('Original Date') or ''
        r['Original Date of Letter/Mail'] = orig_raw
        d_orig = parse_indian_date(orig_raw)

        close_raw = r.get('Final Closure Date') or ''
        disp_raw = r.get('Dispatched Date') or ''
        d_close = parse_indian_date(close_raw) or parse_indian_date(disp_raw)

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

        # Parse Source and Channel
        r['Source'] = clean_text(r.get('Source') or '')
        r['Received Through'] = clean_text(r.get('Received Through') or r.get('Received Through ') or '')

        # Parse Sender Department & Officer Designation
        dept_raw = clean_text(r.get('Received From Department /Wing') or r.get('Received From Department') or r.get('Received From Department / Wing') or '')
        r['Received From Department /Wing'] = dept_raw

        desig_raw = clean_text(r.get('Received From Officer designation ') or r.get('Received From Officer designation') or r.get('Received From Officer Designation') or '')
        r['Received From Officer designation '] = desig_raw

        # Parse Nature of Request, Due/Event Date, and Remarks
        nature_raw = clean_text(r.get('Nature of Request') or r.get('Nature of request') or r.get('Nature') or '')
        r['nature_of_request'] = nature_raw if nature_raw != 'N/A' else 'General Action'
        r['Nature of Request'] = r['nature_of_request']

        due_date_raw = clean_text(r.get('Due date/Event date') or r.get('Due Date / Event Date') or r.get('Due date / Event date') or r.get('Due Date') or '')
        r['due_date_raw'] = due_date_raw if due_date_raw != 'N/A' else ''
        r['Due date/Event date'] = r['due_date_raw']
        d_due = parse_indian_date(r['due_date_raw'])
        r['due_date_iso'] = d_due.strftime("%Y-%m-%d") if d_due else ''

        # Parse Follow-up & Delay Days
        followup_raw = clean_text(r.get('Further Fallow up Required with Concerned Department ') or r.get('Further Fallow up Required with Concerned Department') or r.get('Further Follow up Required with Concerned Department ') or r.get('Further Follow up Required with Concerned Department') or r.get('Further Follow-up Required') or '')
        r['Further Fallow up Required with Concerned Department '] = followup_raw
        r['Further Follow-up Required'] = followup_raw

        last_followup_raw = clean_text(r.get('Last Follow-up Date') or r.get('Last Followup Date') or '')
        r['Last Follow-up Date'] = last_followup_raw

        delay_days_raw = clean_text(r.get('Delay days') or r.get('Delay Days') or '')
        r['Delay days'] = delay_days_raw

        decision_raw = clean_text(r.get('Final Decision ') or r.get('Final Decision') or '')
        r['Final Decision '] = decision_raw

        # Calculate event timing / countdown
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
        
        # Detect URLs in remarks for direct meeting access
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

        # ================= INTELLIGENT VIP / CRITICAL PRIORITY DETECTION =================
        vip_reasons = []
        if any(d in dept_raw.lower() for d in VIP_DEPTS):
            vip_reasons.append(f"VIP Dept: {dept_raw}")
        if any(s in stage_raw.lower() for s in VIP_STAGES):
            vip_reasons.append(f"VIP Stage: {stage_raw}")
        if followup_raw.lower() in ['yes', 'y', 'true']:
            vip_reasons.append("Follow-up Required with Dept")
        if any(dg in desig_raw.lower() for dg in VIP_DESIGS):
            vip_reasons.append(f"VIP Officer: {desig_raw}")

        is_vip = len(vip_reasons) > 0
        r['is_vip'] = is_vip
        r['vip_reasons'] = vip_reasons

        orig_priority = (r.get('Priority') or '').strip()
        if orig_priority not in ['Critical', 'High', 'Normal']:
            orig_priority = ''

        if is_vip:
            if (orig_priority == 'Critical' or 
                followup_raw.lower() in ['yes', 'y', 'true'] or 
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

    # Sort events: Upcoming first (0..n), then No Date, then Past (-1..-n)
    scheduled_events.sort(key=lambda x: (
        0 if 0 <= x['event_delta_days'] < 9999 else (1 if x['event_delta_days'] == 9999 else 2),
        x['event_delta_days'] if 0 <= x['event_delta_days'] < 9999 else (-x['event_delta_days'] if x['event_delta_days'] < 0 else 0)
    ))

    upcoming_events_count = sum(1 for e in scheduled_events if e['event_delta_days'] >= 0 or e['status_group'] != 'Completed')

    total_records = len(records)
    completed_count = sum(1 for r in records if r['status_group'] == 'Completed')
    inprogress_count = sum(1 for r in records if r['status_group'] == 'In-Progress')
    pending_count = sum(1 for r in records if r['status_group'] == 'Pending')
    onhold_closed_count = sum(1 for r in records if r['status_group'] == 'On Hold / Closed')
    critical_count = sum(1 for r in records if r['Priority'] == 'Critical')
    high_count = sum(1 for r in records if r['Priority'] == 'High')
    vip_total_count = sum(1 for r in records if r.get('is_vip'))

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
        elif r['status_group'] == 'Pending':
            om['pending'] += 1
            om['pending_files'].append(file_entry)
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
            'vip_count': data['vip_count'],
            'projects': sorted(list(data['projects'])),
            'pending_files': data['pending_files'],
            'inprogress_files': data['inprogress_files'],
            'na_stage_files': data['na_stage_files'],
            'completed_files': data['completed_files']
        })

    # ================= DESIGNATION / RANK-WISE HIERARCHICAL ANALYTICS =================
    rank_order = ['DSP', 'CI', 'SI', 'Unassigned', 'Other']
    rank_summary = {}
    for o in officer_analytics:
        rnk = o['rank'] if o['rank'] and o['rank'] != 'Other' else ('Unassigned' if o['pro_id'] == 'N/A' else 'Other')
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

    # Group by Projects
    projects_map = {}
    for r in records:
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
        elif r['status_group'] == 'Pending':
            pm['pending'] += 1
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
            'pending': data['pending'],
            'onhold': data['onhold'],
            'critical': data['critical'],
            'vip_count': data['vip_count'],
            'officers': sorted(list(data['officers']))
        })

    # Group by Channels
    channels_map = {}
    for r in records:
        ch = (r.get('Received Through') or r.get('Received Through ') or '').strip()
        if not ch or ch in ['N/A', 'None']:
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
        f"• ⭐ High Priority / VIP Files: {vip_total_count} Files",
        "",
        "───────────────────────────────",
        "🎖️ *WORKLOAD BY DESIGNATION / RANK:*",
        "───────────────────────────────"
    ]

    for da in designation_analytics:
        if da['active_total'] > 0 or da['total'] > 0:
            wa_lines.append(f"🔹 *{da['rank']} Rank ({da['officers_count']} Officers)*: *{da['active_total']} Active Files* ({da['pending']} Pend / {da['inprogress']} In-Prog / {da['onhold']} N/A)")

    wa_lines.append("")
    wa_lines.append("───────────────────────────────")
    wa_lines.append("👤 *OFFICER-WISE DETAILED BREAKDOWN:*")
    wa_lines.append("───────────────────────────────")

    for da in designation_analytics:
        rank_active_officers = [o for o in da['officers'] if o['active_total'] > 0]
        if not rank_active_officers:
            continue
        wa_lines.append("")
        wa_lines.append(f"👮 *[{da['rank']} RANK OFFICERS]*")
        for idx, o in enumerate(rank_active_officers, 1):
            wa_lines.append(f"{idx}. *{o['officer']}*: Total Active: {o['active_total']} ({o['pending']} Pending, {o['inprogress']} In-Prog, {o['onhold']} Stage N/A)")
            
            if o['pending_files']:
                wa_lines.append("   🟡 *Pending Action:*")
                for f in o['pending_files']:
                    vip_tag = " ⭐[VIP]" if f.get('is_vip') else ""
                    wa_lines.append(f"   - {f['id']}{vip_tag} [{f['project']}]: ({f['age']})")
            
            if o['inprogress_files']:
                wa_lines.append("   🔵 *In-Progress:*")
                for f in o['inprogress_files']:
                    vip_tag = " ⭐[VIP]" if f.get('is_vip') else ""
                    wa_lines.append(f"   - {f['id']}{vip_tag} [{f['project']}]: ({f['age']})")
            
            if o['na_stage_files']:
                wa_lines.append("   ⚪ *Stage N/A (Notice Required):*")
                for f in o['na_stage_files']:
                    vip_tag = " ⭐[VIP]" if f.get('is_vip') else ""
                    wa_lines.append(f"   - {f['id']}{vip_tag} [{f['project']}]: ({f['age']})")

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
        f"📊 *Total: {total_records}* | 🟡 *Pend: {pending_count}* | 🔵 *In-Prog: {inprogress_count}* | ⚪ *N/A: {onhold_closed_count}* | 🟢 *Done: {completed_count}* | ⭐ *VIP: {vip_total_count}*",
        "",
        "───────────────────────────────",
        "🎖️ *RANK SUMMARY [Active = Pend / InProg / N/A]:*",
        "───────────────────────────────"
    ]

    for da in designation_analytics:
        if da['active_total'] > 0 or da['total'] > 0:
            wa_short_lines.append(f"• *{da['rank']}*: *{da['active_total']}* ({da['pending']} / {da['inprogress']} / {da['onhold']})")

    wa_short_lines.append("")
    wa_short_lines.append("───────────────────────────────")
    wa_short_lines.append("👮 *OFFICER WORKLOAD:*")
    wa_short_lines.append("───────────────────────────────")

    for idx, o in enumerate(active_officers, 1):
        wa_short_lines.append(f"{idx}. {o['officer']}: *{o['active_total']}* ({o['pending']} / {o['inprogress']} / {o['onhold']})")

    wa_short_lines.append("")
    wa_short_lines.append("───────────────────────────────")
    wa_short_lines.append("🔗 *Update Status in Google Sheet:*")
    wa_short_lines.append(GOOGLE_SHEET_URL)
    wa_short_lines.append("───────────────────────────────")
    wa_short_lines.append("_Generated from AP Police TS Executive Command Portal_")
    whatsapp_short_text = "\n".join(wa_short_lines)

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

    active_events = [e for e in scheduled_events if e['event_delta_days'] >= 0 or e['status_group'] != 'Completed']
    if not active_events:
        events_lines.append("")
        events_lines.append("• No upcoming scheduled meetings/workshops at present.")
    else:
        for idx, ev in enumerate(active_events, 1):
            date_str = ev['due_date_raw'] if ev['due_date_raw'] else 'Date TBA'
            timing = f" ({ev['event_timing_label']})" if ev['event_timing_label'] != '—' else ''
            nature = ev['nature_of_request']
            icon = "💻" if any(k in nature.lower() for k in ['vc', 'video', 'conference']) else ("🎓" if any(k in nature.lower() for k in ['workshop', 'seminar', 'training']) else "🏢")
            
            events_lines.append("")
            events_lines.append(f"{idx}. {icon} *{nature}*")
            events_lines.append(f"   • *File ID:* {ev['id']} [{ev['project']}]")
            events_lines.append(f"   • *Date / Time:* 📅 {date_str}{timing}")
            events_lines.append(f"   • *Officer:* {ev['officer']}")
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
        'records': records,
        'pending_abstract': {
            'total_pending': pending_count,
            'total_inprogress': inprogress_count,
            'total_na_stage': onhold_closed_count,
            'vip_count': vip_total_count,
            'officers_active': active_officers,
            'designations': designation_analytics,
            'whatsapp_text': whatsapp_text,
            'whatsapp_short_text': whatsapp_short_text,
            'whatsapp_events_text': whatsapp_events_text
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
        print(f" Designation-Wise Summary Active: DSP, CI, SI, Unassigned")
        print(f"================================================================")
        httpd.serve_forever()

if __name__ == '__main__':
    start_server()
