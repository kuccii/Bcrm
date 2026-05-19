"""Bridge: Sync CRM data to call-agent leads system.

Usage:
    python sync_to_call_agent.py              # Full sync: CRM data -> leads.xlsx
    python sync_to_call_agent.py --export     # Just export CRM with statuses
    python sync_to_call_agent.py --import     # Import call-agent outcomes back to CRM
"""
import json
import os
import sys
from datetime import datetime
from typing import List, Dict, Optional

BASE_DIR = r"C:\Users\impat\OneDrive\business-scanner"
AGENT_DIR = r"C:\Users\impat\call-agent"
CRM_DATA_FILE = os.path.join(BASE_DIR, "crm-data.js")
CRM_SYNC_FILE = os.path.join(BASE_DIR, "crm-sync.json")
LEADS_FILE = os.path.join(AGENT_DIR, "data", "leads.xlsx")
AGENT_OUTCOMES_DIR = os.path.join(AGENT_DIR, "data", "call_logs")


def parse_crm_js() -> List[Dict]:
    """Parse crm-data.js and return the CRM array."""
    if not os.path.exists(CRM_DATA_FILE):
        print(f"CRM data file not found: {CRM_DATA_FILE}")
        return []

    with open(CRM_DATA_FILE, 'r', encoding='utf-8') as f:
        content = f.read()

    # Find const CRM = [...]
    marker = 'const CRM = '
    idx = content.find(marker)
    if idx == -1:
        print("Could not find 'const CRM = ' in file")
        return []

    json_str = content[idx + len(marker):]

    # Find matching closing bracket
    depth = 0
    start = 0
    for i, ch in enumerate(json_str):
        if ch == '[':
            if depth == 0:
                start = i
            depth += 1
        elif ch == ']':
            depth -= 1
            if depth == 0:
                json_str = json_str[start:i + 1]
                break

    try:
        data = json.loads(json_str)
        print(f"Loaded {len(data)} businesses from CRM")
        return data
    except json.JSONDecodeError as e:
        print(f"JSON parse error: {e}")
        return []


def load_sync_statuses() -> Dict:
    """Load synced statuses from crm-sync.json (if exists)."""
    if os.path.exists(CRM_SYNC_FILE):
        with open(CRM_SYNC_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    return {}


def save_sync_statuses(statuses: Dict):
    """Save statuses to crm-sync.json."""
    with open(CRM_SYNC_FILE, 'w', encoding='utf-8') as f:
        json.dump(statuses, f, indent=2)
    print(f"Saved {len(statuses)} lead statuses to {CRM_SYNC_FILE}")


def import_call_outcomes():
    """Import outcomes from call-agent call logs back into the sync file."""
    if not os.path.exists(AGENT_OUTCOMES_DIR):
        print(f"Call logs directory not found: {AGENT_OUTCOMES_DIR}")
        return

    statuses = load_sync_statuses()
    imported = 0

    for filename in os.listdir(AGENT_OUTCOMES_DIR):
        if not filename.endswith('.json'):
            continue

        filepath = os.path.join(AGENT_OUTCOMES_DIR, filename)
        try:
            with open(filepath, 'r', encoding='utf-8') as f:
                log = json.load(f)
        except (json.JSONDecodeError, IOError):
            continue

        # Extract phone number
        phone = log.get('phone', '')
        if not phone:
            continue

        # Extract outcome
        outcome_data = log.get('outcome', {})
        if isinstance(outcome_data, dict):
            outcome = outcome_data.get('outcome', 'called')
            notes = outcome_data.get('notes', '')
            appointment = outcome_data.get('appointment_booked', False)
        else:
            outcome = str(outcome_data)
            notes = ''
            appointment = False

        # Map outcome to our status
        status_map = {
            'interested': 'Interested',
            'appointment_booked': 'Follow Up',
            'callback': 'Follow Up',
            'not_interested': 'Not Interested',
            'no_answer': 'Called',
            'failed': 'Called',
            'completed': 'Called'
        }
        mapped_status = status_map.get(outcome, 'Called')
        if appointment:
            mapped_status = 'Follow Up'

        # Find matching CRM lead by phone
        for lead_id, data in statuses.items():
            if data.get('phone') == phone:
                statuses[lead_id]['status'] = mapped_status
                if notes:
                    existing = statuses[lead_id].get('notes', '')
                    statuses[lead_id]['notes'] = f"{existing}\n{datetime.now().strftime('%Y-%m-%d')}: {notes}"
                statuses[lead_id]['call_log'] = filename
                imported += 1
                break

    if imported:
        save_sync_statuses(statuses)
        print(f"Imported {imported} call outcomes from agent logs")
    else:
        print("No new call outcomes to import")


def export_to_excel():
    """Export CRM leads with statuses to Excel for call-agent import."""
    crm_data = parse_crm_js()
    if not crm_data:
        return

    statuses = load_sync_statuses()

    try:
        from openpyxl import Workbook, load_workbook
    except ImportError:
        print("openpyxl not installed. Install with: pip install openpyxl")
        return

    # Build lead rows
    leads = []
    for item in crm_data:
        phone = item.get('phone', '') or item.get('mobile', '')
        if not phone:
            continue

        # Get status from sync file, or mark as new
        lead_id = str(item.get('id', ''))
        sync = statuses.get(lead_id, {})
        status = sync.get('status', 'new')

        # Only export pending leads (not already called/closed)
        if status in ['Closed Won', 'Closed Lost', 'Not Interested']:
            continue

        # Determine industry for agent's offer matching
        cat = (item.get('cat', '') or '').lower()
        industry = 'general'
        for kw, ind in [('restaurant', 'restaurant'), ('cafe', 'restaurant'),
                        ('hotel', 'hotel'), ('lodge', 'hotel'),
                        ('salon', 'salon'), ('beauty', 'salon'),
                        ('shop', 'shop'), ('store', 'shop'), ('retail', 'shop')]:
            if kw in cat:
                industry = ind
                break

        # Determine priority
        ws = item.get('ws', '')
        if ws == 'none':
            priority = 'high'
        elif ws == 'broken':
            priority = 'medium'
        else:
            priority = 'low'

        leads.append({
            'phone': phone,
            'business_name': item.get('name', ''),
            'industry': industry,
            'contact_name': '',
            'priority': priority,
            'status': 'pending' if status in ['new', 'Called', ''] else status.lower(),
            'crm_id': lead_id,
            'notes': sync.get('notes', ''),
            'lead_score': 0,
        })

    # Write to Excel — always recreate from scratch
    wb = Workbook()
    ws = wb.active
    columns = ['phone', 'business_name', 'industry', 'contact_name',
               'priority', 'status', 'crm_id', 'notes', 'lead_score',
               'appointment_date']
    ws.append(columns)

    for i, lead in enumerate(leads, start=2):
        ws.cell(row=i, column=1, value=lead['phone'])
        ws.cell(row=i, column=2, value=lead['business_name'])
        ws.cell(row=i, column=3, value=lead['industry'])
        ws.cell(row=i, column=4, value=lead['contact_name'])
        ws.cell(row=i, column=5, value=lead['priority'])
        ws.cell(row=i, column=6, value=lead['status'])
        ws.cell(row=i, column=7, value=lead['crm_id'])
        ws.cell(row=i, column=8, value=lead['notes'])
        ws.cell(row=i, column=9, value=int(lead['lead_score']))
        ws.cell(row=i, column=10, value='')

    # Remove empty trailing rows
    max_row = len(leads) + 1
    if max_row < ws.max_row:
        ws.delete_rows(max_row + 1, ws.max_row - max_row)

    wb.save(LEADS_FILE)
    print(f"Exported {len(leads)} leads to {LEADS_FILE}")
    
    # Stats
    pending = sum(1 for l in leads if l['status'] == 'pending')
    high = sum(1 for l in leads if l['priority'] == 'high')
    print(f"  Pending calls: {pending}")
    print(f"  High priority: {high}")


if __name__ == '__main__':
    if '--import' in sys.argv:
        import_call_outcomes()
    elif '--export' in sys.argv:
        export_to_excel()
    else:
        # Full sync: import outcomes first, then export to Excel
        print("=== Sync CRM Data to Call Agent ===\n")
        import_call_outcomes()
        print()
        export_to_excel()
        print("\nDone. Start the call-agent to begin calling.")
        print(f"  dashboard: cd {AGENT_DIR} && python dashboard.py")
        print(f"  webhook:   cd {AGENT_DIR} && python webhook_server.py")
