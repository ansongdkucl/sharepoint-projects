import os
import re
import sys
import argparse
import time
import requests
import json
from pathlib import Path
from openpyxl import load_workbook
from netmiko import ConnectHandler
from dotenv import load_dotenv

# ============================================================
# 1. SETUP & CONFIGURATION
# ============================================================
load_dotenv()

USERNAME = os.getenv("username")
PASSWORD = os.getenv("passwordAD")

# Webhook URL
TEAMS_WEBHOOK_URL = "https://liveuclac.webhook.office.com/webhookb2/4bf41ac1-0d61-4760-9dea-e0f6184dde8a@1faf88fe-a998-4c5b-93c9-210a11d9a5c2/IncomingWebhook/46bcd5e8d47e4de5b575fe10f189b1e1/43bfe760-7689-4d0b-96fd-46b265519580/V2OHjE1yNbt-Vi6fJNAiAdSwjM90ZnyOKY49V-zdJi0dA1"

DEFAULT_PATH = Path("/mnt/c/Users/cceadan/University College London/Estates IT - Project Documentation - Patching Schedule/90 TCR - Level 3A Patching Schedule.xlsx")

def send_teams_notification(status, message, details=None):
    if not TEAMS_WEBHOOK_URL or "placeholder" in TEAMS_WEBHOOK_URL:
        return 
    color = {"SUCCESS": "28A745", "WARNING": "FFC107", "CRITICAL": "DC3545"}.get(status, "0078D7")
    facts = []
    if details:
        for entry in details:
            facts.append({"name": f"Switch {entry['ip']}", "value": f"Port {entry['port']} -> VLAN {entry['vlan']}"})
    payload = {
        "@type": "MessageCard",
        "@context": "http://schema.org/extensions",
        "themeColor": color,
        "summary": "Aruba Config Update",
        "sections": [{
            "activityTitle": f"**Aruba Configurator: {status}**",
            "activitySubtitle": f"File: {DEFAULT_PATH.name}",
            "text": message,
            "facts": facts
        }]
    }
    try:
        requests.post(TEAMS_WEBHOOK_URL, json=payload, timeout=10)
    except Exception as e:
        print(f"[!] Teams Alert Failed: {e}")

def get_lock_owner(path):
    lock_path = path.parent / f"~${path.name}"
    if not lock_path.exists(): return "Unknown (Closed)"
    try:
        with open(lock_path, 'rb') as f:
            content = f.read().decode('latin-1', errors='ignore')
            match = re.search(r'[a-zA-Z\s]{3,}', content)
            return match.group(0).strip() if match else "a Colleague"
    except: return "a Colleague"

def run_aruba_config(switch_ip, port, vlan_id):
    device = {'device_type': 'aruba_osswitch', 'host': switch_ip, 'username': USERNAME, 'password': PASSWORD}
    commands = ["aruba-central support-mode", "conf t", f"int {port}", f"vlan access {vlan_id}"]
    try:
        with ConnectHandler(**device, conn_timeout=15) as net_connect:
            net_connect.send_config_set(commands)
            mac_out = net_connect.send_command(f"show mac-address-table int {port}")
            mac_match = re.search(r'([0-9a-fA-F]{2}[:.-]){5}[0-9a-fA-F]{2}', mac_out)
            found_mac = mac_match.group(0) if mac_match else "Unknown"
            arp_out = net_connect.send_command(f"show arp | inc {found_mac}")
            ip_match = re.search(r'(\d{1,3}\.){3}\d{1,3}', arp_out)
            found_ip = ip_match.group(0) if ip_match else "No ARP"
            return {"mac": found_mac, "ip": found_ip, "status": "Success"}
    except Exception as e:
        return {"mac": "Error", "ip": "Error", "status": str(e)[:20]}

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--safe", action="store_true", help="Confirm each change manually")
    parser.add_argument("--dry-run", action="store_true", help="Show changes without applying them")
    args = parser.parse_args()

    if not DEFAULT_PATH.exists(): 
        print(f"[!] File not found: {DEFAULT_PATH}")
        return

    # 1. Pre-run lock check
    if (DEFAULT_PATH.parent / f"~${DEFAULT_PATH.name}").exists():
        owner = get_lock_owner(DEFAULT_PATH)
        print(f"\n[!] ABORTED: File is currently open by {owner}.")
        return

    start_mtime = os.path.getmtime(DEFAULT_PATH)
    config_summary = []

    try:
        wb = load_workbook(DEFAULT_PATH, data_only=False)
        ws = wb.active
        print(f"[*] File Loaded. Sheet: {ws.title}")
    except Exception as e:
        print(f"Error: {e}"); return

    # 2. Header Mapping
    read_map = {}   # UPPERCASE
    write_map = {}  # lowercase
    for col in range(1, ws.max_column + 1):
        v2 = str(ws.cell(row=2, column=col).value or "").strip()
        v3 = str(ws.cell(row=3, column=col).value or "").strip()
        for h in [v2, v3]:
            if h == "VLAN":      read_map["vlan"] = col
            if h == "SWITCH IP": read_map["switch"] = col
            if h == "PORT":      read_map["port"] = col
            if h == "ROOM":      read_map["room"] = col
            if h == "OUTLET":    read_map["outlet"] = col
            if h == "DEVICE":    read_map["device"] = col
            if h == "vlan":      write_map["vlan"] = col
            if h == "switch":    write_map["switch"] = col
            if h == "port":      write_map["port"] = col
            if h == "mac":       write_map["mac"] = col
            if h == "ip":        write_map["ip"] = col

    if not all(k in read_map for k in ["vlan", "switch", "port"]):
        print("[!] ERROR: Header mapping failed. Check row 2 and 3 for VLAN/SWITCH IP/PORT.")
        return

    # 3. Processing
    rows_processed = 0
    rows_skipped = 0

    if args.dry_run:
        print("\n" + "!"*20 + " DRY RUN ACTIVE " + "!"*20)

    for row_idx in range(4, ws.max_row + 1):
        s_val = ws.cell(row=row_idx, column=read_map["switch"]).value
        p_val = ws.cell(row=row_idx, column=read_map["port"]).value
        v_target = ws.cell(row=row_idx, column=read_map["vlan"]).value
        v_current = ws.cell(row=row_idx, column=write_map.get("vlan", 1)).value

        if not s_val or not p_val or v_target is None:
            continue

        # Convert Port date objects back to strings (1-1 logic)
        if hasattr(p_val, 'strftime'):
            p_val = f"{p_val.month}-{p_val.day}"

        # Change Detection
        is_different = str(v_target).strip() != str(v_current).strip()
        is_new = v_current is None

        if not (is_different or is_new):
            rows_skipped += 1
            continue

        # Get Context Data
        room = ws.cell(row=row_idx, column=read_map.get("room", 1)).value or "N/A"
        outlet = ws.cell(row=row_idx, column=read_map.get("outlet", 1)).value or "N/A"

        if args.dry_run:
            print(f"[DRY-RUN] Row {row_idx}: Room {room} | Port {p_val} | Target VLAN: {v_target} (Current: {v_current})")
            rows_processed += 1
            continue

        if args.safe:
            print("\n" + "="*50)
            print(f"LOCATION: Room {room} | Outlet {outlet}")
            print(f"PROPOSED: Row {row_idx} | {s_val} | Port {p_val} -> VLAN {v_target}")
            if input("Apply Config? (y/n): ").lower() != 'y': continue

        # EXECUTION
        res = run_aruba_config(str(s_val), str(p_val), str(v_target))
        
        if res["status"] == "Success":
            config_summary.append({"ip": s_val, "port": p_val, "vlan": v_target})
            # Update cells
            if "switch" in write_map: ws.cell(row=row_idx, column=write_map["switch"], value=s_val)
            if "port" in write_map:   ws.cell(row=row_idx, column=write_map["port"], value=p_val)
            if "vlan" in write_map:   ws.cell(row=row_idx, column=write_map["vlan"], value=v_target)
            if "mac" in write_map:    ws.cell(row=row_idx, column=write_map["mac"], value=res["mac"])
            if "ip" in write_map:     ws.cell(row=row_idx, column=write_map["ip"], value=res["ip"])
            rows_processed += 1
            print(f"    [DONE] Row {row_idx} | VLAN {v_target} Configured")

    # 4. Save & Teams Summary
    print(f"\n[*] Finished. Changes Pending: {rows_processed} | Already Correct: {rows_skipped}")

    if args.dry_run:
        print("[!] Dry run complete. No switches were touched and file was not saved.")
        return

    if rows_processed > 0:
        if os.path.getmtime(DEFAULT_PATH) != start_mtime:
            owner = get_lock_owner(DEFAULT_PATH)
            send_teams_notification("CRITICAL", f"Conflict! {owner} modified the file.")
            print("\n[!] File changed while script was running. Save aborted.")
            return

        try:
            wb.save(DEFAULT_PATH)
            send_teams_notification("SUCCESS", f"Successfully updated {rows_processed} ports.", details=config_summary)
            print(f"[+] Success: Spreadsheet updated.")
        except PermissionError:
            owner = get_lock_owner(DEFAULT_PATH)
            print(f"\n[!] SAVE FAILED: {owner} has the file open.")

if __name__ == "__main__":
    main()