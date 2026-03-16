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
    """Sends a formatted message card to a Microsoft Teams channel."""
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
    """Tries to read the username from the hidden Excel lock file (~$)."""
    lock_path = path.parent / f"~${path.name}"
    if not lock_path.exists(): return "Unknown (Closed)"
    try:
        with open(lock_path, 'rb') as f:
            content = f.read().decode('latin-1', errors='ignore')
            match = re.search(r'[a-zA-Z\s]{3,}', content)
            return match.group(0).strip() if match else "a Colleague"
    except: return "a Colleague"

def run_aruba_config(switch_ip, port, vlan_id):
    """SSH Connection logic for Aruba switches."""
    device = {'device_type': 'aruba_osswitch', 'host': switch_ip, 'username': USERNAME, 'password': PASSWORD}
    commands = ["aruba-central support-mode", "conf t", f"int {port}", f"vlan access {vlan_id}"]
    
    try:
        with ConnectHandler(**device, conn_timeout=15) as net_connect:
            net_connect.send_config_set(commands)
            mac_out = net_connect.send_command(f"show mac-address-table int {port}")
            mac_match = re.search(r'([0-9a-fA-F]{2}[:.-]){5}[0-9a-fA-F]{2}', mac_out)
            found_mac = mac_match.group(0) if mac_match else "Unknown"
            
            # Fetch IP
            arp_out = net_connect.send_command(f"show arp | inc {found_mac}")
            ip_match = re.search(r'(\d{1,3}\.){3}\d{1,3}', arp_out)
            found_ip = ip_match.group(0) if ip_match else "No ARP"
            
            return {"mac": found_mac, "ip": found_ip, "status": "Success"}
    except Exception as e:
        return {"mac": "Error", "ip": "Error", "status": str(e)[:20]}

# ============================================================
# 2. MAIN EXECUTION
# ============================================================
def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--safe", action="store_true")
    args = parser.parse_args()

    if not DEFAULT_PATH.exists(): 
        print(f"[!] File not found: {DEFAULT_PATH}")
        return

    # 1. Pre-run lock check
    if (DEFAULT_PATH.parent / f"~${DEFAULT_PATH.name}").exists():
        owner = get_lock_owner(DEFAULT_PATH)
        send_teams_notification("WARNING", f"Script aborted. **{owner}** is currently editing the file.")
        print(f"[!] Locked by {owner}. Aborting.")
        return

    start_mtime = os.path.getmtime(DEFAULT_PATH)
    config_summary = [] 

    try:
        wb = load_workbook(DEFAULT_PATH, data_only=False)
        ws = wb.active
    except Exception as e:
        print(f"Error: {e}"); return

    # 2. Header Mapping
    read_map = {}
    write_map = {}
    for col in range(1, ws.max_column + 1):
        v2 = str(ws.cell(row=2, column=col).value or "").strip().upper()
        v3 = str(ws.cell(row=3, column=col).value or "").strip().upper()
        
        for h in [v2, v3]:
            # Inputs
            if h == "VLAN":      read_map["vlan"] = col
            if h == "SWITCH IP": read_map["switch"] = col
            if h == "PORT":      read_map["port"] = col
            # Context (Read-Only)
            if h == "ROOM":      read_map["room"] = col
            if h == "OUTLET":    read_map["outlet"] = col
            if h == "DEVICE":    read_map["device"] = col
            # Outputs
            if h == "MAC":       write_map["mac"] = col
            if h == "IP":        write_map["ip"] = col

    if not all(k in read_map for k in ["vlan", "switch", "port"]):
        print("[!] ERROR: Required headers (VLAN, SWITCH IP, PORT) not found.")
        return

    # 3. Processing Rows
    rows_processed = 0
    for row_idx in range(4, ws.max_row + 1):
        s_val = ws.cell(row=row_idx, column=read_map["switch"]).value
        p_val = ws.cell(row=row_idx, column=read_map["port"]).value
        v_val = ws.cell(row=row_idx, column=read_map["vlan"]).value
        
        # Get Context for Display
        room = ws.cell(row=row_idx, column=read_map.get("room", 1)).value or "N/A"
        outlet = ws.cell(row=row_idx, column=read_map.get("outlet", 1)).value or "N/A"
        device = ws.cell(row=row_idx, column=read_map.get("device", 1)).value or "N/A"

        if not s_val or not p_val: continue

        # Fix Excel Date Bug (e.g. 1-1 becoming 2002-01-01)
        if hasattr(p_val, 'strftime'):
            p_val = f"{p_val.month}-{p_val.day}"

        if args.safe:
            print("\n" + "="*50)
            print(f"LOCATION: Room {room} | Outlet {outlet}")
            print(f"DEVICE  : {device}")
            print(f"PROPOSED: Row {row_idx} | {s_val} | Port {p_val} -> VLAN {v_val}")
            if input("Apply? (y/n): ").lower() != 'y': continue

        res = run_aruba_config(str(s_val), str(p_val), str(v_val))
        
        if res["status"] == "Success":
            config_summary.append({"ip": s_val, "port": p_val, "vlan": v_val})
            if "mac" in write_map:
                ws.cell(row=row_idx, column=write_map["mac"], value=res["mac"])
            if "ip" in write_map:
                ws.cell(row=row_idx, column=write_map["ip"], value=res["ip"])
            
            rows_processed += 1
            print(f"    [DONE] Row {row_idx} | MAC: {res['mac']}")

    # 4. Final Safety Check & Save
    if rows_processed > 0:
        if os.path.getmtime(DEFAULT_PATH) != start_mtime:
            owner = get_lock_owner(DEFAULT_PATH)
            msg = f"**Conflict Detected!** {owner} modified the file. Save aborted."
            send_teams_notification("CRITICAL", msg)
            print(f"\n[!] {msg}")
            return

        try:
            wb.save(DEFAULT_PATH)
            msg = f"Successfully updated **{rows_processed}** ports."
            send_teams_notification("SUCCESS", msg, details=config_summary)
            print(f"\n[+] {msg}")
        except PermissionError:
            owner = get_lock_owner(DEFAULT_PATH)
            print(f"\n[!] Save failed. {owner} has file open.")
    else:
        print("\n[!] No rows were processed.")

if __name__ == "__main__":
    main()