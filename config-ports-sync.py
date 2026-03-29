import os
import re
import sys
import argparse
import time
import datetime
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
TEAMS_WEBHOOK_URL = "https://liveuclac.webhook.office.com/webhookb2/4bf41ac1-0d61-4760-9dea-e0f6184dde8a@1faf88fe-a998-4c5b-93c9-210a11d9a5c2/IncomingWebhook/46bcd5e8d47e4de5b575fe10f189b1e1/43bfe760-7689-4d0b-96fd-46b265519580/V2OHjE1yNbt-Vi6fJNAiAdSwjM90ZnyOKY49V-zdJi0dA1"

# Environment Detection
is_github = os.getenv('GITHUB_ACTIONS') == 'true'

# Path Management
FILE_PATH = Path("/mnt/c/Users/cceadan/OneDrive - University College London/Estates IT - Project Documentation - Patching Schedule/90TCR - Daniel Test.xlsx")
LOG_FILE = FILE_PATH.parent / "automation_audit.log"

def log_event(message):
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open(LOG_FILE, "a") as f:
        f.write(f"[{timestamp}] {message}\n")

def send_teams_notification(status, message, details=None):
    if not TEAMS_WEBHOOK_URL: return 
    color = {"SUCCESS": "28A745", "WARNING": "FFC107", "CRITICAL": "DC3545"}.get(status, "0078D7")
    
    sections = [{"activityTitle": f"**Aruba Sync: {status}**", "text": message}]
    if details:
        facts = [{"name": f"SW {d['ip']} Port {d['port']}", "value": f"VLAN {d['vlan']} | MAC: {d['mac']}"} for d in details]
        sections[0]["facts"] = facts

    payload = {"@type": "MessageCard", "@context": "http://schema.org/extensions", "themeColor": color, "sections": sections}
    try:
        requests.post(TEAMS_WEBHOOK_URL, json=payload, timeout=10)
    except Exception as e:
        print(f"[!] Webhook Failed: {e}")

def run_aruba_config(switch_ip, port, vlan_id):
    device = {'device_type': 'aruba_osswitch', 'host': switch_ip, 'username': USERNAME, 'password': PASSWORD, 'global_delay_factor': 2}
    commands = ["conf t", f"int {port}", f"vlan access {vlan_id}", "exit"]
    result = {"mac": "Unknown", "ip": "No ARP", "status": "Error"}
    
    try:
        with ConnectHandler(**device, conn_timeout=15) as net_connect:
            net_connect.send_config_set(commands)
            
            # MAC Discovery with Fallback
            mac_out = net_connect.send_command(f"show mac-address {port}")
            if "Invalid" in mac_out or not any(c.isdigit() for c in mac_out):
                mac_out = net_connect.send_command(f"show mac-address-table int {port}")
            
            mac_match = re.search(r'([0-9a-fA-F]{2}[:.-]){5}[0-9a-fA-F]{2}|([0-9a-fA-F]{4}[.-]){2}[0-9a-fA-F]{4}', mac_out)
            
            if mac_match:
                result["mac"] = mac_match.group(0)
                # ARP Discovery
                arp_out = net_connect.send_command(f"show arp")
                ip_match = re.search(rf'(\d{{1,3}}\.\d{{1,3}}\.\d{{1,3}}\.\d{{1,3}}).*{re.escape(result["mac"])}', arp_out, re.IGNORECASE)
                if ip_match:
                    result["ip"] = ip_match.group(1)
            
            result["status"] = "Success"
            return result
    except Exception as e:
        result["status"] = str(e)[:25]
        return result

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--safe", action="store_true")
    args = parser.parse_args()

    # 1. Pre-Flight Checks
    if not FILE_PATH.exists():
        print(f"[!] Path Error: {FILE_PATH} not found."); return

    if (FILE_PATH.parent / f"~${FILE_PATH.name}").exists():
        msg = "Aborted: File is open in Excel."
        print(f"[!] {msg}")
        if is_github: send_teams_notification("CRITICAL", msg)
        return

    # 2. Load Workbook
    try:
        wb = load_workbook(FILE_PATH, data_only=False)
        ws = wb.active
        print(f"[*] Loaded: {FILE_PATH.name}")
    except Exception as e:
        print(f"[!] Load Error: {e}"); return

    # 3. Header Mapping (Restored logic)
    read_map = {}; write_map = {}
    for col in range(1, ws.max_column + 1):
        v2 = str(ws.cell(row=2, column=col).value or "").strip().upper()
        v3 = str(ws.cell(row=3, column=col).value or "").strip().lower()
        if "VLAN" in v2: read_map["vlan_target"] = col
        if "SWITCH IP" in v2: read_map["switch"] = col
        if "PORT" in v2: read_map["port"] = col
        if v3 == "vlan": write_map["vlan_curr"] = col
        if v3 == "mac":  write_map["mac"] = col
        if v3 == "ip":   write_map["ip"] = col
        if v3 == "time": write_map["time"] = col

    # Verify Mapping Success
    if not all(k in read_map for k in ["vlan_target", "switch", "port"]):
        print("[!] Header Map Failed. Ensure row 2 has 'VLAN', 'SWITCH IP', 'PORT'."); return

    updates = 0
    config_summary = []

    # 4. Processing Loop
    for row in range(4, ws.max_row + 1):
        s_ip = ws.cell(row=row, column=read_map["switch"]).value
        port = ws.cell(row=row, column=read_map["port"]).value
        v_target = ws.cell(row=row, column=read_map["vlan_target"]).value
        v_curr = ws.cell(row=row, column=write_map.get("vlan_curr", 1)).value

        if not s_ip or not port or v_target is None: continue
        if hasattr(port, 'strftime'): port = f"1/1/{port.day}"

        if str(v_target).strip() != str(v_curr).strip():
            print(f"\n[ACTION] Row {row}: Port {port} | {v_curr} -> {v_target}")
            
            # Auto-approve for GitHub, prompt for --safe
            if args.safe and not is_github:
                if input("    Apply? (y/n): ").lower() != 'y': continue

            res = run_aruba_config(str(s_ip), str(port), str(v_target))
            
            if res["status"] == "Success":
                ws.cell(row=row, column=write_map["vlan_curr"], value=v_target)
                if "mac" in write_map: ws.cell(row=row, column=write_map["mac"], value=res["mac"])
                if "ip" in write_map:  ws.cell(row=row, column=write_map["ip"], value=res["ip"])
                if "time" in write_map: ws.cell(row=row, column=write_map["time"], value=datetime.datetime.now().strftime("%H:%M"))
                
                updates += 1
                config_summary.append({"ip": s_ip, "port": port, "vlan": v_target, "mac": res["mac"]})
                log_event(f"SUCCESS: Row {row} | Port {port} | MAC: {res['mac']} | IP: {res['ip']}")
                print(f"    [DONE] MAC: {res['mac']} | IP: {res['ip']}")
            else:
                print(f"    [FAIL] {res['status']}")

    # 5. Save
    if updates > 0:
        try:
            wb.save(FILE_PATH)
            send_teams_notification("SUCCESS", f"Synced {updates} ports.", details=config_summary)
            print(f"\n[+] Spreadsheet Updated.")
        except PermissionError:
            print("\n[!] Save blocked by Excel. Hardware was updated, but record not saved.")
            send_teams_notification("WARNING", "Hardware updated, but Excel save failed (File Open).")

if __name__ == "__main__":
    main()