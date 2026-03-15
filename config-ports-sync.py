import os
import re
import sys
import argparse
import time
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

# WSL Path to SharePoint Sync
DEFAULT_PATH = Path("/mnt/c/Users/cceadan/University College London/Estates IT - Project Documentation - Patching Schedule/90 TCR - Level 3A Patching Schedule.xlsx")

def check_sync_status(path):
    """Checks if the file is currently 'Sync Pending' or 'Locked'."""
    if not path.exists():
        return False
    try:
        # Try to open for appending briefly to see if Windows/OneDrive has a lock
        with open(path, 'a'):
            pass
        return True
    except OSError:
        return False

def run_aruba_config(switch_ip, port, vlan_id):
    device = {
        'device_type': 'aruba_osswitch',
        'host': switch_ip,
        'username': USERNAME,
        'password': PASSWORD,
    }
    
    commands = [
        "aruba-central support-mode",
        "conf t",
        f"int {port}",
        f"vlan access {vlan_id}"
    ]
    
    try:
        with ConnectHandler(**device, conn_timeout=15) as net_connect:
            print(f"    [SSH] Connected to {switch_ip}. Configuring Port {port}...")
            net_connect.send_config_set(commands)
            
            # Fetch MAC
            mac_out = net_connect.send_command(f"show mac-address-table int {port}")
            mac_match = re.search(r'([0-9a-fA-F]{2}[:.-]){5}[0-9a-fA-F]{2}', mac_out)
            found_mac = mac_match.group(0) if mac_match else "Unknown"
            
            # Fetch IP
            arp_out = net_connect.send_command(f"show arp | inc {found_mac}")
            ip_match = re.search(r'(\d{1,3}\.){3}\d{1,3}', arp_out)
            ip_addr = ip_match.group(0) if ip_match else "No ARP"
                
            return {"mac": found_mac, "ip": ip_addr, "status": "Success"}
    except Exception as e:
        return {"mac": "N/A", "ip": "N/A", "status": f"Error: {str(e)[:20]}"}

# ============================================================
# 2. MAIN EXECUTION
# ============================================================
def main():
    parser = argparse.ArgumentParser(description="Aruba Port Configurator")
    parser.add_argument("--safe", action="store_true", help="Enable Safe Mode")
    args = parser.parse_args()

    print(f"\n[*] Checking Sync Status for: {DEFAULT_PATH.name}")
    if not check_sync_status(DEFAULT_PATH):
        print("[!] WARNING: File is currently 'Sync Pending' or open in Excel.")
        print("[!] Please close Excel or wait for OneDrive arrows to turn into a green check.")
        confirm = input("    Continue anyway? (y/n): ").lower()
        if confirm != 'y': return

    try:
        wb = load_workbook(DEFAULT_PATH, data_only=True)
        ws = wb.active
        print(f"[*] File Opened. Active Sheet: {ws.title}")
    except Exception as e:
        print(f"[!] FATAL: Could not open file. {e}")
        return

    # --- STEP 1: HEADER MAPPING ---
    read_map = {}
    write_map = {}
    for col in range(1, ws.max_column + 1):
        v2 = str(ws.cell(row=2, column=col).value or "").strip()
        v3 = str(ws.cell(row=3, column=col).value or "").strip()
        for h in [v2, v3]:
            if h == "VLAN":      read_map["vlan"] = col
            if h == "SWITCH IP": read_map["switch"] = col
            if h == "PORT":      read_map["port"] = col
            if h == "vlan":      write_map["vlan"] = col
            if h == "switch":    write_map["switch"] = col
            if h == "port":      write_map["port"] = col
            if h == "mac":       write_map["mac"] = col
            if h == "ip":        write_map["ip"] = col

    if not all(k in read_map for k in ["vlan", "switch", "port"]):
        print("[!] ERROR: Headers not found.")
        return

    # --- STEP 2: PROCESSING ---
    rows_processed = 0
    for row_idx in range(4, ws.max_row + 1):
        v_val = ws.cell(row=row_idx, column=read_map["vlan"]).value
        s_val = ws.cell(row=row_idx, column=read_map["switch"]).value
        p_val = ws.cell(row=row_idx, column=read_map["port"]).value

        if not s_val or not p_val: continue

        if args.safe:
            print(f"\n>> PROPOSED: Row {row_idx} | {s_val} Port {p_val}")
            if input("   Apply? (y/n): ").lower() != 'y': continue

        result = run_aruba_config(str(s_val), str(p_val), str(v_val))
        
        # Write to lowercase result columns
        for key in ["switch", "port", "vlan", "mac", "ip"]:
            if key in write_map:
                val = result.get(key, locals().get(f"{key}_req")) if key not in ["mac", "ip"] else result[key]
                ws.cell(row=row_idx, column=write_map[key], value=val)

        print(f"    [DONE] Row {row_idx} | MAC: {result['mac']}")
        rows_processed += 1

    # --- STEP 3: SYNC-AWARE SAVE ---
    if rows_processed > 0:
        for i in range(5):
            try:
                wb.save(DEFAULT_PATH)
                print(f"\n[+] SUCCESS: {rows_processed} rows updated. OneDrive sync triggered.")
                break
            except PermissionError:
                print(f"[!] Sync Lock detected. Retrying save in 5s... ({i+1}/5)")
                time.sleep(5)
    else:
        print("\n[!] No rows were processed.")

if __name__ == "__main__":
    main()