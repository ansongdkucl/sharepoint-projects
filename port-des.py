import os
import pandas as pd
from netmiko import ConnectHandler
from dotenv import load_dotenv
from pathlib import Path

# ============================================================
# 1. SETUP & CONFIGURATION
# ============================================================
load_dotenv()

USERNAME = os.getenv("username")
PASSWORD = os.getenv("passwordAD")
EXCEL_FILE = Path("des.xlsx")

def apply_descriptions():
    if not EXCEL_FILE.exists():
        print(f"[!] Error: {EXCEL_FILE} not found.")
        return

    try:
        df = pd.read_excel(EXCEL_FILE)
        df.columns = [c.strip().lower() for c in df.columns]
    except Exception as e:
        print(f"[!] Failed to read Excel: {e}")
        return

    switches = df['switch'].unique()

    for sw_ip in switches:
        port_configs = df[df['switch'] == sw_ip]
        
        device = {
            'device_type': 'aruba_osswitch',
            'host': str(sw_ip).strip(),
            'username': USERNAME,
            'password': PASSWORD,
            'global_delay_factor': 1,
            'fast_cli': False,
            'session_log': f"netmiko-{sw_ip}.log",
        }

        try:
            print(f"[*] Connecting to {sw_ip}...")
            with ConnectHandler(**device) as net_connect:
                print(f"[*] Applying {len(port_configs)} descriptions (No Quotes)...")
                commands = [
                    "aruba-central support-mode",
                    "configure terminal",
                ]
                for _, row in port_configs.iterrows():
                    port = str(row['interfaces']).strip()
                    desc = str(row['description']).strip()
                    commands.extend([
                        f"interface {port}",
                        f"description {desc}",
                    ])

                net_connect.send_config_set(commands)
                net_connect.send_command("write memory", expect_string=r"#", read_timeout=30)
                print(f"[+] Success for {sw_ip}!")

        except Exception as e:
            print(f"[!] Error on {sw_ip}: {e}")

if __name__ == "__main__":
    apply_descriptions()