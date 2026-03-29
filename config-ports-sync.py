import os
import re
import sys
import argparse
from datetime import datetime
from pathlib import Path

import requests
from dotenv import load_dotenv
from netmiko import ConnectHandler
from openpyxl import load_workbook

# ============================================================
# 1. SETUP & CONFIGURATION
# ============================================================
load_dotenv()

USERNAME = os.getenv("username")
PASSWORD = os.getenv("passwordAD")
TEAMS_WEBHOOK_URL = os.getenv("TEAMS_WEBHOOK_URL", "")

RUN_ACTOR = (
    os.getenv("GITHUB_ACTOR")
    or os.getenv("USER")
    or os.getenv("USERNAME")
    or "unknown"
)

RUN_SOURCE = (
    "GitHub Actions"
    if os.getenv("GITHUB_ACTIONS", "").lower() == "true"
    else "Manual"
)

RUN_AT = datetime.now().astimezone().strftime("%Y-%m-%d %H:%M:%S %Z")

# --- PATH AUTO-DISCOVERY ---
ONEDRIVE_PATH = Path(
    "/mnt/c/Users/cceadan/OneDrive - University College London/Estates IT - Project Documentation - Patching Schedule/90TCR - Daniel Test.xlsx"
)
DOWNLOADS_PATH = Path(
    "/mnt/c/Users/cceadan/Downloads/90 TCR - Daniel-Test.xlsx"
)

if ONEDRIVE_PATH.exists():
    DEFAULT_PATH = ONEDRIVE_PATH
elif DOWNLOADS_PATH.exists():
    DEFAULT_PATH = DOWNLOADS_PATH
else:
    DEFAULT_PATH = ONEDRIVE_PATH


def log(message):
    print(message, flush=True)


def send_teams_notification(status, message, details=None):
    if not TEAMS_WEBHOOK_URL:
        return

    color = {
        "SUCCESS": "28A745",
        "WARNING": "FFC107",
        "CRITICAL": "DC3545",
    }.get(status, "0078D7")

    facts = [
        {"name": "Run By", "value": RUN_ACTOR},
        {"name": "Run At", "value": RUN_AT},
        {"name": "Source", "value": RUN_SOURCE},
        {"name": "File", "value": DEFAULT_PATH.name},
    ]

    if details:
        for entry in details:
            facts.append(
                {
                    "name": f"Switch {entry['ip']}",
                    "value": (
                        f"Port {entry['port']} -> VLAN {entry['vlan']}\n"
                        f"By: {entry['changed_by']}\n"
                        f"At: {entry['changed_at']}\n"
                        f"Source: {entry['source']}"
                    ),
                }
            )

    payload = {
        "@type": "MessageCard",
        "@context": "http://schema.org/extensions",
        "themeColor": color,
        "summary": "Aruba Config Update",
        "sections": [
            {
                "activityTitle": f"**Aruba Configurator: {status}**",
                "activitySubtitle": f"File: {DEFAULT_PATH.name}",
                "text": message,
                "facts": facts,
            }
        ],
    }

    try:
        response = requests.post(TEAMS_WEBHOOK_URL, json=payload, timeout=10)
        response.raise_for_status()
    except Exception as exc:
        log(f"[!] Teams Alert Failed: {exc}")


def get_lock_owner(path):
    lock_path = path.parent / f"~${path.name}"
    if not lock_path.exists():
        return "Unknown (Closed)"

    try:
        with open(lock_path, "rb") as file_handle:
            content = file_handle.read().decode("latin-1", errors="ignore")
            match = re.search(r"[a-zA-Z\s]{3,}", content)
            return match.group(0).strip() if match else "a Colleague"
    except Exception:
        return "a Colleague"


def run_aruba_config(switch_ip, port, vlan_id):
    device = {
        "device_type": "aruba_osswitch",
        "host": switch_ip,
        "username": USERNAME,
        "password": PASSWORD,
    }

    commands = [
        "aruba-central support-mode",
        "conf t",
        f"int {port}",
        f"vlan access {vlan_id}",
    ]

    try:
        with ConnectHandler(**device, conn_timeout=15) as net_connect:
            net_connect.send_config_set(commands)

            mac_out = net_connect.send_command(f"show mac-address-table int {port}")
            mac_match = re.search(
                r"([0-9a-fA-F]{2}[:.-]){5}[0-9a-fA-F]{2}",
                mac_out,
            )
            found_mac = mac_match.group(0) if mac_match else "Unknown"

            arp_out = net_connect.send_command(f"show arp | inc {found_mac}")
            ip_match = re.search(r"(\d{1,3}\.){3}\d{1,3}", arp_out)
            found_ip = ip_match.group(0) if ip_match else "No ARP"

            return {
                "mac": found_mac,
                "ip": found_ip,
                "status": "Success",
            }

    except Exception as exc:
        return {
            "mac": "Error",
            "ip": "Error",
            "status": str(exc)[:200],
        }


def confirm_change(safe_mode, switch_ip, port, current_vlan, target_vlan, row_idx):
    if not safe_mode:
        return True

    if not os.isatty(0):
        raise RuntimeError(
            "--safe was supplied, but no interactive terminal is available."
        )

    print("\n" + "=" * 60)
    print(f"[ACTION REQUIRED] Row {row_idx}: Port {port}")
    print(f"VLAN Mismatch: [{current_vlan}] -> Target: [{target_vlan}]")
    reply = input(f"Update Switch {switch_ip} Port {port}? (y/n): ").strip().lower()
    return reply == "y"


def build_header_maps(ws):
    read_map = {}
    write_map = {}

    for col in range(1, ws.max_column + 1):
        v2 = str(ws.cell(row=2, column=col).value or "").strip()
        v3 = str(ws.cell(row=3, column=col).value or "").strip()

        for header in (v2, v3):
            if header == "VLAN":
                read_map["vlan"] = col
            elif header == "SWITCH IP":
                read_map["switch"] = col
            elif header == "PORT":
                read_map["port"] = col
            elif header == "ROOM":
                read_map["room"] = col
            elif header == "OUTLET":
                read_map["outlet"] = col
            elif header == "DEVICE":
                read_map["device"] = col
            elif header == "vlan":
                write_map["vlan"] = col
            elif header == "switch":
                write_map["switch"] = col
            elif header == "port":
                write_map["port"] = col
            elif header == "mac":
                write_map["mac"] = col
            elif header == "ip":
                write_map["ip"] = col

    return read_map, write_map


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--safe",
        action="store_true",
        help="Confirm each change manually",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Show changes without applying them",
    )
    args = parser.parse_args()

    log("[*] Script starting")
    log(f"[*] Safe mode: {args.safe}")
    log(f"[*] Dry run: {args.dry_run}")
    log(f"[*] Candidate file path: {DEFAULT_PATH}")
    log(f"[*] Run By: {RUN_ACTOR}")
    log(f"[*] Run At: {RUN_AT}")
    log(f"[*] Source: {RUN_SOURCE}")

    if not DEFAULT_PATH.exists():
        log(f"[!] File not found: {DEFAULT_PATH}")
        sys.exit(1)

    if ONEDRIVE_PATH.exists():
        log(f"[*] Using OneDrive Source: {ONEDRIVE_PATH.name}")
    elif DOWNLOADS_PATH.exists():
        log(f"[*] Using Backup Source (Downloads): {DOWNLOADS_PATH.name}")

    if not USERNAME or not PASSWORD:
        log("[!] ERROR: Missing username/passwordAD environment variables.")
        sys.exit(1)

    lock_file = DEFAULT_PATH.parent / f"~${DEFAULT_PATH.name}"
    if lock_file.exists():
        owner = get_lock_owner(DEFAULT_PATH)
        log(f"[!] ABORTED: File is currently open by {owner}.")
        sys.exit(1)

    start_mtime = os.path.getmtime(DEFAULT_PATH)
    config_summary = []

    try:
        wb = load_workbook(DEFAULT_PATH, data_only=False)
        ws = wb.active
        log(f"[*] Workbook loaded: {DEFAULT_PATH}")
        log(f"[*] Active sheet: {ws.title}")
        log(f"[*] Max rows: {ws.max_row}, Max cols: {ws.max_column}")
    except Exception as exc:
        log(f"[!] Error loading workbook: {exc}")
        sys.exit(1)

    read_map, write_map = build_header_maps(ws)

    log(f"[*] Read header map: {read_map}")
    log(f"[*] Write header map: {write_map}")

    if not all(key in read_map for key in ("vlan", "switch", "port")):
        log("[!] ERROR: Header mapping failed. Check row 2 and row 3.")
        sys.exit(1)

    rows_processed = 0
    rows_skipped = 0
    rows_failed = 0
    rows_declined = 0
    candidate_changes = 0

    if args.dry_run:
        log("!!!!!!!!!!!!!!!!!!!! DRY RUN ACTIVE !!!!!!!!!!!!!!!!!!!!")

    for row_idx in range(4, ws.max_row + 1):
        switch_ip = ws.cell(row=row_idx, column=read_map["switch"]).value
        port = ws.cell(row=row_idx, column=read_map["port"]).value
        target_vlan = ws.cell(row=row_idx, column=read_map["vlan"]).value

        current_vlan = None
        if "vlan" in write_map:
            current_vlan = ws.cell(row=row_idx, column=write_map["vlan"]).value

        if not switch_ip or not port or target_vlan is None:
            continue

        if hasattr(port, "strftime"):
            port = f"{port.month}-{port.day}"

        switch_ip = str(switch_ip).strip()
        port = str(port).strip()
        target_vlan = str(target_vlan).strip()
        current_vlan_str = "" if current_vlan is None else str(current_vlan).strip()

        is_new = current_vlan is None or current_vlan_str == ""
        is_different = target_vlan != current_vlan_str

        if not (is_new or is_different):
            rows_skipped += 1
            continue

        candidate_changes += 1

        room = ws.cell(row=row_idx, column=read_map.get("room", 1)).value or "N/A"
        outlet = ws.cell(row=row_idx, column=read_map.get("outlet", 1)).value or "N/A"

        if args.dry_run:
            log(
                f"[DRY-RUN] Row {row_idx}: "
                f"Room {room} | Outlet {outlet} | "
                f"Switch {switch_ip} | Port {port} | "
                f"Current VLAN {current_vlan_str or 'None'} -> Target VLAN {target_vlan}"
            )
            rows_processed += 1
            continue

        try:
            should_apply = confirm_change(
                safe_mode=args.safe,
                switch_ip=switch_ip,
                port=port,
                current_vlan=current_vlan_str or "None",
                target_vlan=target_vlan,
                row_idx=row_idx,
            )
        except RuntimeError as exc:
            log(f"[!] {exc}")
            sys.exit(1)

        if not should_apply:
            rows_declined += 1
            log(f"[SKIPPED] Row {row_idx} declined by user.")
            continue

        changed_by = RUN_ACTOR
        changed_at = datetime.now().astimezone().strftime("%Y-%m-%d %H:%M:%S %Z")
        source = RUN_SOURCE

        log(
            f"[*] Applying Row {row_idx} | "
            f"Switch {switch_ip} | Port {port} | "
            f"{current_vlan_str or 'None'} -> {target_vlan} | "
            f"By: {changed_by} | At: {changed_at} | Source: {source}"
        )

        result = run_aruba_config(switch_ip, port, target_vlan)

        if result["status"] == "Success":
            if "vlan" in write_map:
                ws.cell(row=row_idx, column=write_map["vlan"], value=target_vlan)
            if "mac" in write_map:
                ws.cell(row=row_idx, column=write_map["mac"], value=result["mac"])
            if "ip" in write_map:
                ws.cell(row=row_idx, column=write_map["ip"], value=result["ip"])

            config_summary.append(
                {
                    "ip": switch_ip,
                    "port": port,
                    "vlan": target_vlan,
                    "changed_by": changed_by,
                    "changed_at": changed_at,
                    "source": source,
                }
            )

            rows_processed += 1
            log(
                f"[DONE] Row {row_idx} | {switch_ip} | Port {port} -> VLAN {target_vlan} | "
                f"By: {changed_by} | At: {changed_at} | Source: {source}"
            )
        else:
            rows_failed += 1
            log(
                f"[FAILED] Row {row_idx} | {switch_ip} | Port {port} | "
                f"Error: {result['status']} | By: {changed_by} | At: {changed_at} | Source: {source}"
            )

    log(
        f"[*] Summary | Candidates: {candidate_changes} | Applied: {rows_processed} | "
        f"Already Correct: {rows_skipped} | Declined: {rows_declined} | Failed: {rows_failed}"
    )

    if args.dry_run:
        log("[*] Dry run finished successfully.")
        sys.exit(0)

    if rows_processed > 0:
        if os.path.getmtime(DEFAULT_PATH) != start_mtime:
            owner = get_lock_owner(DEFAULT_PATH)
            send_teams_notification(
                "CRITICAL",
                f"Conflict detected. {owner} modified the file while the script was running.",
            )
            log("[!] File changed while script was running. Save aborted.")
            sys.exit(1)

        try:
            wb.save(DEFAULT_PATH)
            log("[+] Success: Spreadsheet updated.")
        except PermissionError:
            owner = get_lock_owner(DEFAULT_PATH)
            log(f"[!] SAVE FAILED: {owner} has the file open.")
            sys.exit(1)
        except Exception as exc:
            log(f"[!] SAVE FAILED: {exc}")
            sys.exit(1)

    if rows_failed > 0:
        send_teams_notification(
            "WARNING",
            f"Completed with errors. Applied {rows_processed} changes, but {rows_failed} failed.",
            details=config_summary,
        )
        sys.exit(1)

    if rows_processed > 0:
        send_teams_notification(
            "SUCCESS",
            f"Successfully updated {rows_processed} ports.",
            details=config_summary,
        )
        sys.exit(0)

    log("[*] No spreadsheet changes needed.")
    sys.exit(0)


if __name__ == "__main__":
    main()