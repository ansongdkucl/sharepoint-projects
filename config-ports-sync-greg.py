import os, re, sys, argparse, requests, platform
from datetime import datetime
from pathlib import Path
from collections import defaultdict
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from netmiko import ConnectHandler
from dotenv import load_dotenv

# ============================================================
# 1. SETUP, VALUES & CONFIGURATION
# ============================================================
load_dotenv()

USERNAME = os.getenv("username")
PASSWORD = os.getenv("passwordAD")
TEAMS_WEBHOOK_URL = os.getenv("TEAMS_WEBHOOK_URL", "")

RUN_ACTOR = os.getenv("GITHUB_ACTOR") or os.getenv("USER") or os.getenv("USERNAME") or "unknown"
RUN_SOURCE = "GitHub Actions" if os.getenv("GITHUB_ACTIONS", "").lower() == "true" else "Manual"

WORKBOOK_NAME = "FC-MSA-CI.xlsx"
WORKBOOK_DIR = r"University College London\ISD.ITSD.CO.Technical Specialists - patching"
WINDOWS_WORKBOOK_PATH = rf"C:\Users\anson\{WORKBOOK_DIR}\{WORKBOOK_NAME}"


def normalize_workbook_path(value):
    path_text = str(value).strip().strip('"')

    if os.name != "nt":
        m = re.match(r"^([a-zA-Z]):\\(.*)$", path_text)
        if m:
            drive, rest = m.groups()
            rest = rest.replace("\\", "/")
            return Path(f"/mnt/{drive.lower()}/{rest}")

    return Path(path_text)


# Fixed workbook path for testing - supports Windows and WSL/Linux
DEFAULT_PATH = normalize_workbook_path(os.getenv("WORKBOOK_PATH") or WINDOWS_WORKBOOK_PATH)

# ============================================================
# 2. GENERAL & CELL HELPERS
# ============================================================
def now_str():
    return datetime.now().astimezone().strftime("%Y-%m-%d %H:%M:%S %Z")

def log(msg):
    print(msg, flush=True)

def clean_text(v):
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v).strip()

def normalize_header(v):
    return re.sub(r"\s+", " ", clean_text(v).replace("\n", " ")).strip().lower()

def normalize_mac(v):
    return re.sub(r"[^0-9a-fA-F]", "", str(v or "")).lower()

def first_match(pattern, text):
    m = re.search(pattern, text)
    return m.group(0) if m else "Unknown"

def get_lock_owner(path):
    lk = path.parent / f"~${path.name}"
    if not lk.exists():
        return "Unknown (Closed)"
    try:
        return first_match(
            r"[a-zA-Z\s]{3,}",
            lk.read_bytes().decode("latin-1", errors="ignore")
        ).strip() or "a Colleague"
    except Exception:
        return "a Colleague"

def confirm_change(safe, ip, port, cur, tgt, row, s_name):
    if not safe:
        return True
    if not os.isatty(0):
        raise RuntimeError("--safe was supplied, but no interactive terminal is available.")
    print(
        f"\n{'='*60}\n"
        f"[ACTION REQUIRED] Sheet {s_name} Row {row}: {ip} {port}\n"
        f"Live VLAN: [{cur}] -> Target VLAN: [{tgt}]"
    )
    return input("Apply change? (y/n): ").strip().lower() == "y"

# ============================================================
# 3. TEAMS ENGINE
# ============================================================
def build_adaptive_card(status, message, details=None):
    color = {
        "SUCCESS": "Good",
        "WARNING": "Warning",
        "CRITICAL": "Attention",
    }.get(status, "Accent")

    facts = [
        {"title": "Run By", "value": RUN_ACTOR},
        {"title": "Run At", "value": now_str()},
        {"title": "Source", "value": RUN_SOURCE},
        {"title": "File", "value": DEFAULT_PATH.name},
    ]

    body = [
        {
            "type": "TextBlock",
            "text": f"Aruba Configurator: {status}",
            "weight": "Bolder",
            "size": "Medium",
            "color": color,
            "wrap": True,
        },
        {
            "type": "TextBlock",
            "text": message,
            "wrap": True,
        },
        {
            "type": "FactSet",
            "facts": facts,
        },
    ]

    for e in (details or []):
        body.append({
            "type": "TextBlock",
            "text": (
                f"**{e['sheet']} | {e['ip']}**\n\n"
                f"Port {e['port']} -> VLAN {e['target_vlan']}\n\n"
                f"Previous Live VLAN: {e['old_vlan']}\n\n"
                f"By: {e['changed_by']}\n\n"
                f"At: {e['changed_at']}\n\n"
                f"Source: {e['source']}"
            ),
            "wrap": True,
            "spacing": "Medium",
        })

    return {
        "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
        "type": "AdaptiveCard",
        "version": "1.4",
        "body": body,
    }

def send_teams_notification(status, message, details=None):
    if not TEAMS_WEBHOOK_URL:
        return

    try:
        requests.post(
            TEAMS_WEBHOOK_URL,
            json={"adaptive_card": build_adaptive_card(status, message, details)},
            timeout=10
        ).raise_for_status()
    except Exception as exc:
        log(f"[!] Teams Alert Failed: {exc}")

# ============================================================
# 4. EXCEL SCANNING / BLOCK PARSING
# ============================================================
def row_text_map(ws, row):
    return {
        c: clean_text(ws.cell(row=row, column=c).value)
        for c in range(1, ws.max_column + 1)
        if ws.cell(row=row, column=c).value is not None
    }

def is_primary_header_row(ws, row):
    v = {normalize_header(x) for x in row_text_map(ws, row).values()}
    return (
        "vlan" in v
        and "port" in v
        and any(x in v for x in ("switch ip", "switch id"))
        and any(x in v for x in ("data outlet id", "service description", "room name"))
    )

def is_secondary_header_row(ws, row):
    return any(
        x in {normalize_header(v) for v in row_text_map(ws, row).values()}
        for x in ("bldg", "cr", "outlet")
    )

def resolve_block_columns(ws, h_row):
    ex = {}
    norm = defaultdict(list)

    for col, val in row_text_map(ws, h_row).items():
        ex.setdefault(val, col)
        norm[normalize_header(val)].append(col)

    def find_col(keys, is_ex=False, min_c=0):
        if is_ex:
            return next((ex[k] for k in keys if k in ex), None)
        return max(([c for k in keys for c in norm.get(k, []) if c > min_c]), default=None)

    in_vlan = find_col(["VLAN", "Target VLAN", "Requested VLAN"], True) or find_col(
        ["vlan", "target vlan", "requested vlan"]
    )
    sw_ip = find_col(["SWITCH IP", "Switch IP", "Switch ID", "switch ip", "switch id"], True) or find_col(
        ["switch ip", "switch id"]
    )
    in_prt = find_col(["Port", "PORT"], True) or find_col(["port"])

    if not all([in_vlan, sw_ip, in_prt]):
        return {"input_vlan": None, "switch_ip": None, "input_port": None}

    mx_in = max(c for c in (in_vlan, sw_ip, in_prt) if c)

    out = {
        k: find_col(v, min_c=mx_in)
        for k, v in {
            "out_switch": ["switch", "configured switch", "live switch"],
            "out_port": ["port", "configured port", "live port"],
            "out_vlan": ["vlan", "configured vlan", "live vlan", "done vlan"],
            "out_mac": ["mac", "mac address"],
            "out_ip": ["ip", "ip address"],
            "out_time": ["time", "last checked", "checked at"],
        }.items()
    }

    out["out_notes"] = find_col(["notes", "status", "Status"], True) or find_col(
        ["status", "notes"], min_c=mx_in
    )

    if out["out_notes"] and any(c and c > out["out_notes"] for c in out.values() if c != out["out_notes"]):
        out["out_notes"] = None

    return {"input_vlan": in_vlan, "switch_ip": sw_ip, "input_port": in_prt, **out}

def collect_sheet_blocks(ws):
    blocks = []
    h_rows = [r for r in range(1, ws.max_row + 1) if is_primary_header_row(ws, r)]

    for idx, hr in enumerate(h_rows):
        nxt = h_rows[idx + 1] if idx + 1 < len(h_rows) else ws.max_row + 1
        cols = resolve_block_columns(ws, hr)

        if all(cols[x] for x in ["input_vlan", "switch_ip", "input_port"]):
            title = "Unknown"
            for r in range(hr - 1, max(hr - 3, 0), -1):
                v = clean_text(ws.cell(row=r, column=1).value)
                if v and normalize_header(v) not in {"bldg", "data outlet id"}:
                    title = v
                    break

            blocks.append({
                "sheet_name": ws.title,
                "section_name": title if title != "Unknown" else ws.title,
                "header_row": hr,
                "data_start": hr + 2 if is_secondary_header_row(ws, hr + 1) else hr + 1,
                "data_end": nxt - 1,
                "columns": cols
            })

    return blocks

def write_result_columns(ws, row, cols, **kwargs):
    mapping = {
        "out_switch": kwargs.get("switch_ip"),
        "out_port": kwargs.get("port"),
        "out_vlan": kwargs.get("vlan"),
        "out_mac": kwargs.get("mac"),
        "out_ip": kwargs.get("ip"),
        "out_time": kwargs.get("checked_at"),
        "out_notes": kwargs.get("notes"),
    }
    for k, v in mapping.items():
        if cols.get(k):
            ws.cell(row=row, column=cols[k], value=v)

def is_yellow_fill(fill):
    if not fill or not fill.fill_type:
        return False

    for color in (fill.fgColor, fill.start_color):
        if not color:
            continue
        if color.type == "rgb" and color.rgb:
            rgb = color.rgb.upper()[-6:]
            try:
                red = int(rgb[0:2], 16)
                green = int(rgb[2:4], 16)
                blue = int(rgb[4:6], 16)
            except ValueError:
                continue
            if red >= 220 and green >= 180 and blue <= 140:
                return True
        if color.type == "indexed" and color.indexed in {6, 27, 36, 44}:
            return True
        if color.type == "theme" and fill.fill_type == "solid":
            return True

    return False

def is_highlighted_row(ws, row):
    if is_yellow_fill(ws.row_dimensions[row].fill):
        return True
    return any(is_yellow_fill(ws.cell(row=row, column=c).fill) for c in range(1, ws.max_column + 1))

def clear_yellow_highlight(ws, row):
    if is_yellow_fill(ws.row_dimensions[row].fill):
        ws.row_dimensions[row].fill = PatternFill(fill_type=None)
    for c in range(1, ws.max_column + 1):
        cell = ws.cell(row=row, column=c)
        if is_yellow_fill(cell.fill):
            cell.fill = PatternFill(fill_type=None)

# ============================================================
# 5. SWITCH INFRASTRUCTURE LOGIC
# ============================================================
def run_cmd_safe(conn, cmd, t=30):
    try:
        return conn.send_command(cmd, read_timeout=t)
    except Exception:
        return ""

def get_port_live_details(conn, port):
    mac = "Unknown"

    for cmd in [f"show mac-address-table interface {port}", f"show mac-address-table int {port}"]:
        mac = first_match(
            r"(?:[0-9a-fA-F]{2}[:.-]){5}[0-9a-fA-F]{2}|(?:[0-9a-fA-F]{4}\.){2}[0-9a-fA-F]{4}",
            run_cmd_safe(conn, cmd)
        )
        if mac != "Unknown":
            break

    if mac == "Unknown":
        mac = next(
            (
                first_match(
                    r"(?:[0-9a-fA-F]{2}[:.-]){5}[0-9a-fA-F]{2}|(?:[0-9a-fA-F]{4}\.){2}[0-9a-fA-F]{4}",
                    l
                )
                for l in run_cmd_safe(conn, "show mac-address-table", 60).splitlines()
                if port in l
            ),
            "Unknown"
        )

    ip, m_norm = "Unknown", normalize_mac(mac)
    if m_norm:
        for cmd in ("show arp", "show arp all-vrfs"):
            ip = next(
                (
                    first_match(r"\b(?:\d{1,3}\.){3}\d{1,3}\b", l)
                    for l in run_cmd_safe(conn, cmd, 60).splitlines()
                    if m_norm in normalize_mac(l)
                ),
                "Unknown"
            )
            if ip != "Unknown":
                break

    return {"mac": mac, "ip": ip}

def parse_switch_port(port):
    m = re.match(r"^(\d+/\d+/)(\d+)$", clean_text(port))
    if not m:
        return None
    prefix, number = m.groups()
    return prefix, int(number)

def build_vlan_change_groups(items):
    groups = []
    sortable = defaultdict(list)

    for idx, item in enumerate(items):
        parsed = parse_switch_port(item["port"])
        if parsed:
            prefix, number = parsed
            sortable[(item["target_vlan"], prefix)].append((number, idx, item))
        else:
            groups.append({
                "order": idx,
                "interface": item["port"],
                "target_vlan": item["target_vlan"],
                "items": [item],
            })

    for (vlan, prefix), ports in sortable.items():
        ports.sort(key=lambda x: x[0])
        start_num, start_idx, first_item = ports[0]
        prev_num, prev_idx = start_num, start_idx
        current_items = [first_item]

        for number, idx, item in ports[1:]:
            if number == prev_num + 1:
                current_items.append(item)
                prev_num, prev_idx = number, idx
                continue

            interface = (
                f"{prefix}{start_num}"
                if start_num == prev_num
                else f"{prefix}{start_num}-{prefix}{prev_num}"
            )
            groups.append({
                "order": start_idx,
                "interface": interface,
                "target_vlan": vlan,
                "items": current_items,
            })
            start_num, start_idx = number, idx
            prev_num, prev_idx = number, idx
            current_items = [item]

        interface = (
            f"{prefix}{start_num}"
            if start_num == prev_num
            else f"{prefix}{start_num}-{prefix}{prev_num}"
        )
        groups.append({
            "order": start_idx,
            "interface": interface,
            "target_vlan": vlan,
            "items": current_items,
        })

    return sorted(groups, key=lambda x: x["order"])

def apply_vlan_change(conn, interface, vlan):
    for cmd in ["configure terminal", f"interface {interface}", f"vlan access {vlan}", "end"]:
        conn.send_command_timing(cmd)

# ============================================================
# 6. PIPELINE CONTROLLER
# ============================================================
def main():
    global DEFAULT_PATH

    parser = argparse.ArgumentParser()
    parser.add_argument("--safe", action="store_true")
    parser.add_argument("--dry-run", action="store_true")
    parser.add_argument(
        "--workbook",
        default=None,
        help="Workbook path. Accepts Windows paths under WSL, for example C:\\Users\\anson\\...\\FC-MSA-CI.xlsx",
    )
    args = parser.parse_args()
    if args.workbook:
        DEFAULT_PATH = normalize_workbook_path(args.workbook)

    log(f"[*] Script starting\n[*] Safe mode: {args.safe}\n[*] Dry run: {args.dry_run}\n[*] Path: {DEFAULT_PATH}")
    log(f"[*] Python executable: {sys.executable}")
    log(f"[*] Platform: {platform.platform()}")
    log(f"[*] os.name: {os.name}")
    log(f"[*] Workbook exists check: {DEFAULT_PATH.exists()}")

    if not DEFAULT_PATH.exists():
        log(f"[!] Workbook not found: {DEFAULT_PATH}")
        sys.exit(1)

    if not USERNAME:
        log("[!] Missing environment variable: username")
        sys.exit(1)

    if not PASSWORD:
        log("[!] Missing environment variable: passwordAD")
        sys.exit(1)

    lk = DEFAULT_PATH.parent / f"~${DEFAULT_PATH.name}"
    if lk.exists():
        log(f"[!] ABORTED: Open by {get_lock_owner(DEFAULT_PATH)}.")
        sys.exit(1)

    start_mtime = os.path.getmtime(DEFAULT_PATH)

    try:
        wb = load_workbook(DEFAULT_PATH, data_only=False)
        log(f"[*] Workbook loaded successfully: {DEFAULT_PATH.name}")
        log(f"[*] Sheets found: {wb.sheetnames}")
    except Exception as exc:
        log(f"[!] Load error: {exc}")
        sys.exit(1)

    all_blocks = [(ws, b) for ws in wb.worksheets for b in collect_sheet_blocks(ws)]
    log(f"[*] Total blocks found: {len(all_blocks)}")

    if not all_blocks:
        log("[*] No blocks.")
        sys.exit(0)

    stats = {"chk": 0, "ok": 0, "chg": 0, "fail": 0, "dec": 0}
    wb_touch = False
    summary = []
    no_change_highlights = []

    if args.dry_run:
        log("!!!!!!!!!!!!!!!!!!!! DRY RUN ACTIVE !!!!!!!!!!!!!!!!!!!!")

    for ws, b in all_blocks:
        log(
            f"[*] Processing sheet '{b['sheet_name']}' "
            f"section '{b['section_name']}' "
            f"rows {b['data_start']} to {b['data_end']}"
        )

        cols = b["columns"]
        rows_by_sw = defaultdict(list)
        highlighted_rows = 0
        skipped_highlights = []

        for r in range(b["data_start"], b["data_end"] + 1):
            if not is_highlighted_row(ws, r):
                continue

            sw, pt, tg = [
                clean_text(ws.cell(row=r, column=cols[k]).value)
                for k in ["switch_ip", "input_port", "input_vlan"]
            ]
            if sw and pt and tg:
                highlighted_rows += 1
                rows_by_sw[sw].append({
                    "row_idx": r,
                    "port": pt,
                    "target_vlan": tg,
                    "cell": ws.cell(row=r, column=cols["input_vlan"])
                })
            else:
                skipped_highlights.append(r)

        log(f"[*] Highlighted rows queued in this block: {highlighted_rows}")
        if skipped_highlights:
            log(f"[!] Highlighted rows skipped due to missing switch/port/VLAN: {skipped_highlights}")
        log(f"[*] Unique switches in this block: {len(rows_by_sw)}")

        for sw_ip, entries in rows_by_sw.items():
            log(f"[*] Connecting to switch: {sw_ip} ({len(entries)} row(s))")

            try:
                with ConnectHandler(
                    device_type="aruba_aoscx",
                    host=sw_ip,
                    username=USERNAME,
                    password=PASSWORD,
                    conn_timeout=20,
                    fast_cli=False
                ) as conn:
                    log(f"[+] Connected to {sw_ip}")

                    try:
                        conn.send_command_timing("no page")
                        conn.send_command_timing("aruba-central support-mode")
                    except Exception:
                        pass

                    cur_v_map = {}
                    for l in conn.send_command("show int br", read_timeout=60).splitlines():
                        m = re.match(r"^\s*(\d+/\d+/\d+)\s+(\S+)", l.rstrip())
                        if m:
                            cur_v_map[m.group(1).strip()] = m.group(2).strip()

                    approved_changes = []
                    pending = []

                    for e in entries:
                        r = e["row_idx"]
                        pt = e["port"]
                        tg = e["target_vlan"]
                        cur_v = cur_v_map.get(pt, "Unknown")
                        stats["chk"] += 1

                        log(f"[*] Row {r} | Port {pt} | Current VLAN {cur_v} | Target VLAN {tg}")

                        if cur_v == tg:
                            stats["ok"] += 1
                            log(f"[OK] Row {r} already on correct VLAN")
                            if not args.dry_run:
                                for k, v in {
                                    "out_switch": sw_ip,
                                    "out_port": pt,
                                    "out_vlan": cur_v,
                                    "out_time": now_str(),
                                    "out_notes": "No change needed",
                                }.items():
                                    if cols.get(k):
                                        ws.cell(row=r, column=cols[k], value=v)
                                clear_yellow_highlight(ws, r)
                                wb_touch = True
                                no_change_highlights.append({
                                    "sheet": b["sheet_name"],
                                    "row": r,
                                    "ip": sw_ip,
                                    "port": pt,
                                    "vlan": tg,
                                })
                            continue

                        if args.dry_run:
                            log(f"[DRY-RUN] Row {r} | {sw_ip} | Port {pt} | {cur_v} -> {tg}")
                            continue

                        if not confirm_change(args.safe, sw_ip, pt, cur_v, tg, r, b["sheet_name"]):
                            dt = get_port_live_details(conn, e["port"])
                            stats["dec"] += 1
                            log(f"[SKIPPED] Row {r} declined")
                            write_result_columns(
                                ws,
                                r,
                                cols,
                                switch_ip=sw_ip,
                                port=pt,
                                vlan=cur_v,
                                mac=dt["mac"],
                                ip=dt["ip"],
                                checked_at=now_str(),
                                notes=f"Current VLAN: {cur_v}"
                            )
                            wb_touch = True
                            continue

                        approved_changes.append({
                            "row_idx": r,
                            "port": pt,
                            "target_vlan": tg,
                            "old_vlan": cur_v,
                            "cell": e["cell"]
                        })

                    for group in build_vlan_change_groups(approved_changes):
                        rows = ", ".join(str(item["row_idx"]) for item in group["items"])

                        try:
                            log(
                                f"[*] Applying change on {sw_ip} {group['interface']}: "
                                f"VLAN {group['target_vlan']} (row(s): {rows})"
                            )
                            apply_vlan_change(conn, group["interface"], group["target_vlan"])
                            pending.extend(group["items"])
                        except Exception as exc:
                            for item in group["items"]:
                                stats["fail"] += 1
                                log(f"[FAILED] Row {item['row_idx']} apply error: {exc}")
                                write_result_columns(
                                    ws,
                                    item["row_idx"],
                                    cols,
                                    switch_ip=sw_ip,
                                    port=item["port"],
                                    vlan=item["old_vlan"],
                                    mac="Unknown",
                                    ip="Unknown",
                                    checked_at=now_str(),
                                    notes=f"Error: {str(exc)[:50]}"
                                )
                                wb_touch = True

                    if pending:
                        log(f"[*] Verifying {len(pending)} changed row(s) on {sw_ip}")

                        post_v_map = {}
                        for l in conn.send_command("show int br", read_timeout=60).splitlines():
                            m = re.match(r"^\s*(\d+/\d+/\d+)\s+(\S+)", l.rstrip())
                            if m:
                                post_v_map[m.group(1).strip()] = m.group(2).strip()

                        for item in pending:
                            v_fin = post_v_map.get(item["port"], "Unknown")
                            dt = get_port_live_details(conn, item["port"])

                            if v_fin == item["target_vlan"]:
                                stats["chg"] += 1
                                log(f"[DONE] Row {item['row_idx']} verified successfully")
                                write_result_columns(
                                    ws,
                                    item["row_idx"],
                                    cols,
                                    switch_ip=sw_ip,
                                    port=item["port"],
                                    vlan=v_fin,
                                    mac=dt["mac"],
                                    ip=dt["ip"],
                                    checked_at=now_str()
                                )
                                clear_yellow_highlight(ws, item["row_idx"])
                                summary.append({
                                    "sheet": b["sheet_name"],
                                    "ip": sw_ip,
                                    "port": item["port"],
                                    "target_vlan": item["target_vlan"],
                                    "old_vlan": item["old_vlan"],
                                    "changed_by": RUN_ACTOR,
                                    "changed_at": now_str(),
                                    "source": RUN_SOURCE
                                })
                            else:
                                stats["fail"] += 1
                                log(f"[FAILED] Row {item['row_idx']} verify failed. Live VLAN is {v_fin}")
                                write_result_columns(
                                    ws,
                                    item["row_idx"],
                                    cols,
                                    switch_ip=sw_ip,
                                    port=item["port"],
                                    vlan=v_fin,
                                    mac=dt["mac"],
                                    ip=dt["ip"],
                                    checked_at=now_str(),
                                    notes="Failed to change."
                                )

                            wb_touch = True

            except Exception as exc:
                log(f"[!] Switch connection/processing failure on {sw_ip}: {exc}")
                stats["fail"] += len(entries)
                if not args.dry_run:
                    for e in entries:
                        write_result_columns(
                            ws,
                            e["row_idx"],
                            cols,
                            switch_ip=sw_ip,
                            port=e["port"],
                            vlan="Unknown",
                            mac="Error",
                            ip="Error",
                            checked_at=now_str(),
                            notes="Switch connection failed"
                        )
                    wb_touch = True

    if args.dry_run:
        log("[*] Dry run complete")
        sys.exit(0)

    if wb_touch:
        if os.path.getmtime(DEFAULT_PATH) != start_mtime:
            send_teams_notification("CRITICAL", f"Conflict! Opened by {get_lock_owner(DEFAULT_PATH)}")
            sys.exit(1)
        try:
            wb.save(DEFAULT_PATH)
            log(f"[+] Workbook saved successfully: {DEFAULT_PATH}")
        except Exception as e:
            log(f"[!] Save Error: {e}")
            sys.exit(1)

    st = "CRITICAL" if stats["fail"] > 0 else "SUCCESS" if stats["chg"] > 0 else "INFO"
    log(
        f"[*] Final summary | Checked: {stats['chk']} | Already OK: {stats['ok']} | "
        f"Changed: {stats['chg']} | Failed: {stats['fail']} | Declined: {stats['dec']}"
    )

    final_message = f"Execution completed. Changed: {stats['chg']}, Failed: {stats['fail']}"
    if no_change_highlights:
        row_notes = [
            f"{e['sheet']} row {e['row']} ({e['port']} VLAN {e['vlan']})"
            for e in no_change_highlights
        ]
        shown = ", ".join(row_notes[:20])
        if len(row_notes) > 20:
            shown = f"{shown}, and {len(row_notes) - 20} more"
        log(f"[*] No change needed on highlighted rows: {shown}")
        final_message = f"{final_message}\n\nNo change needed on highlighted rows: {shown}"

    send_teams_notification(
        st,
        final_message,
        summary if stats["chg"] > 0 else None
    )

if __name__ == "__main__":
    main()
