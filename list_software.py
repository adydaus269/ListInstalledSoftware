# save as list_software.py
import winreg
from datetime import datetime
import csv
import os
import sys
from collections import OrderedDict

REG_PATHS = [
    (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"),
    (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"),
    (winreg.HKEY_CURRENT_USER,  r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"),
]

def _read_str_value(key, name, default=""):
    try:
        val, _ = winreg.QueryValueEx(key, name)
        if isinstance(val, str):
            return val.strip()
        return str(val)
    except FileNotFoundError:
        return default
    except OSError:
        return default

def _parse_install_date(raw: str) -> str:
    s = (raw or "").strip()
    if not s:
        return ""
    if len(s) == 8 and s.isdigit():
        # Try YYYYMMDD
        try:
            return datetime.strptime(s, "%Y%m%d").strftime("%Y-%m-%d")
        except ValueError:
            pass
    # Common fallbacks
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y"):
        try:
            return datetime.strptime(s, fmt).strftime("%Y-%m-%d")
        except ValueError:
            continue
    return s

def get_installed_software():
    rows = []

    for root, path in REG_PATHS:
        try:
            with winreg.OpenKey(root, path) as reg_key:
                subkey_count = winreg.QueryInfoKey(reg_key)[0]
                for i in range(subkey_count):
                    try:
                        subkey_name = winreg.EnumKey(reg_key, i)
                        with winreg.OpenKey(root, f"{path}\\{subkey_name}") as sk:
                            name = _read_str_value(sk, "DisplayName")
                            if not name:
                                name = _read_str_value(sk, "QuietDisplayName")
                            if not name:
                                continue

                            version    = _read_str_value(sk, "DisplayVersion")
                            publisher  = _read_str_value(sk, "Publisher")
                            raw_date   = _read_str_value(sk, "InstallDate")
                            inst_date  = _parse_install_date(raw_date)
                            installloc = _read_str_value(sk, "InstallLocation")
                            uninst     = _read_str_value(sk, "UninstallString")

                            rows.append({
                                "Software Name": name,
                                "Version": version,
                                "Publisher": publisher,
                                "Install Date": inst_date,
                                "Install Location": installloc,
                                "Uninstall String": uninst,
                            })
                    except OSError:
                        continue
        except FileNotFoundError:
            continue

    # Dedup
    from collections import OrderedDict
    unique = OrderedDict()
    for r in rows:
        key = (r["Software Name"], r["Version"], r["Publisher"], r["Install Date"])
        if key not in unique:
            unique[key] = r
    return list(unique.values())

def default_output_dir():
    docs = os.path.join(os.path.expanduser("~"), "Documents")
    return docs if os.path.isdir(docs) else os.getcwd()

def write_csv(path, data):
    fieldnames = [
        "Software Name", "Version", "Publisher", "Install Date",
        "Install Location", "Uninstall String"
    ]
    with open(path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(data)

def write_txt(path, data):
    def safe(s): return (s or "").replace("\n", " ").strip()
    with open(path, "w", encoding="utf-8") as f:
        header = f"{'Software Name':60}  {'Version':20}  {'Publisher':30}  {'Install Date':12}"
        f.write(header + "\n")
        f.write("-" * len(header) + "\n")
        for r in sorted(data, key=lambda x: (x["Software Name"].lower(), x.get("Version") or "")):
            f.write(f"{safe(r['Software Name']):60}  {safe(r['Version']):20}  "
                    f"{safe(r['Publisher']):30}  {safe(r['Install Date']):12}\n")
        f.write("\n\n# Extra fields\n")
        for r in data:
            f.write(f"\n[{safe(r['Software Name'])}]"
                    f"\nInstall Location: {safe(r['Install Location'])}"
                    f"\nUninstall String: {safe(r['Uninstall String'])}\n")

def main():
    try:
        print("Collecting installed software…")
        data = get_installed_software()
        print(f"Found {len(data)} programs.")
        fmt = input("Save format (csv/txt) [csv]: ").strip().lower() or "csv"
        if fmt not in ("csv", "txt"):
            print("Invalid format. Use 'csv' or 'txt'.")
            return 2
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        suggested = f"installed_software_{ts}.{fmt}"
        out_dir = input(f"Output folder [{default_output_dir()}]: ").strip() or default_output_dir()
        os.makedirs(out_dir, exist_ok=True)
        out_name = input(f"Output file name [{suggested}]: ").strip() or suggested
        out_path = os.path.join(out_dir, out_name)
        if fmt == "csv":
            write_csv(out_path, data)
        else:
            write_txt(out_path, data)
        print(f"Saved to: {out_path}")
        return 0
    except Exception as e:
        print(f"ERROR: {e}")
        return 1

if __name__ == "__main__":
    import sys
    sys.exit(main())
