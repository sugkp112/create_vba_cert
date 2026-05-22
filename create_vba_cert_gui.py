import shutil
import subprocess
import sys
import traceback
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox


def timestamp_now():
    return datetime.now().strftime("%Y%m%d_%H%M%S")


def get_script_dir():
    return Path(__file__).resolve().parent


def run_powershell(ps_script, encoding="cp932"):
    return subprocess.run(
        [
            "powershell",
            "-NoProfile",
            "-ExecutionPolicy", "Bypass",
            "-Command", ps_script
        ],
        capture_output=True,
        text=True,
        encoding=encoding,
        errors="replace"
    )


def check_powershell():
    return shutil.which("powershell") is not None


def ps_escape(text: str) -> str:
    return text.replace("'", "''")


def get_existing_cert_info(subject_name: str):
    safe_subject = ps_escape(subject_name)
    ps_script = rf"""
$ErrorActionPreference = "Stop"
$keyword = "{safe_subject}"
$storePath = "Cert:\CurrentUser\My"
$certs = Get-ChildItem -Path $storePath | Where-Object {{ $_.Subject -like ("*" + $keyword + "*") }}
if ($certs) {{
    foreach ($c in $certs) {{
        Write-Output ($storePath + "|" + $c.Subject + "|" + $c.Thumbprint + "|" + $c.NotBefore.ToString("yyyy-MM-dd") + "|" + $c.NotAfter.ToString("yyyy-MM-dd"))
    }}
}}
"""
    result = run_powershell(ps_script)
    items = []
    if result.returncode == 0 and result.stdout:
        for line in result.stdout.splitlines():
            line = line.strip()
            if not line or "|" not in line:
                continue
            parts = line.split("|")
            if len(parts) == 5:
                items.append({
                    "store_path": parts[0].strip(),
                    "subject": parts[1].strip(),
                    "thumbprint": parts[2].strip(),
                    "not_before": parts[3].strip(),
                    "not_after": parts[4].strip(),
                })
    return items, result


def get_existing_cert_files(output_dir: Path, keyword: str):
    items = []

    if not output_dir.exists():
        return items

    keyword_lower = keyword.lower().strip()

    for pattern in ("*.pfx", "*.cer"):
        for file_path in output_dir.glob(pattern):
            name_lower = file_path.name.lower()
            if keyword_lower in name_lower:
                items.append({
                    "file_name": file_path.name,
                    "file_path": str(file_path),
                    "file_size": file_path.stat().st_size,
                    "file_type": file_path.suffix.lower(),
                })

    items.sort(key=lambda x: x["file_name"].lower())
    return items


def create_certificate(
    cert_name: str,
    password: str,
    valid_years: int,
    output_dir: Path,
    remove_old: bool,
):
    current_year = str(datetime.now().year)
    base_file_name = f"ChukyoShikenki_VBA_Cert_{current_year}"
    stamp = timestamp_now()

    pfx_path = output_dir / f"{base_file_name}_{stamp}.pfx"
    cer_path = output_dir / f"{base_file_name}_{stamp}.cer"
    password_txt_path = output_dir / f"{base_file_name}_{stamp}_password.txt"
    log_path = output_dir / f"{base_file_name}_{stamp}_log.txt"

    safe_cert_name = ps_escape(cert_name)
    safe_password = ps_escape(password)
    safe_subject = ps_escape(f"CN={cert_name}")
    safe_pfx = ps_escape(str(pfx_path))
    safe_cer = ps_escape(str(cer_path))

    ps_remove_old = ""
    if remove_old:
        ps_remove_old = rf"""
$oldCerts = Get-ChildItem -Path "Cert:\CurrentUser\My" | Where-Object {{ $_.Subject -eq "{safe_subject}" }}
foreach ($old in $oldCerts) {{
    Remove-Item -Path ("Cert:\CurrentUser\My\" + $old.Thumbprint) -Force
}}
"""

    ps_script = rf"""
$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

{ps_remove_old}

$notBefore = Get-Date
$notAfter = $notBefore.AddYears({valid_years})

$cert = New-SelfSignedCertificate `
    -Subject "CN={safe_cert_name}" `
    -Type CodeSigningCert `
    -CertStoreLocation "Cert:\CurrentUser\My" `
    -NotBefore $notBefore `
    -NotAfter $notAfter

$password = ConvertTo-SecureString -String "{safe_password}" -Force -AsPlainText

Export-PfxCertificate `
    -Cert "Cert:\CurrentUser\My\$($cert.Thumbprint)" `
    -FilePath "{safe_pfx}" `
    -Password $password | Out-Null

Export-Certificate `
    -Cert "Cert:\CurrentUser\My\$($cert.Thumbprint)" `
    -FilePath "{safe_cer}" | Out-Null

Write-Output ("THUMBPRINT=" + $cert.Thumbprint)
Write-Output ("NOT_BEFORE=" + $notBefore.ToString("yyyy-MM-dd"))
Write-Output ("NOT_AFTER=" + $notAfter.ToString("yyyy-MM-dd"))
Write-Output ("PFX_PATH={safe_pfx}")
Write-Output ("CER_PATH={safe_cer}")
"""

    stdout_text = ""
    stderr_text = ""
    thumbprint = ""
    not_before = ""
    not_after = ""

    try:
        result = run_powershell(ps_script, encoding="utf-8")
        stdout_text = result.stdout or ""
        stderr_text = result.stderr or ""

        if result.returncode != 0:
            raise RuntimeError(f"PowerShell exited with code {result.returncode}")

        for line in stdout_text.splitlines():
            line = line.strip()
            if line.startswith("THUMBPRINT="):
                thumbprint = line.split("=", 1)[1].strip()
            elif line.startswith("NOT_BEFORE="):
                not_before = line.split("=", 1)[1].strip()
            elif line.startswith("NOT_AFTER="):
                not_after = line.split("=", 1)[1].strip()

        if not pfx_path.exists():
            raise FileNotFoundError(f"PFX file not found: {pfx_path}")
        if not cer_path.exists():
            raise FileNotFoundError(f"CER file not found: {cer_path}")

        if pfx_path.stat().st_size <= 0:
            raise RuntimeError(f"PFX file is empty: {pfx_path}")
        if cer_path.stat().st_size <= 0:
            raise RuntimeError(f"CER file is empty: {cer_path}")

        password_txt_path.write_text(
            "\n".join([
                f"Certificate Name: {cert_name}",
                f"Thumbprint: {thumbprint}",
                f"Password: {password}",
                f"Valid From: {not_before}",
                f"Valid Until: {not_after}",
                f"PFX File: {pfx_path}",
                f"CER File: {cer_path}",
            ]),
            encoding="utf-8"
        )

        log_path.write_text(
            "\n".join([
                "Status: SUCCESS",
                f"Timestamp: {datetime.now().isoformat()}",
                f"Certificate Name: {cert_name}",
                f"Thumbprint: {thumbprint}",
                f"Valid Years: {valid_years}",
                f"Valid From: {not_before}",
                f"Valid Until: {not_after}",
                f"Remove Old Same-Name Certificates: {remove_old}",
                f"Output Folder: {output_dir}",
                f"PFX File: {pfx_path}",
                f"CER File: {cer_path}",
                f"Password File: {password_txt_path}",
                "",
                "=== PowerShell STDOUT ===",
                stdout_text,
                "",
                "=== PowerShell STDERR ===",
                stderr_text,
            ]),
            encoding="utf-8"
        )

        return {
            "success": True,
            "thumbprint": thumbprint,
            "not_before": not_before,
            "not_after": not_after,
            "pfx_path": pfx_path,
            "cer_path": cer_path,
            "password_txt_path": password_txt_path,
            "log_path": log_path,
        }

    except Exception as e:
        error_text = traceback.format_exc()
        try:
            log_path.write_text(
                "\n".join([
                    "Status: FAILED",
                    f"Timestamp: {datetime.now().isoformat()}",
                    f"Certificate Name: {cert_name}",
                    f"Valid Years: {valid_years}",
                    f"Remove Old Same-Name Certificates: {remove_old}",
                    f"Output Folder: {output_dir}",
                    "",
                    f"Error: {e}",
                    "",
                    "=== Traceback ===",
                    error_text,
                    "",
                    "=== PowerShell STDOUT ===",
                    stdout_text,
                    "",
                    "=== PowerShell STDERR ===",
                    stderr_text,
                ]),
                encoding="utf-8"
            )
        except Exception:
            pass

        return {
            "success": False,
            "error": str(e),
            "log_path": log_path,
        }


class App:
    def __init__(self, root):
        self.root = root
        self.root.title("VBA Certificate Generator")
        self.root.geometry("760x500")
        self.root.resizable(False, False)

        self.current_year = str(datetime.now().year)
        self.default_cert_name = f"ChukyoShikenki VBA Cert {self.current_year}"

        self.cert_name_var = tk.StringVar(value=self.default_cert_name)
        self.password_var = tk.StringVar(value="12345678")
        self.valid_years_var = tk.StringVar(value="10")
        self.output_mode_var = tk.StringVar(value="script")
        self.selected_folder_var = tk.StringVar(value=str(get_script_dir()))
        self.remove_old_var = tk.BooleanVar(value=False)

        self.build_ui()

    def build_ui(self):
        pad_x = 12
        pad_y = 8

        tk.Label(self.root, text="Certificate Name").grid(row=0, column=0, sticky="w", padx=pad_x, pady=pad_y)
        tk.Entry(self.root, textvariable=self.cert_name_var, width=65).grid(row=0, column=1, columnspan=3, sticky="we", padx=pad_x, pady=pad_y)

        tk.Label(self.root, text="PFX Password").grid(row=1, column=0, sticky="w", padx=pad_x, pady=pad_y)
        tk.Entry(self.root, textvariable=self.password_var, width=30).grid(row=1, column=1, sticky="w", padx=pad_x, pady=pad_y)

        tk.Label(self.root, text="Valid Years").grid(row=2, column=0, sticky="w", padx=pad_x, pady=pad_y)
        tk.Entry(self.root, textvariable=self.valid_years_var, width=10).grid(row=2, column=1, sticky="w", padx=pad_x, pady=pad_y)

        tk.Label(self.root, text="Output Folder").grid(row=3, column=0, sticky="nw", padx=pad_x, pady=pad_y)

        frame_mode = tk.Frame(self.root)
        frame_mode.grid(row=3, column=1, columnspan=3, sticky="w", padx=pad_x, pady=pad_y)

        tk.Radiobutton(
            frame_mode,
            text="Use script folder",
            variable=self.output_mode_var,
            value="script",
            command=self.update_output_display
        ).grid(row=0, column=0, sticky="w")

        tk.Radiobutton(
            frame_mode,
            text="Use selected folder",
            variable=self.output_mode_var,
            value="selected",
            command=self.update_output_display
        ).grid(row=0, column=1, sticky="w", padx=(20, 0))

        self.folder_entry = tk.Entry(frame_mode, textvariable=self.selected_folder_var, width=56)
        self.folder_entry.grid(row=1, column=0, columnspan=2, sticky="w", pady=(8, 0))

        self.browse_button = tk.Button(frame_mode, text="Browse...", command=self.choose_folder)
        self.browse_button.grid(row=1, column=2, sticky="w", padx=(8, 0), pady=(8, 0))

        tk.Checkbutton(
            self.root,
            text="Remove old same-name certificates before creating a new one",
            variable=self.remove_old_var
        ).grid(row=4, column=1, columnspan=3, sticky="w", padx=pad_x, pady=pad_y)

        tk.Button(self.root, text="Check Existing Certificates", command=self.check_existing).grid(
            row=5, column=1, sticky="w", padx=pad_x, pady=pad_y
        )

        tk.Button(self.root, text="Create Certificate", command=self.create).grid(
            row=5, column=2, sticky="w", padx=pad_x, pady=pad_y
        )

        tk.Label(self.root, text="Status").grid(row=6, column=0, sticky="nw", padx=pad_x, pady=pad_y)
        self.status_text = tk.Text(self.root, width=88, height=18)
        self.status_text.grid(row=6, column=1, columnspan=3, padx=pad_x, pady=pad_y, sticky="nsew")

        self.update_output_display()

    def choose_folder(self):
        folder = filedialog.askdirectory(initialdir=self.selected_folder_var.get() or str(get_script_dir()))
        if folder:
            self.selected_folder_var.set(folder)
            self.output_mode_var.set("selected")
            self.update_output_display()

    def update_output_display(self):
        if self.output_mode_var.get() == "script":
            self.folder_entry.configure(state="disabled")
            self.browse_button.configure(state="disabled")
        else:
            self.folder_entry.configure(state="normal")
            self.browse_button.configure(state="normal")

    def log(self, text):
        self.status_text.insert("end", text + "\n")
        self.status_text.see("end")
        self.root.update_idletasks()

    def clear_log(self):
        self.status_text.delete("1.0", "end")

    def get_output_dir(self):
        if self.output_mode_var.get() == "script":
            return get_script_dir()
        return Path(self.selected_folder_var.get()).resolve()

    def check_existing(self):
        self.clear_log()

        if not check_powershell():
            messagebox.showerror("Error", "PowerShell was not found on this system.")
            return

        cert_name = self.cert_name_var.get().strip()
        if not cert_name:
            messagebox.showerror("Error", "Certificate name is required.")
            return

        output_dir = self.get_output_dir()

        lines = []
        lines.append("=== Existing Certificate Check ===")
        lines.append("")

        cert_items, result = get_existing_cert_info(cert_name)
        lines.append("[Windows Certificate Store]")
        if result.returncode != 0:
            lines.append("Check failed.")
        elif not cert_items:
            lines.append("No matching certificates found.")
        else:
            lines.append(f"Found {len(cert_items)} matching certificate(s):")
            for idx, item in enumerate(cert_items, start=1):
                lines.append(
                    f"{idx}. {item['subject']} | "
                    f"{item['not_before']} -> {item['not_after']} | "
                    f"{item['thumbprint']}"
                )

        lines.append("")
        lines.append("[Output Folder Files]")
        file_items = get_existing_cert_files(output_dir, cert_name)
        if not file_items:
            lines.append("No matching .pfx/.cer files found.")
        else:
            lines.append(f"Found {len(file_items)} matching file(s):")
            for idx, item in enumerate(file_items, start=1):
                lines.append(
                    f"{idx}. {item['file_name']} | "
                    f"{item['file_size']} bytes"
                )

        lines.append("")
        lines.append(f"Folder: {output_dir}")

        self.log("\n".join(lines))

    def create(self):
        self.clear_log()

        if not check_powershell():
            messagebox.showerror("Error", "PowerShell was not found on this system.")
            return

        cert_name = self.cert_name_var.get().strip()
        password = self.password_var.get()
        valid_years_text = self.valid_years_var.get().strip()

        if not cert_name:
            messagebox.showerror("Error", "Certificate name is required.")
            return

        if not password:
            messagebox.showerror("Error", "PFX password is required.")
            return

        try:
            valid_years = int(valid_years_text)
            if valid_years <= 0:
                raise ValueError
        except ValueError:
            messagebox.showerror("Error", "Valid years must be a positive integer.")
            return

        output_dir = self.get_output_dir()
        output_dir.mkdir(parents=True, exist_ok=True)

        existing_items, _ = get_existing_cert_info(cert_name)
        if existing_items and not self.remove_old_var.get():
            proceed = messagebox.askyesno(
                "Existing Certificates Found",
                "Existing certificates with the same subject were found.\n\n"
                "Do you want to continue without removing them?"
            )
            if not proceed:
                return

        start_lines = [
            "=== Creating Certificate ===",
            "",
            f"Name         : {cert_name}",
            f"Valid Years  : {valid_years}",
            f"Output Folder: {output_dir}",
            f"Remove Old   : {self.remove_old_var.get()}",
            "",
            "Working..."
        ]
        self.log("\n".join(start_lines))

        result = create_certificate(
            cert_name=cert_name,
            password=password,
            valid_years=valid_years,
            output_dir=output_dir,
            remove_old=self.remove_old_var.get(),
        )

        self.clear_log()

        if result["success"]:
            summary_lines = [
                "=== Certificate Created Successfully ===",
                "",
                "[Certificate]",
                f"Name       : {cert_name}",
                f"Thumbprint : {result['thumbprint']}",
                f"Valid From : {result['not_before']}",
                f"Valid Until: {result['not_after']}",
                "",
                "[Output Files]",
                f"PFX       : {Path(result['pfx_path']).name}",
                f"CER       : {Path(result['cer_path']).name}",
                f"Password  : {Path(result['password_txt_path']).name}",
                f"Log       : {Path(result['log_path']).name}",
                "",
                f"Folder    : {output_dir}",
            ]
            self.log("\n".join(summary_lines))

            try:
                subprocess.Popen(["explorer", str(output_dir)])
            except Exception:
                pass

            messagebox.showinfo(
                "Success",
                "Certificate files were created successfully.\n\n"
                f"Thumbprint:\n{result['thumbprint']}"
            )
        else:
            error_lines = [
                "=== Certificate Creation Failed ===",
                "",
                "[Error]",
                f"{result['error']}",
                "",
                "[Log File]",
                f"{result['log_path']}",
            ]
            self.log("\n".join(error_lines))

            messagebox.showerror(
                "Failed",
                "Certificate creation failed.\n\n"
                f"See log file:\n{result['log_path']}"
            )


def main():
    if sys.platform != "win32":
        print("This script only supports Windows.")
        sys.exit(1)

    root = tk.Tk()
    app = App(root)
    root.mainloop()


if __name__ == "__main__":
    main()