#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
VBA Certificate Generator

A small Windows GUI utility for creating a self-signed code-signing
certificate for Excel / Word VBA macro projects.

What it does:
- Creates a self-signed CodeSigningCert in CurrentUser\My.
- Exports a .pfx file for signing VBA projects.
- Exports a .cer file that can be installed as a trusted certificate.
- Writes a log file for troubleshooting.
- Optionally writes a password reminder file.

Requirements:
- Windows
- PowerShell
- Python 3.9+
- No third-party Python packages
"""

from __future__ import annotations

import re
import shutil
import subprocess
import sys
import traceback
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox


APP_TITLE = "VBA Certificate Generator"
DEFAULT_VALID_YEARS = 10
DEFAULT_CERT_NAME_PREFIX = "VBA Code Signing Certificate"
DEFAULT_PFX_PASSWORD = ""
OUTPUT_SUBFOLDER = "generated_certificates"


# -----------------------------
# Data models
# -----------------------------

@dataclass
class CertificateInfo:
    store_path: str
    subject: str
    thumbprint: str
    not_before: str
    not_after: str


@dataclass
class GeneratedCertificate:
    thumbprint: str
    not_before: str
    not_after: str
    pfx_path: Path
    cer_path: Path
    password_txt_path: Path | None
    log_path: Path


@dataclass
class CertificateCreationResult:
    success: bool
    data: GeneratedCertificate | None = None
    error: str = ""
    log_path: Path | None = None


# -----------------------------
# General helpers
# -----------------------------

def timestamp_now() -> str:
    return datetime.now().strftime("%Y%m%d_%H%M%S")


def get_script_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def get_default_output_dir() -> Path:
    return get_script_dir() / OUTPUT_SUBFOLDER


def check_powershell() -> bool:
    return shutil.which("powershell") is not None


def ps_single_quote(text: str) -> str:
    """Escape a Python string for a PowerShell single-quoted string."""
    return text.replace("'", "''")


def safe_file_part(text: str, fallback: str = "VBA_Cert") -> str:
    """
    Convert certificate name into a safe filename part.
    This avoids Windows filename-invalid characters.
    """
    text = text.strip()
    text = re.sub(r"[\\/:*?\"<>|]+", "_", text)
    text = re.sub(r"\s+", "_", text)
    text = re.sub(r"_+", "_", text).strip("._ ")
    return text or fallback


def run_powershell(ps_script: str, encoding: str = "utf-8") -> subprocess.CompletedProcess:
    return subprocess.run(
        [
            "powershell",
            "-NoProfile",
            "-ExecutionPolicy",
            "Bypass",
            "-Command",
            ps_script,
        ],
        capture_output=True,
        text=True,
        encoding=encoding,
        errors="replace",
    )


# -----------------------------
# Certificate operations
# -----------------------------

def get_existing_cert_info(subject_keyword: str) -> tuple[list[CertificateInfo], subprocess.CompletedProcess]:
    keyword = ps_single_quote(subject_keyword)
    ps_script = f"""
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

$keyword = '{keyword}'
$storePath = 'Cert:\\CurrentUser\\My'
$certs = Get-ChildItem -Path $storePath | Where-Object {{ $_.Subject -like ('*' + $keyword + '*') }}

if ($certs) {{
    foreach ($c in $certs) {{
        Write-Output ($storePath + '|' + $c.Subject + '|' + $c.Thumbprint + '|' + $c.NotBefore.ToString('yyyy-MM-dd') + '|' + $c.NotAfter.ToString('yyyy-MM-dd'))
    }}
}}
"""
    result = run_powershell(ps_script)
    items: list[CertificateInfo] = []

    if result.returncode == 0 and result.stdout:
        for line in result.stdout.splitlines():
            line = line.strip()
            if not line or "|" not in line:
                continue
            parts = line.split("|")
            if len(parts) == 5:
                items.append(
                    CertificateInfo(
                        store_path=parts[0].strip(),
                        subject=parts[1].strip(),
                        thumbprint=parts[2].strip(),
                        not_before=parts[3].strip(),
                        not_after=parts[4].strip(),
                    )
                )

    return items, result


def get_existing_cert_files(output_dir: Path, keyword: str) -> list[dict[str, object]]:
    items: list[dict[str, object]] = []
    if not output_dir.exists():
        return items

    keyword_lower = keyword.lower().strip()
    for pattern in ("*.pfx", "*.cer", "*_password.txt", "*_log.txt"):
        for file_path in output_dir.glob(pattern):
            if keyword_lower and keyword_lower not in file_path.name.lower():
                continue
            items.append(
                {
                    "file_name": file_path.name,
                    "file_path": str(file_path),
                    "file_size": file_path.stat().st_size,
                    "file_type": file_path.suffix.lower(),
                    "modified": datetime.fromtimestamp(file_path.stat().st_mtime).strftime("%Y-%m-%d %H:%M:%S"),
                }
            )

    items.sort(key=lambda x: str(x["file_name"]).lower())
    return items


def create_certificate(
    cert_name: str,
    password: str,
    valid_years: int,
    output_dir: Path,
    remove_old: bool,
    save_password_file: bool,
) -> CertificateCreationResult:
    output_dir.mkdir(parents=True, exist_ok=True)

    file_name_part = safe_file_part(cert_name)
    stamp = timestamp_now()

    pfx_path = output_dir / f"{file_name_part}_{stamp}.pfx"
    cer_path = output_dir / f"{file_name_part}_{stamp}.cer"
    password_txt_path = output_dir / f"{file_name_part}_{stamp}_password.txt"
    log_path = output_dir / f"{file_name_part}_{stamp}_log.txt"

    safe_cert_name = ps_single_quote(cert_name)
    safe_subject = ps_single_quote(f"CN={cert_name}")
    safe_password = ps_single_quote(password)
    safe_pfx = ps_single_quote(str(pfx_path))
    safe_cer = ps_single_quote(str(cer_path))

    ps_remove_old = ""
    if remove_old:
        ps_remove_old = f"""
$oldCerts = Get-ChildItem -Path 'Cert:\\CurrentUser\\My' | Where-Object {{ $_.Subject -eq '{safe_subject}' }}
foreach ($old in $oldCerts) {{
    Remove-Item -Path ('Cert:\\CurrentUser\\My\\' + $old.Thumbprint) -Force
}}
"""

    ps_script = f"""
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

{ps_remove_old}

$notBefore = Get-Date
$notAfter = $notBefore.AddYears({valid_years})

$cert = New-SelfSignedCertificate `
    -Subject 'CN={safe_cert_name}' `
    -Type CodeSigningCert `
    -CertStoreLocation 'Cert:\\CurrentUser\\My' `
    -NotBefore $notBefore `
    -NotAfter $notAfter

$password = ConvertTo-SecureString -String '{safe_password}' -Force -AsPlainText

Export-PfxCertificate `
    -Cert ('Cert:\\CurrentUser\\My\\' + $cert.Thumbprint) `
    -FilePath '{safe_pfx}' `
    -Password $password | Out-Null

Export-Certificate `
    -Cert ('Cert:\\CurrentUser\\My\\' + $cert.Thumbprint) `
    -FilePath '{safe_cer}' | Out-Null

Write-Output ('THUMBPRINT=' + $cert.Thumbprint)
Write-Output ('NOT_BEFORE=' + $notBefore.ToString('yyyy-MM-dd'))
Write-Output ('NOT_AFTER=' + $notAfter.ToString('yyyy-MM-dd'))
Write-Output ('PFX_PATH={safe_pfx}')
Write-Output ('CER_PATH={safe_cer}')
"""

    stdout_text = ""
    stderr_text = ""

    try:
        result = run_powershell(ps_script)
        stdout_text = result.stdout or ""
        stderr_text = result.stderr or ""

        if result.returncode != 0:
            raise RuntimeError(f"PowerShell exited with code {result.returncode}")

        thumbprint = ""
        not_before = ""
        not_after = ""

        for line in stdout_text.splitlines():
            line = line.strip()
            if line.startswith("THUMBPRINT="):
                thumbprint = line.split("=", 1)[1].strip()
            elif line.startswith("NOT_BEFORE="):
                not_before = line.split("=", 1)[1].strip()
            elif line.startswith("NOT_AFTER="):
                not_after = line.split("=", 1)[1].strip()

        if not thumbprint:
            raise RuntimeError("Could not read certificate thumbprint from PowerShell output.")

        if not pfx_path.exists() or pfx_path.stat().st_size <= 0:
            raise FileNotFoundError(f"PFX file was not created correctly: {pfx_path}")
        if not cer_path.exists() or cer_path.stat().st_size <= 0:
            raise FileNotFoundError(f"CER file was not created correctly: {cer_path}")

        saved_password_path: Path | None = None
        if save_password_file:
            password_txt_path.write_text(
                "\n".join(
                    [
                        "WARNING: This file contains the PFX password.",
                        "Do not upload it to GitHub or share it publicly.",
                        "",
                        f"Certificate Name: {cert_name}",
                        f"Thumbprint: {thumbprint}",
                        f"Password: {password}",
                        f"Valid From: {not_before}",
                        f"Valid Until: {not_after}",
                        f"PFX File: {pfx_path}",
                        f"CER File: {cer_path}",
                    ]
                ),
                encoding="utf-8",
            )
            saved_password_path = password_txt_path

        log_path.write_text(
            "\n".join(
                [
                    "Status: SUCCESS",
                    f"Timestamp: {datetime.now().isoformat()}",
                    f"Certificate Name: {cert_name}",
                    f"Thumbprint: {thumbprint}",
                    f"Valid Years: {valid_years}",
                    f"Valid From: {not_before}",
                    f"Valid Until: {not_after}",
                    f"Remove Old Same-Name Certificates: {remove_old}",
                    f"Save Password Reminder File: {save_password_file}",
                    f"Output Folder: {output_dir}",
                    f"PFX File: {pfx_path}",
                    f"CER File: {cer_path}",
                    f"Password File: {saved_password_path or '(not saved)'}",
                    "",
                    "=== PowerShell STDOUT ===",
                    stdout_text,
                    "",
                    "=== PowerShell STDERR ===",
                    stderr_text,
                ]
            ),
            encoding="utf-8",
        )

        return CertificateCreationResult(
            success=True,
            data=GeneratedCertificate(
                thumbprint=thumbprint,
                not_before=not_before,
                not_after=not_after,
                pfx_path=pfx_path,
                cer_path=cer_path,
                password_txt_path=saved_password_path,
                log_path=log_path,
            ),
        )

    except Exception as e:
        error_text = traceback.format_exc()
        try:
            log_path.write_text(
                "\n".join(
                    [
                        "Status: FAILED",
                        f"Timestamp: {datetime.now().isoformat()}",
                        f"Certificate Name: {cert_name}",
                        f"Valid Years: {valid_years}",
                        f"Remove Old Same-Name Certificates: {remove_old}",
                        f"Save Password Reminder File: {save_password_file}",
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
                    ]
                ),
                encoding="utf-8",
            )
        except Exception:
            pass

        return CertificateCreationResult(success=False, error=str(e), log_path=log_path)


# -----------------------------
# GUI
# -----------------------------

class App:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry("820x560")
        self.root.minsize(820, 560)

        current_year = datetime.now().year
        self.default_cert_name = f"{DEFAULT_CERT_NAME_PREFIX} {current_year}"

        self.cert_name_var = tk.StringVar(value=self.default_cert_name)
        self.password_var = tk.StringVar(value=DEFAULT_PFX_PASSWORD)
        self.valid_years_var = tk.StringVar(value=str(DEFAULT_VALID_YEARS))
        self.output_mode_var = tk.StringVar(value="script")
        self.selected_folder_var = tk.StringVar(value=str(get_default_output_dir()))
        self.remove_old_var = tk.BooleanVar(value=False)
        self.save_password_file_var = tk.BooleanVar(value=False)

        self._build_ui()
        self._update_output_display()

    def _build_ui(self) -> None:
        self.root.columnconfigure(1, weight=1)
        self.root.rowconfigure(8, weight=1)

        pad_x = 12
        pad_y = 7

        tk.Label(self.root, text="Certificate Name").grid(row=0, column=0, sticky="w", padx=pad_x, pady=pad_y)
        tk.Entry(self.root, textvariable=self.cert_name_var, width=70).grid(
            row=0, column=1, columnspan=3, sticky="we", padx=pad_x, pady=pad_y
        )

        tk.Label(self.root, text="PFX Password").grid(row=1, column=0, sticky="w", padx=pad_x, pady=pad_y)
        tk.Entry(self.root, textvariable=self.password_var, width=32, show="*").grid(
            row=1, column=1, sticky="w", padx=pad_x, pady=pad_y
        )

        tk.Button(self.root, text="Show / Hide", command=self._toggle_password).grid(
            row=1, column=2, sticky="w", padx=pad_x, pady=pad_y
        )

        tk.Label(self.root, text="Valid Years").grid(row=2, column=0, sticky="w", padx=pad_x, pady=pad_y)
        tk.Entry(self.root, textvariable=self.valid_years_var, width=10).grid(
            row=2, column=1, sticky="w", padx=pad_x, pady=pad_y
        )

        tk.Label(self.root, text="Output Folder").grid(row=3, column=0, sticky="nw", padx=pad_x, pady=pad_y)

        frame_mode = tk.Frame(self.root)
        frame_mode.grid(row=3, column=1, columnspan=3, sticky="we", padx=pad_x, pady=pad_y)
        frame_mode.columnconfigure(0, weight=1)

        tk.Radiobutton(
            frame_mode,
            text=f"Use default folder: ./{OUTPUT_SUBFOLDER}",
            variable=self.output_mode_var,
            value="script",
            command=self._update_output_display,
        ).grid(row=0, column=0, sticky="w")

        tk.Radiobutton(
            frame_mode,
            text="Use selected folder",
            variable=self.output_mode_var,
            value="selected",
            command=self._update_output_display,
        ).grid(row=0, column=1, sticky="w", padx=(20, 0))

        self.folder_entry = tk.Entry(frame_mode, textvariable=self.selected_folder_var, width=72)
        self.folder_entry.grid(row=1, column=0, columnspan=2, sticky="we", pady=(8, 0))

        self.browse_button = tk.Button(frame_mode, text="Browse...", command=self._choose_folder)
        self.browse_button.grid(row=1, column=2, sticky="w", padx=(8, 0), pady=(8, 0))

        tk.Checkbutton(
            self.root,
            text="Remove old same-name certificates before creating a new one",
            variable=self.remove_old_var,
        ).grid(row=4, column=1, columnspan=3, sticky="w", padx=pad_x, pady=pad_y)

        tk.Checkbutton(
            self.root,
            text="Save password reminder file next to the certificate files",
            variable=self.save_password_file_var,
        ).grid(row=5, column=1, columnspan=3, sticky="w", padx=pad_x, pady=pad_y)

        button_frame = tk.Frame(self.root)
        button_frame.grid(row=6, column=1, columnspan=3, sticky="w", padx=pad_x, pady=pad_y)

        tk.Button(button_frame, text="Check Existing Certificates", command=self.check_existing).grid(
            row=0, column=0, sticky="w", padx=(0, 8)
        )
        tk.Button(button_frame, text="Create Certificate", command=self.create).grid(
            row=0, column=1, sticky="w", padx=(0, 8)
        )
        tk.Button(button_frame, text="Open Output Folder", command=self.open_output_folder).grid(
            row=0, column=2, sticky="w", padx=(0, 8)
        )
        tk.Button(button_frame, text="Clear Log", command=self.clear_log).grid(row=0, column=3, sticky="w")

        note = (
            "Note: This tool creates a self-signed certificate for VBA macro signing. "
            "Generated .pfx and password files should not be uploaded to GitHub."
        )
        tk.Label(self.root, text=note, fg="#475569", wraplength=760, justify="left").grid(
            row=7, column=0, columnspan=4, sticky="w", padx=pad_x, pady=(0, 4)
        )

        tk.Label(self.root, text="Status").grid(row=8, column=0, sticky="nw", padx=pad_x, pady=pad_y)

        text_frame = tk.Frame(self.root)
        text_frame.grid(row=8, column=1, columnspan=3, padx=pad_x, pady=pad_y, sticky="nsew")
        text_frame.rowconfigure(0, weight=1)
        text_frame.columnconfigure(0, weight=1)

        self.status_text = tk.Text(text_frame, width=92, height=18, wrap="word")
        self.status_text.grid(row=0, column=0, sticky="nsew")

        scrollbar = tk.Scrollbar(text_frame, orient="vertical", command=self.status_text.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.status_text.configure(yscrollcommand=scrollbar.set)

    def _toggle_password(self) -> None:
        current = self.root.focus_get()
        # Find the password entry by scanning children.
        for child in self.root.winfo_children():
            if isinstance(child, tk.Entry) and str(child.cget("textvariable")) == str(self.password_var):
                child.configure(show="" if child.cget("show") == "*" else "*")
                child.focus_set()
                return
        if current:
            current.focus_set()

    def _choose_folder(self) -> None:
        folder = filedialog.askdirectory(initialdir=self.selected_folder_var.get() or str(get_default_output_dir()))
        if folder:
            self.selected_folder_var.set(folder)
            self.output_mode_var.set("selected")
            self._update_output_display()

    def _update_output_display(self) -> None:
        if self.output_mode_var.get() == "script":
            self.selected_folder_var.set(str(get_default_output_dir()))
            self.folder_entry.configure(state="disabled")
            self.browse_button.configure(state="disabled")
        else:
            self.folder_entry.configure(state="normal")
            self.browse_button.configure(state="normal")

    def log(self, text: str) -> None:
        self.status_text.insert("end", text + "\n")
        self.status_text.see("end")
        self.root.update_idletasks()

    def clear_log(self) -> None:
        self.status_text.delete("1.0", "end")

    def get_output_dir(self) -> Path:
        if self.output_mode_var.get() == "script":
            return get_default_output_dir()
        return Path(self.selected_folder_var.get()).resolve()

    def open_output_folder(self) -> None:
        output_dir = self.get_output_dir()
        output_dir.mkdir(parents=True, exist_ok=True)
        try:
            subprocess.Popen(["explorer", str(output_dir)])
        except Exception as exc:
            messagebox.showerror("Error", f"Could not open folder:\n{exc}")

    def _validate_input(self) -> tuple[str, str, int] | None:
        cert_name = self.cert_name_var.get().strip()
        password = self.password_var.get()
        valid_years_text = self.valid_years_var.get().strip()

        if not cert_name:
            messagebox.showerror("Error", "Certificate name is required.")
            return None

        if not password:
            messagebox.showerror("Error", "PFX password is required.")
            return None

        if len(password) < 8:
            proceed = messagebox.askyesno(
                "Weak Password",
                "The PFX password is shorter than 8 characters.\n\n"
                "Do you want to continue?"
            )
            if not proceed:
                return None

        try:
            valid_years = int(valid_years_text)
            if not 1 <= valid_years <= 50:
                raise ValueError
        except ValueError:
            messagebox.showerror("Error", "Valid years must be an integer between 1 and 50.")
            return None

        return cert_name, password, valid_years

    def check_existing(self) -> None:
        self.clear_log()

        if not check_powershell():
            messagebox.showerror("Error", "PowerShell was not found on this system.")
            return

        cert_name = self.cert_name_var.get().strip()
        if not cert_name:
            messagebox.showerror("Error", "Certificate name is required.")
            return

        output_dir = self.get_output_dir()

        lines: list[str] = [
            "=== Existing Certificate Check ===",
            "",
            f"Search Keyword: {cert_name}",
            f"Output Folder : {output_dir}",
            "",
        ]

        cert_items, result = get_existing_cert_info(cert_name)

        lines.append("[Windows Certificate Store: CurrentUser\\My]")
        if result.returncode != 0:
            lines.append("Check failed.")
            lines.append(result.stderr or result.stdout or "(no PowerShell output)")
        elif not cert_items:
            lines.append("No matching certificates found.")
        else:
            lines.append(f"Found {len(cert_items)} matching certificate(s):")
            for idx, item in enumerate(cert_items, start=1):
                lines.append(
                    f"{idx}. {item.subject} | "
                    f"{item.not_before} -> {item.not_after} | "
                    f"{item.thumbprint}"
                )

        lines.append("")
        lines.append("[Output Folder Files]")
        file_items = get_existing_cert_files(output_dir, safe_file_part(cert_name))
        if not file_items:
            lines.append("No matching output files found.")
        else:
            lines.append(f"Found {len(file_items)} matching file(s):")
            for idx, item in enumerate(file_items, start=1):
                lines.append(
                    f"{idx}. {item['file_name']} | "
                    f"{item['file_size']} bytes | "
                    f"{item['modified']}"
                )

        self.log("\n".join(lines))

    def create(self) -> None:
        self.clear_log()

        if sys.platform != "win32":
            messagebox.showerror("Error", "This tool only supports Windows.")
            return

        if not check_powershell():
            messagebox.showerror("Error", "PowerShell was not found on this system.")
            return

        validated = self._validate_input()
        if not validated:
            return

        cert_name, password, valid_years = validated
        output_dir = self.get_output_dir()
        output_dir.mkdir(parents=True, exist_ok=True)

        existing_items, _ = get_existing_cert_info(cert_name)
        if existing_items and not self.remove_old_var.get():
            proceed = messagebox.askyesno(
                "Existing Certificates Found",
                "Existing certificates with a matching subject were found.\n\n"
                "Do you want to continue without removing them?"
            )
            if not proceed:
                return

        self.log(
            "\n".join(
                [
                    "=== Creating Certificate ===",
                    "",
                    f"Name              : {cert_name}",
                    f"Valid Years       : {valid_years}",
                    f"Output Folder     : {output_dir}",
                    f"Remove Old        : {self.remove_old_var.get()}",
                    f"Save Password File: {self.save_password_file_var.get()}",
                    "",
                    "Working...",
                ]
            )
        )

        result = create_certificate(
            cert_name=cert_name,
            password=password,
            valid_years=valid_years,
            output_dir=output_dir,
            remove_old=self.remove_old_var.get(),
            save_password_file=self.save_password_file_var.get(),
        )

        self.clear_log()

        if result.success and result.data:
            data = result.data
            self.log(
                "\n".join(
                    [
                        "=== Certificate Created Successfully ===",
                        "",
                        "[Certificate]",
                        f"Name       : {cert_name}",
                        f"Thumbprint : {data.thumbprint}",
                        f"Valid From : {data.not_before}",
                        f"Valid Until: {data.not_after}",
                        "",
                        "[Output Files]",
                        f"PFX       : {data.pfx_path.name}",
                        f"CER       : {data.cer_path.name}",
                        f"Password  : {data.password_txt_path.name if data.password_txt_path else '(not saved)'}",
                        f"Log       : {data.log_path.name}",
                        "",
                        f"Folder    : {output_dir}",
                        "",
                        "[Next Steps]",
                        "1. Import the .pfx file if you need to sign VBA projects on another PC.",
                        "2. Use the .cer file to trust this certificate on your own/internal PCs.",
                        "3. Do not upload generated .pfx or password files to GitHub.",
                    ]
                )
            )

            try:
                subprocess.Popen(["explorer", str(output_dir)])
            except Exception:
                pass

            messagebox.showinfo(
                "Success",
                "Certificate files were created successfully.\n\n"
                f"Thumbprint:\n{data.thumbprint}"
            )
        else:
            self.log(
                "\n".join(
                    [
                        "=== Certificate Creation Failed ===",
                        "",
                        "[Error]",
                        result.error or "(unknown error)",
                        "",
                        "[Log File]",
                        str(result.log_path or "(not available)"),
                    ]
                )
            )

            messagebox.showerror(
                "Failed",
                "Certificate creation failed.\n\n"
                f"See log file:\n{result.log_path}"
            )


def main() -> None:
    if sys.platform != "win32":
        print("This script only supports Windows.")
        sys.exit(1)

    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
