# VBA Certificate Generator

A small Windows GUI tool for creating a **self-signed code-signing certificate** for Excel / Word VBA macro projects.

This tool is intended for personal or internal-company VBA tools, for example `.xlsm` / `.xlsb` files that contain VBA macros.

## What this tool does

- Creates a self-signed VBA code-signing certificate in the current Windows user's certificate store.
- Exports a `.pfx` file for signing VBA projects.
- Exports a `.cer` file for installing/trusting the certificate on your own or internal PCs.
- Writes a log file for troubleshooting.
- Optionally writes a password reminder file.

## Requirements

- Windows
- PowerShell
- Python 3.9 or later
- No third-party Python packages

## How to run

Double-click:

```bat
run.bat
```

Or run from PowerShell / Command Prompt:

```bat
python create_vba_cert_gui.py
```

## Recommended settings

- Certificate Name: `VBA Code Signing Certificate 2026` or your organization/tool name
- PFX Password: use a private password, preferably at least 8 characters
- Valid Years: `10`
- Output Folder: default folder is usually fine

## Output files

The tool creates files under:

```text
generated_certificates/
```

Example:

```text
VBA_Code_Signing_Certificate_2026_20260522_103000.pfx
VBA_Code_Signing_Certificate_2026_20260522_103000.cer
VBA_Code_Signing_Certificate_2026_20260522_103000_log.txt
```

If you enable the password reminder option, it also creates:

```text
VBA_Code_Signing_Certificate_2026_20260522_103000_password.txt
```

## How to use the certificate in Excel VBA

1. Open your Excel macro file, such as `.xlsm`.
2. Press `Alt + F11` to open the VBA editor.
3. Open `Tools` → `Digital Signature`.
4. Choose the certificate created by this tool.
5. Save the workbook.
6. Close and reopen Excel to test the signature.

## Important security notes

- This is a **self-signed certificate**, not a commercial certificate from a public CA.
- It is suitable for personal use, testing, and internal company workflows.
- Do not upload generated `.pfx` files, password files, or private keys to GitHub.
- The included `.gitignore` is configured to avoid committing generated certificate files.

## Japanese note

このツールは、Excel / Word の VBA マクロにデジタル署名するための自己署名証明書を作成する Windows 用 GUI ツールです。個人利用または社内利用を想定しています。生成された `.pfx` とパスワードファイルは GitHub にアップロードしないでください。

## Chinese note

这个工具用于为 Excel / Word 的 VBA 宏项目创建自签名代码签名证书。适合个人电脑、公司内部工具或测试阶段使用。生成的 `.pfx` 文件和密码文件不要上传到 GitHub。
