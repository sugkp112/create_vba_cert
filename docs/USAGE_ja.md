# VBA Certificate Generator 使用メモ

このツールは、Excel / Word の VBA マクロにデジタル署名するための自己署名証明書を作成する Windows 用ツールです。

## 対象

- 個人利用
- 社内利用
- テスト用 VBA ツール
- Excel `.xlsm` / `.xlsb` のマクロ署名

## 実行方法

```bat
run.bat
```

または：

```bat
python create_vba_cert_gui.py
```

## 注意

生成された `.pfx` ファイルとパスワードファイルは秘密情報です。GitHub にアップロードしないでください。
