# 使用说明：VBA 代码签名证书生成工具

## 这个工具适合什么情况？

当你制作了带 VBA 宏的 Excel / Word 文件，例如：

- `.xlsm`
- `.xlsb`
- `.docm`

并且希望给 VBA 项目添加数字签名时，可以使用这个工具。

它尤其适合：

- 自己电脑上使用的 Excel VBA 工具
- 公司内部少量电脑使用的 VBA 工具
- 测试阶段的宏工具
- 希望减少“宏已被禁用”提示的内部文件

## 不适合什么情况？

它不适合直接作为商业公开发行证书使用。因为它生成的是自签名证书，不是公开 CA 机构签发的商业代码签名证书。

## 基本流程

### 1. 运行工具

双击：

```text
run.bat
```

或者运行：

```text
python create_vba_cert_gui.py
```

### 2. 创建证书

填写：

```text
Certificate Name  证书名称
PFX Password      PFX 密码
Valid Years       有效年限
Output Folder     输出文件夹
```

然后点击：

```text
Create Certificate
```

### 3. 在 Excel VBA 中签名

1. 打开 Excel 文件。
2. 按 `Alt + F11`。
3. 菜单选择 `Tools` → `Digital Signature`。
4. 选择刚才创建的证书。
5. 保存 Excel 文件。

### 4. 信任证书

如果要让同一台电脑或内部电脑信任这个证书，可以使用 `.cer` 文件导入到受信任证书区域。

## 文件安全

不要上传这些文件到 GitHub：

```text
*.pfx
*_password.txt
generated_certificates/
```

这些已经写入 `.gitignore`。
