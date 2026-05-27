# Windows 打包说明

本文说明如何把本项目打包为 Windows 可分发版本，让普通用户解压后双击 `论文格式修改助手.exe` 运行，不需要自行安装 Python。

## 1. 安装打包依赖

建议在 Windows 上使用 Python 3，并先安装项目依赖：

```powershell
python -m pip install -U pip
python -m pip install -r requirements.txt
```

本项目使用 PyInstaller 打包。如果 `requirements.txt` 中已包含 `pyinstaller`，执行上面的命令即可；也可以单独安装：

```powershell
python -m pip install pyinstaller
```

Windows 下如需处理 `.doc`、`.wps`、修订或 Word COM 自动化相关能力，通常还需要安装 `pywin32`。当前 `requirements.txt` 已声明 Windows 平台依赖 `pywin32`。

## 2. 执行打包脚本

在项目根目录运行：

```powershell
powershell -ExecutionPolicy Bypass -File packaging/build_windows.ps1
```

脚本会执行以下操作：

- 检查项目入口 `wfp.py`
- 检查 `python` 和 `pip`
- 安装 `requirements.txt` 中的依赖
- 检查并在缺失时尝试安装 PyInstaller
- 清理旧的 `build/` 和 `dist/`
- 使用 PyInstaller 生成 Windows GUI 程序

如果存在 `sample_configs/`，脚本会把它一起加入打包产物；如果不存在，会自动跳过。

## 3. 打包产物位置

默认产物在：

```text
dist/论文格式修改助手/
```

可执行文件在：

```text
dist/论文格式修改助手/论文格式修改助手.exe
```

`dist/`、`build/`、自动生成的 `*.spec`、`*.exe` 和发布用 `*.zip` 不应提交到 git。

## 4. 用户如何运行

普通用户下载 Windows zip 后：

1. 解压 zip。
2. 打开解压后的文件夹。
3. 双击 `论文格式修改助手.exe`。
4. 按界面选择 Word 文档并执行排版。

建议发布前在一台没有开发环境的 Windows 电脑或干净虚拟机中手动双击测试。

## 5. 为什么第一版使用 onedir

第一版优先使用 `--onedir` 文件夹版，而不是 `--onefile` 单文件版，原因是：

- 依赖文件保持展开状态，问题更容易定位。
- 首次启动通常比 onefile 更稳定，避免每次运行都临时解压。
- Word、WPS、pywin32、GUI 相关依赖出现缺失时，更容易检查具体文件。
- 后续需要加入配置、模板、示例文件时，onedir 更直观。

等发布流程稳定后，再评估是否增加 onefile 作为可选产物。

## 6. 已知限制

- Windows exe 只适用于 Windows。
- 如果使用 `.doc` / `.wps` 转换，可能需要本机安装 Microsoft Word 或 WPS Office。
- 只处理 `.docx` 时通常不需要安装 Word。
- 当前 exe 未签名，杀毒软件或 Windows SmartScreen 可能提示风险。
- 第一次启动可能稍慢。

## 7. 压缩为 zip

打包完成并手动测试后，可以在项目根目录执行：

```powershell
Compress-Archive -Path "dist/论文格式修改助手" -DestinationPath "dist/论文格式修改助手-windows.zip" -Force
```

也可以在资源管理器中右键 `dist/论文格式修改助手/` 文件夹，选择“发送到 -> 压缩(zipped)文件夹”。

## 8. 上传到 GitHub Release

当前任务不上传 Release。后续发布时可按以下流程操作：

1. 在 GitHub 仓库页面打开 Releases。
2. 点击 Draft a new release。
3. 选择或创建版本标签，例如 `v0.1.0`。
4. 填写发布标题和变更说明。
5. 上传 `dist/论文格式修改助手-windows.zip`。
6. 确认测试说明和已知限制后发布。

发布前建议同步检查 [RELEASE_CHECKLIST.md](../RELEASE_CHECKLIST.md)。
