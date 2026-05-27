$OutputEncoding = [Console]::OutputEncoding = [Text.UTF8Encoding]::new($false)
$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

$AppName = "论文格式修改助手"
$EntryScript = "wfp.py"
$SampleConfigDir = "sample_configs"

function Write-Step {
    param([string]$Message)
    Write-Host ""
    Write-Host "==> $Message" -ForegroundColor Cyan
}

function Assert-CommandSucceeded {
    param(
        [int]$ExitCode,
        [string]$FailureMessage
    )
    if ($ExitCode -ne 0) {
        throw $FailureMessage
    }
}

function Resolve-ProjectRoot {
    $currentDir = (Get-Location).Path
    if (Test-Path -LiteralPath (Join-Path $currentDir $EntryScript)) {
        return (Resolve-Path -LiteralPath $currentDir).Path
    }

    $scriptParent = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot "..")).Path
    if (Test-Path -LiteralPath (Join-Path $scriptParent $EntryScript)) {
        Write-Host "当前目录未找到 $EntryScript，已切换到脚本所在项目根目录：$scriptParent" -ForegroundColor Yellow
        return $scriptParent
    }

    throw "未找到 $EntryScript。请在项目根目录运行：powershell -ExecutionPolicy Bypass -File packaging/build_windows.ps1"
}

Write-Step "定位项目根目录"
$ProjectRoot = Resolve-ProjectRoot
Set-Location -LiteralPath $ProjectRoot
Write-Host "项目目录：$ProjectRoot"

Write-Step "检查入口文件"
if (-not (Test-Path -LiteralPath $EntryScript)) {
    throw "当前项目目录不包含 $EntryScript，停止打包。"
}
Write-Host "已找到入口文件：$EntryScript"

Write-Step "检查 Python"
if (-not (Get-Command python -ErrorAction SilentlyContinue)) {
    throw "未找到 python 命令。请先安装 Python 3，并确认 python 已加入 PATH。"
}
& python --version
Assert-CommandSucceeded $LASTEXITCODE "Python 检查失败。"

Write-Step "检查 pip"
& python -m pip --version
Assert-CommandSucceeded $LASTEXITCODE "pip 检查失败。请确认当前 Python 环境已安装 pip。"

if (Test-Path -LiteralPath "requirements.txt") {
    Write-Step "安装 requirements.txt 依赖"
    & python -m pip install -r "requirements.txt"
    Assert-CommandSucceeded $LASTEXITCODE "requirements.txt 依赖安装失败。"
} else {
    Write-Host "未找到 requirements.txt，跳过项目依赖安装。"
}

Write-Step "检查 PyInstaller"
& python -m PyInstaller --version
if ($LASTEXITCODE -ne 0) {
    Write-Host "未检测到 PyInstaller，正在尝试安装 pyinstaller..." -ForegroundColor Yellow
    & python -m pip install pyinstaller
    Assert-CommandSucceeded $LASTEXITCODE "PyInstaller 安装失败。请手动执行：python -m pip install pyinstaller"

    & python -m PyInstaller --version
    Assert-CommandSucceeded $LASTEXITCODE "PyInstaller 安装后仍不可用，请检查 Python 环境。"
}

Write-Step "清理旧打包产物"
foreach ($path in @("build", "dist")) {
    if (Test-Path -LiteralPath $path) {
        Remove-Item -LiteralPath $path -Recurse -Force
        Write-Host "已删除：$path"
    } else {
        Write-Host "未找到：$path，跳过。"
    }
}

Write-Step "生成 Windows onedir GUI 程序"
$pyInstallerArgs = @(
    "--noconfirm",
    "--clean",
    "--onedir",
    "--windowed",
    "--name",
    $AppName
)

if (Test-Path -LiteralPath $SampleConfigDir) {
    $pyInstallerArgs += @("--add-data", "$SampleConfigDir;$SampleConfigDir")
    Write-Host "已包含示例配置目录：$SampleConfigDir"
} else {
    Write-Host "未找到 $SampleConfigDir，跳过 --add-data。"
}

$pyInstallerArgs += $EntryScript

Write-Host "提示：如果使用 .doc / .wps 转换、Word COM 自动化或相关功能，请确保 requirements.txt 中的 pywin32 已安装。"
Write-Host "提示：PyInstaller 通常会自动收集 pythoncom、pywintypes、win32com；若运行 exe 时缺少相关模块，可在后续补充 hook 或 hidden-import。"

& python -m PyInstaller @pyInstallerArgs
Assert-CommandSucceeded $LASTEXITCODE "PyInstaller 打包失败。"

$DistPath = Join-Path $ProjectRoot "dist"
$ExePath = Join-Path $DistPath (Join-Path $AppName "$AppName.exe")

Write-Step "打包完成"
Write-Host "dist 目录：$DistPath" -ForegroundColor Green
Write-Host "exe 路径：$ExePath" -ForegroundColor Green
Write-Host ""
Write-Host "请手动双击 exe 测试 GUI 是否能正常启动：" -ForegroundColor Yellow
Write-Host $ExePath
Write-Host ""
Write-Host "注意：不要把 dist/、build/ 或自动生成的 *.spec 提交进 git。"

