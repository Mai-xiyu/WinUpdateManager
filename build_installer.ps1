$ErrorActionPreference = 'Stop'

Write-Host '1. 正在发布项目 (Release, win-x64, 单文件)...' -ForegroundColor Cyan
dotnet publish 'WinUpdateManager\WinUpdateManager.csproj' -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -o 'publish'

Write-Host '2. 查找 Inno Setup 编译器...' -ForegroundColor Cyan
$isccPaths = @(
    "C:\Program Files (x86)\Inno Setup 6\ISCC.exe",
    "C:\Program Files\Inno Setup 6\ISCC.exe",
    "$env:LOCALAPPDATA\Programs\Inno Setup 6\ISCC.exe"
)

$isccPath = $isccPaths | Where-Object { Test-Path $_ } | Select-Object -First 1

if (-not $isccPath) {
    Write-Host "未找到 Inno Setup。正在尝试通过 winget 安装..." -ForegroundColor Yellow
    winget install -e --id JRSoftware.InnoSetup --accept-package-agreements --accept-source-agreements
    $isccPath = $isccPaths | Where-Object { Test-Path $_ } | Select-Object -First 1
    if (-not $isccPath) {
        Write-Error "Inno Setup 安装失败或未找到，请手动安装 Inno Setup 6。"
    }
}

Write-Host "3. 正在编译安装包..." -ForegroundColor Cyan
& $isccPath "installer.iss"

Write-Host '打包完成！安装包位于 Output 目录中。' -ForegroundColor Green
