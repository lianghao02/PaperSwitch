param(
    [string]$PowerShellHost = 'pwsh.exe'
)

$ErrorActionPreference = 'Stop'
$sourceLauncher = Join-Path $PSScriptRoot 'run.ps1'
$fixtureRoot = Join-Path ([IO.Path]::GetTempPath()) ('PaperSwitch 入口測試 ' + [guid]::NewGuid().ToString('N'))
$fixtureScripts = Join-Path $fixtureRoot 'dotnet-src\scripts'
$sourceDir = Join-Path $fixtureRoot 'dotnet-src\src\PaperSwitch'
$ridDir = Join-Path $sourceDir 'bin\Release\net8.0-windows10.0.19041.0\win-x64'
$frameworkDir = Split-Path -Parent $ridDir
$publishDir = Join-Path $fixtureRoot 'dist\publish'
New-Item -ItemType Directory -Path $fixtureScripts -Force | Out-Null
Copy-Item -LiteralPath $sourceLauncher -Destination (Join-Path $fixtureScripts 'run.ps1')
$epoch = [datetime]::UtcNow.AddDays(-2)

function Set-FixtureFile {
    param([string]$RelativePath, [int]$Seconds)
    $path = Join-Path $fixtureRoot $RelativePath
    New-Item -ItemType Directory -Path (Split-Path -Parent $path) -Force | Out-Null
    [IO.File]::WriteAllText($path, '測試佔位檔，僅供 ValidateOnly 使用', [Text.UTF8Encoding]::new($false))
    (Get-Item -LiteralPath $path).LastWriteTimeUtc = $epoch.AddSeconds($Seconds)
}

function Assert-Launcher {
    param([string]$Name, [string]$Expected)
    $output = & $PowerShellHost -NoProfile -ExecutionPolicy Bypass -File (Join-Path $fixtureScripts 'run.ps1') -ValidateOnly 2>&1
    $code = $LASTEXITCODE
    if ($Expected) {
        if ($code -ne 0 -or ([string]($output | Select-Object -Last 1)).Trim() -ne $Expected) {
            throw "測試失敗：$Name"
        }
    } elseif ($code -eq 0) {
        throw "應要求建置卻通過：$Name"
    }
    Write-Output "通過：$Name"
}

# 假 EXE 不會被啟動；沒有 build.ps1，確保唯讀檢查不暗中建置。
Set-FixtureFile 'dotnet-src\src\PaperSwitch\MainWindow.xaml' 100
Assert-Launcher '完全沒有成品時明確失敗' ''
Set-FixtureFile 'dotnet-src\src\PaperSwitch\bin\Release\net8.0-windows10.0.19041.0\win-x64\PaperSwitch.exe' 10
Set-FixtureFile 'dotnet-src\src\PaperSwitch\bin\Release\net8.0-windows10.0.19041.0\win-x64\PaperSwitch.dll' 200
Assert-Launcher 'publish 缺失，以新 DLL 判定既有 Release 有效' (Join-Path $ridDir 'PaperSwitch.exe')
Set-FixtureFile 'dotnet-src\src\PaperSwitch\obj\Release\Generated.g.cs' 900
Set-FixtureFile 'dotnet-src\src\PaperSwitch\bin\Generated.cs' 900
Assert-Launcher 'obj 與 bin 更新不觸發重建' (Join-Path $ridDir 'PaperSwitch.exe')
Set-FixtureFile 'dist\publish\PaperSwitch.exe' 50
Assert-Launcher 'publish 過舊仍可回退有效 Release' (Join-Path $ridDir 'PaperSwitch.exe')
Set-FixtureFile 'dist\publish\PaperSwitch.exe' 300
Assert-Launcher '有效 publish 優先' (Join-Path $publishDir 'PaperSwitch.exe')
Set-FixtureFile 'dotnet-src\src\PaperSwitch\MainWindow.xaml' 400
Assert-Launcher '真正來源較新時要求建置' ''
Set-FixtureFile 'dotnet-src\src\PaperSwitch\bin\Release\net8.0-windows10.0.19041.0\PaperSwitch.exe' 10
Set-FixtureFile 'dotnet-src\src\PaperSwitch\bin\Release\net8.0-windows10.0.19041.0\PaperSwitch.dll' 500
Assert-Launcher '支援未指定 RID 的 Release 成品' (Join-Path $frameworkDir 'PaperSwitch.exe')
Set-FixtureFile 'dotnet-src\src\PaperSwitch\version.txt' 600
Assert-Launcher '版本資源更新也需要建置' ''
Write-Output "共 8 項通過；中文與空白路徑已涵蓋。測試資料保留於：$fixtureRoot"
