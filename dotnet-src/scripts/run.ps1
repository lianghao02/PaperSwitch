param(
    [switch]$ValidateOnly
)

$ErrorActionPreference = "Stop"
[Console]::InputEncoding = [System.Text.UTF8Encoding]::new($false)
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)
$OutputEncoding = [System.Text.UTF8Encoding]::new($false)

$projectRoot = (Resolve-Path (Join-Path $PSScriptRoot "..\..")).Path
$publishedExe = Join-Path $projectRoot "dist\publish\PaperSwitch.exe"
$developmentExe = Join-Path $projectRoot "dotnet-src\src\PaperSwitch\bin\Release\net8.0-windows10.0.19041.0\win-x64\PaperSwitch.exe"
$buildScript = Join-Path $PSScriptRoot "build.ps1"

function Find-PaperSwitchExecutable {
    $targetExe = $null
    if (Test-Path -LiteralPath $publishedExe -PathType Leaf) {
        $targetExe = $publishedExe
    } elseif (Test-Path -LiteralPath $developmentExe -PathType Leaf) {
        $targetExe = $developmentExe
    }

    if ($targetExe) {
        # 檢查原始碼目錄是否有新於執行檔的變更
        $srcDir = Join-Path $projectRoot "dotnet-src\src\PaperSwitch"
        $latestSrc = Get-ChildItem -Path $srcDir -Recurse -File -Include "*.cs","*.xaml","*.csproj" -ErrorAction SilentlyContinue |
            Sort-Object LastWriteTime -Descending | Select-Object -First 1

        if ($latestSrc -and $latestSrc.LastWriteTime -gt (Get-Item $targetExe).LastWriteTime) {
            return $null # 標記為需要重新建置
        }

        return $targetExe
    }

    return $null
}

$executable = Find-PaperSwitchExecutable
if ($ValidateOnly) {
    if ($null -eq $executable) {
        throw "找不到已建置或最新的 PaperSwitch.exe。"
    }

    Write-Output $executable
    exit 0
}

if ($null -eq $executable) {
    & $buildScript
    if ($LASTEXITCODE -ne 0) {
        throw "PaperSwitch 建置失敗，結束碼：$LASTEXITCODE"
    }

    if (Test-Path -LiteralPath $publishedExe -PathType Leaf) {
        $executable = $publishedExe
    } elseif (Test-Path -LiteralPath $developmentExe -PathType Leaf) {
        $executable = $developmentExe
    }
}

if ($null -eq $executable) {
    throw "建置完成後仍找不到 PaperSwitch.exe。"
}

Start-Process -FilePath $executable -WorkingDirectory (Split-Path -Parent $executable)
