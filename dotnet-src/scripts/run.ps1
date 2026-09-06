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
    if (Test-Path -LiteralPath $publishedExe -PathType Leaf) {
        return $publishedExe
    }

    if (Test-Path -LiteralPath $developmentExe -PathType Leaf) {
        return $developmentExe
    }

    return $null
}

$executable = Find-PaperSwitchExecutable
if ($ValidateOnly) {
    if ($null -eq $executable) {
        throw "找不到已建置的 PaperSwitch.exe。"
    }

    Write-Output $executable
    exit 0
}

if ($null -eq $executable) {
    & $buildScript
    if ($LASTEXITCODE -ne 0) {
        throw "PaperSwitch 建置失敗，結束碼：$LASTEXITCODE"
    }

    $executable = Find-PaperSwitchExecutable
}

if ($null -eq $executable) {
    throw "建置完成後仍找不到 PaperSwitch.exe。"
}

Start-Process -FilePath $executable -WorkingDirectory (Split-Path -Parent $executable)
