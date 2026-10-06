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
$frameworkExe = Join-Path $projectRoot "dotnet-src\src\PaperSwitch\bin\Release\net8.0-windows10.0.19041.0\PaperSwitch.exe"
$buildScript = Join-Path $PSScriptRoot "build.ps1"

function Find-PaperSwitchExecutable {
    $srcDir = Join-Path $projectRoot "dotnet-src\src\PaperSwitch"
    # 排除編譯輸出，避免 obj 產生的程式碼造成每次啟動都重建。
    $latestSrc = Get-ChildItem -LiteralPath $srcDir -Recurse -File -ErrorAction Stop |
        Where-Object {
            $relativePath = $_.FullName.Substring($srcDir.Length)
            $relativePath -notmatch '[\\/](bin|obj)[\\/]' -and
            ($_.Extension -in '.cs','.xaml','.csproj','.resx','.ico','.png' -or $_.Name -eq 'version.txt')
        } |
        Sort-Object LastWriteTimeUtc -Descending | Select-Object -First 1

    foreach ($targetExe in @($publishedExe, $developmentExe, $frameworkExe)) {
        if (-not (Test-Path -LiteralPath $targetExe -PathType Leaf)) { continue }
        $buildTime = (Get-Item -LiteralPath $targetExe).LastWriteTimeUtc
        # 開發輸出的 apphost 可能保留舊時間，以實際組件判斷來源是否已編譯。
        if ($targetExe -ne $publishedExe) {
            $assembly = Join-Path (Split-Path -Parent $targetExe) 'PaperSwitch.dll'
            if (-not (Test-Path -LiteralPath $assembly -PathType Leaf)) { continue }
            $buildTime = (Get-Item -LiteralPath $assembly).LastWriteTimeUtc
        }
        if (-not $latestSrc -or $latestSrc.LastWriteTimeUtc -le $buildTime) { return $targetExe }
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

    $executable = Find-PaperSwitchExecutable
}

if ($null -eq $executable) {
    throw "建置完成後仍找不到 PaperSwitch.exe。"
}

Start-Process -FilePath $executable -WorkingDirectory (Split-Path -Parent $executable)
