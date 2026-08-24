[CmdletBinding()]
param(
    [switch]$NoLaunch,
    [switch]$Force
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$projectDir = $scriptDir
$projectName = Split-Path -Leaf $projectDir

# ----------------------------------------------------------------------
# 階段 ⓪：自動清理專案殘留的舊行程，防範 port 8080 被佔用衝突
# ----------------------------------------------------------------------
Get-Process python*, pythonw* -ErrorAction SilentlyContinue | Where-Object { $_.Path -like "*09_PaperSwitch*" } | Stop-Process -Force -ErrorAction SilentlyContinue

$embedDir = Join-Path $projectDir 'python_embed'
$embedPython = Join-Path $embedDir 'python.exe'

$entryPoint = 'app.py'
$reqFile = Join-Path $projectDir 'requirements.txt'

Write-Host '=================================================================' -ForegroundColor Cyan
Write-Host "🚀 【智慧自癒啟動系統】專案：$projectName" -ForegroundColor Yellow
Write-Host '=================================================================' -ForegroundColor Cyan

# ----------------------------------------------------------------------
# 階段 ①：檢查是否已具備現成的 Python 可攜環境 (場景 1：隨身碟 / 已就緒)
# ----------------------------------------------------------------------
$isEnvironmentReady = $false

if ($Force -and (Test-Path -LiteralPath $embedDir)) {
    Write-Host "[強制重建] 偵測到 -Force 參數，正在重置環境..." -ForegroundColor Magenta
    Remove-Item -LiteralPath $embedDir -Recurse -Force
}
if (Test-Path -LiteralPath $embedPython) {
    $oldEap = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    try {
        $testRun = & "$embedPython" -c "import sys, PIL, fitz, pypdf; print('READY')" 2>$null
        if ($LASTEXITCODE -eq 0 -and $testRun -match 'READY') {
            $isEnvironmentReady = $true
        }
    } catch {
        $isEnvironmentReady = $false
    } finally {
        $ErrorActionPreference = $oldEap
    }
}

if (-not $isEnvironmentReady) {
    Write-Host "🔍 偵測到本機尚未就緒 Python 可攜環境，正在啟動自動自癒佈置..." -ForegroundColor Yellow
    Write-Host ''

    $searchPaths = @(
        $projectDir,
        (Join-Path (Split-Path -Parent $projectDir) "00_home\downloads"),
        "D:\Caches",
        (Join-Path $env:USERPROFILE "Downloads")
    )

    $zipPath = $null
    foreach ($sp in $searchPaths) {
        if ($sp -and (Test-Path -LiteralPath $sp)) {
            $found = Get-ChildItem -LiteralPath $sp -Filter "*embed*amd64*.zip" -File -ErrorAction SilentlyContinue | Select-Object -First 1
            if ($found) {
                $zipPath = $found.FullName
                break
            }
        }
    }

    if (-not $zipPath) {
        $zipPath = Join-Path $projectDir 'python-3.13.0-embed-amd64.zip'
    }

    if (-not (Test-Path -LiteralPath $zipPath)) {
        $downloadUrl = "https://www.python.org/ftp/python/3.13.0/python-3.13.0-embed-amd64.zip"
        Write-Host "🌐 [1/4] 本機未發現 ZIP，正在從 Python 官方下載可攜核心..." -ForegroundColor Green
        Invoke-WebRequest -Uri $downloadUrl -OutFile $zipPath -UseBasicParsing
        Write-Host "   ✅ 下載完成：$zipPath" -ForegroundColor Gray
    } else {
        Write-Host "⚡ [1/4] 發現本機 Python ZIP 母檔：$zipPath（略過下載）" -ForegroundColor Green
    }

    Write-Host "📦 [2/4] 正在解壓縮可攜核心至 python_embed/ 資料夾..." -ForegroundColor Green
    if (Test-Path -LiteralPath $embedDir) {
        Remove-Item -LiteralPath $embedDir -Recurse -Force
    }
    Expand-Archive -LiteralPath $zipPath -DestinationPath $embedDir -Force

    Write-Host "⚙️  [3/4] 正在解除環境隔離限制並配置 pip 套件管理器..." -ForegroundColor Green
    $pthFile = Get-ChildItem -LiteralPath $embedDir -Filter "*._pth" -File | Select-Object -First 1
    if ($pthFile) {
        $zipName = [IO.Path]::GetFileNameWithoutExtension($pthFile.Name) + '.zip'
        $pthLines = @(
            $zipName,
            '.',
            'Lib\site-packages',
            'import site'
        )
        $asciiBytes = [System.Text.Encoding]::ASCII.GetBytes(($pthLines -join [Environment]::NewLine) + [Environment]::NewLine)
        [System.IO.File]::WriteAllBytes($pthFile.FullName, $asciiBytes)
    }

    $getPipPath = Join-Path $embedDir 'get-pip.py'
    $cachedGetPip = Join-Path (Split-Path -Parent $projectDir) "00_home\downloads\get-pip.py"
    if (Test-Path -LiteralPath $cachedGetPip) {
        Copy-Item -LiteralPath $cachedGetPip -Destination $getPipPath -Force
    } else {
        $getPipUrl = "https://bootstrap.pypa.io/get-pip.py"
        try {
            Invoke-WebRequest -Uri $getPipUrl -OutFile $getPipPath -UseBasicParsing
        } catch {
            Write-Host "   ⚠️  無法從網路下載 get-pip.py，將嘗試使用已內建套件機制" -ForegroundColor Yellow
        }
    }

    if (Test-Path -LiteralPath $getPipPath) {
        $oldEap = $ErrorActionPreference
        $ErrorActionPreference = 'Continue'
        try {
            $null = & "$embedPython" "$getPipPath" --no-warn-script-location 2>$null
        } finally {
            $ErrorActionPreference = $oldEap
        }
    }

    if ($reqFile -and (Test-Path -LiteralPath $reqFile)) {
        Write-Host "📚 [4/4] 正在自動安裝專案相依套件..." -ForegroundColor Green
        $oldEap = $ErrorActionPreference
        $ErrorActionPreference = 'Continue'
        try {
            & "$embedPython" -m pip install --no-warn-script-location -r "$reqFile"
        } finally {
            $ErrorActionPreference = $oldEap
        }
    }

    Write-Host ''
    Write-Host "🎉 【自癒成功】專案環境已 100% 佈置完成！" -ForegroundColor Cyan
    Write-Host '=================================================================' -ForegroundColor Cyan
}

if ($NoLaunch) {
    Write-Host "✨ 模式為僅建置環境，已順利完成。" -ForegroundColor Green
    return
}

$mainFile = Join-Path $projectDir $entryPoint
if (-not (Test-Path -LiteralPath $mainFile)) {
    throw "找不到專案啟動進入點：$mainFile"
}

Write-Host "🚀 正在啟動 $projectName ($entryPoint)..." -ForegroundColor Green
Write-Host ''

& "$embedPython" "$mainFile"
