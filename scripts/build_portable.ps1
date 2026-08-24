[CmdletBinding()]
param(
    [string]$OutputDirectory
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$projectDir = Split-Path -Parent $PSScriptRoot
$versionPath = Join-Path $projectDir 'version.json'
$embedDir = Join-Path $projectDir 'python_embed'
$embedPython = Join-Path $embedDir 'python.exe'

if (-not (Test-Path -LiteralPath $versionPath)) { throw "Missing version file: $versionPath" }
if (-not (Test-Path -LiteralPath $embedPython)) { throw "Missing portable Python: $embedPython" }

$versionInfo = Get-Content -LiteralPath $versionPath -Raw -Encoding UTF8 | ConvertFrom-Json
if ([string]::IsNullOrWhiteSpace($versionInfo.version)) { throw 'version.json has no version.' }

$packageName = "PaperSwitch_v$($versionInfo.version)_Portable"
if ([string]::IsNullOrWhiteSpace($OutputDirectory)) {
    $OutputDirectory = Join-Path $projectDir 'dist'
}
$OutputDirectory = [System.IO.Path]::GetFullPath($OutputDirectory)
$stagingDir = Join-Path $OutputDirectory "$packageName.staging"
$archivePath = Join-Path $OutputDirectory "$packageName.zip"

$requiredFiles = @(
    'app.py',
    'requirements.txt',
    'version.json',
    'RUN.bat',
    'setup_and_run.ps1',
    'README.md',
    'CHANGELOG.md'
)
foreach ($relativePath in $requiredFiles) {
    if (-not (Test-Path -LiteralPath (Join-Path $projectDir $relativePath))) {
        throw "Required package file is missing: $relativePath"
    }
}

$launcherFiles = @(Get-ChildItem -LiteralPath $projectDir -Filter '*.vbs' -File)
if ($launcherFiles.Count -ne 1) { throw 'Expected exactly one VBS launcher in the project root.' }

$moduleCheck = @'
import importlib
required = ("PIL", "pypdf", "pymupdf", "win32com.client", "dotenv")
missing = []
for module in required:
    try:
        importlib.import_module(module)
    except Exception:
        missing.append(module)
if missing:
    raise SystemExit("Missing portable modules: " + ", ".join(missing))
print("Portable runtime ready")
'@
$moduleCheck | & $embedPython -
if ($LASTEXITCODE -ne 0) { throw 'Portable Python dependency check failed. Packaging stopped.' }

New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
if (Test-Path -LiteralPath $stagingDir) { Remove-Item -LiteralPath $stagingDir -Recurse -Force }
if (Test-Path -LiteralPath $archivePath) { Remove-Item -LiteralPath $archivePath -Force }

try {
    New-Item -ItemType Directory -Path $stagingDir -Force | Out-Null
    foreach ($relativePath in $requiredFiles) {
        Copy-Item -LiteralPath (Join-Path $projectDir $relativePath) -Destination (Join-Path $stagingDir $relativePath) -Force
    }
    Copy-Item -LiteralPath $launcherFiles[0].FullName -Destination (Join-Path $stagingDir $launcherFiles[0].Name) -Force
    Copy-Item -LiteralPath $embedDir -Destination (Join-Path $stagingDir 'python_embed') -Recurse -Force

    foreach ($dataFolder in @('uploads', 'converted')) {
        $targetFolder = Join-Path $stagingDir $dataFolder
        New-Item -ItemType Directory -Path $targetFolder -Force | Out-Null
        New-Item -ItemType File -Path (Join-Path $targetFolder '.gitkeep') -Force | Out-Null
    }

    Compress-Archive -Path (Join-Path $stagingDir '*') -DestinationPath $archivePath -CompressionLevel Optimal
    $verifyDir = Join-Path ([System.IO.Path]::GetTempPath()) ("PaperSwitch_portable_verify_" + [Guid]::NewGuid().ToString('N'))
    try {
        Expand-Archive -LiteralPath $archivePath -DestinationPath $verifyDir -Force
        $verifyPython = Join-Path $verifyDir 'python_embed\python.exe'
        & $verifyPython -m py_compile (Join-Path $verifyDir 'app.py')
        if ($LASTEXITCODE -ne 0) { throw 'Portable archive app.py compile failed.' }
        $runtimeCommand = 'import PIL, pypdf, pymupdf, win32com.client, dotenv; print(''Portable archive runtime ready'')'
        & $verifyPython -c $runtimeCommand
        if ($LASTEXITCODE -ne 0) { throw 'Portable archive dependency import failed.' }
        $archiveRequired = @('app.py', 'version.json', 'RUN.bat', 'python_embed\python.exe', 'uploads\.gitkeep', 'converted\.gitkeep', $launcherFiles[0].Name)
        foreach ($relativePath in $archiveRequired) {
            if (-not (Test-Path -LiteralPath (Join-Path $verifyDir $relativePath))) {
                throw "Portable archive item is missing: $relativePath"
            }
        }
        Write-Output 'Portable archive verification passed.'
    } finally {
        if (Test-Path -LiteralPath $verifyDir) { Remove-Item -LiteralPath $verifyDir -Recurse -Force }
    }
    $sha256 = [System.Security.Cryptography.SHA256]::Create()
    $stream = [System.IO.File]::OpenRead($archivePath)
    try {
        $hash = ([System.BitConverter]::ToString($sha256.ComputeHash($stream))).Replace('-', '').ToLowerInvariant()
    } finally {
        $stream.Dispose()
        $sha256.Dispose()
    }
    $sizeMiB = [math]::Round((Get-Item -LiteralPath $archivePath).Length / 1MB, 1)
    Write-Output "Package created: $archivePath"
    Write-Output "Size: $sizeMiB MiB"
    Write-Output "SHA-256: $hash"
} finally {
    if (Test-Path -LiteralPath $stagingDir) { Remove-Item -LiteralPath $stagingDir -Recurse -Force }
}
