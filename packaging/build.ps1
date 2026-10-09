$ErrorActionPreference = 'Stop'
$ProjectRoot = Split-Path -Parent $PSScriptRoot
$Python = Join-Path $ProjectRoot '.venv\Scripts\python.exe'
$Version = '0.1.36'
$AppNameBase = -join ([char[]](0x8FD0, 0x52A8, 0x4F1A, 0x79E9, 0x5E8F, 0x518C, 0x7CFB, 0x7EDF))
$AppName = "$AppNameBase-v$Version"
$OutputDirectory = Join-Path $ProjectRoot 'outputs'
$Executable = Join-Path $OutputDirectory "$AppName.exe"

& $Python -m pytest
if ($LASTEXITCODE -ne 0) {
    throw "pytest failed with exit code $LASTEXITCODE; aborting build."
}
& $Python -m PyInstaller `
    --noconfirm `
    --clean `
    --onefile `
    --windowed `
    --name $AppName `
    --paths (Join-Path $ProjectRoot 'src') `
    --distpath $OutputDirectory `
    --workpath (Join-Path $ProjectRoot 'build') `
    (Join-Path $ProjectRoot 'launcher.py')

$StartInfo = [System.Diagnostics.ProcessStartInfo]::new()
$StartInfo.FileName = $Executable
$StartInfo.Arguments = '--smoke-test'
$StartInfo.WorkingDirectory = $env:TEMP
$StartInfo.UseShellExecute = $false
$StartInfo.EnvironmentVariables['PATH'] = 'C:\Windows\System32;C:\Windows'
$StartInfo.EnvironmentVariables['QT_QPA_PLATFORM'] = 'offscreen'
$Process = [System.Diagnostics.Process]::Start($StartInfo)
if (-not $Process.WaitForExit(45000)) {
    $Process.Kill()
    throw 'Packaged application smoke test timed out.'
}
if ($Process.ExitCode -ne 0) {
    throw "Packaged application smoke test failed with exit code $($Process.ExitCode)."
}

$Hash = (Get-FileHash -Algorithm SHA256 -LiteralPath $Executable).Hash
Write-Host "Build and clean-environment smoke test succeeded: $Executable"
Write-Host "SHA256: $Hash"
