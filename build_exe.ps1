$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot

$python = Join-Path $PSScriptRoot ".venv\Scripts\python.exe"
if (-not (Test-Path $python)) {
    throw "가상환경이 없습니다. 먼저 python -m venv .venv 후 패키지를 설치하세요."
}

& $python -m pip install -r requirements.txt pyinstaller
if ($LASTEXITCODE -ne 0) { throw "패키지 설치 실패" }

& $python -m PyInstaller --noconfirm --clean CMC_Scheduler.spec
if ($LASTEXITCODE -ne 0) { throw "exe 빌드 실패" }

$dist = Join-Path $PSScriptRoot "dist\CMC스케줄러"
Copy-Item -Force (Join-Path $PSScriptRoot "사용설명서.txt") $dist
if (Test-Path (Join-Path $PSScriptRoot "samples")) {
    $sampleDest = Join-Path $dist "samples"
    if (Test-Path $sampleDest) { Remove-Item -Recurse -Force $sampleDest }
    Copy-Item -Recurse (Join-Path $PSScriptRoot "samples") $sampleDest
    $outputDir = Join-Path $sampleDest "output"
    if (Test-Path $outputDir) { Remove-Item -Recurse -Force $outputDir }
}

Write-Host "빌드 완료: $dist\CMC스케줄러.exe"
