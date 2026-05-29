param(
    [string]$ProjectRoot = (Split-Path -Parent $PSScriptRoot),
    [string]$OutputDir = "dist/nuitka",
    [string]$StageRoot = "$env:TEMP\wps-image-extractor-nuitka"
)

$ErrorActionPreference = "Stop"

$resolvedRoot = (Resolve-Path $ProjectRoot).Path
$resolvedOutput = Join-Path $resolvedRoot $OutputDir
$resolvedStageRoot = $StageRoot
$stageSource = Join-Path $resolvedStageRoot "source"
$stageOutput = Join-Path $resolvedStageRoot "output"
$entryScript = Join-Path $stageSource "main.py"

if (-not (Test-Path (Join-Path $resolvedRoot "main.py"))) {
    throw "未找到入口文件：$(Join-Path $resolvedRoot 'main.py')"
}

Write-Host "Project root: $resolvedRoot"
Write-Host "Output dir : $resolvedOutput"
Write-Host "Stage root : $resolvedStageRoot"

if (Test-Path $resolvedStageRoot) {
    Remove-Item -Path $resolvedStageRoot -Recurse -Force
}

New-Item -ItemType Directory -Path $stageSource -Force | Out-Null
New-Item -ItemType Directory -Path $stageOutput -Force | Out-Null

Copy-Item -Path (Join-Path $resolvedRoot "main.py") -Destination $stageSource -Force
Copy-Item -Path (Join-Path $resolvedRoot "app") -Destination (Join-Path $stageSource "app") -Recurse -Force

if (Test-Path $resolvedOutput) {
    Remove-Item -Path $resolvedOutput -Recurse -Force
}

New-Item -ItemType Directory -Path $resolvedOutput -Force | Out-Null

uv run python -m nuitka `
    --standalone `
    --windows-console-mode=disable `
    --output-filename="WIE.exe" `
    --output-dir="$stageOutput" `
    --assume-yes-for-downloads `
    --enable-plugin=tk-inter `
    --include-package=app `
    --follow-imports `
    --remove-output `
    "$entryScript"

$builtDist = Join-Path $stageOutput "main.dist"
if (-not (Test-Path $builtDist)) {
    throw "未找到 Nuitka 产物目录：$builtDist"
}

Copy-Item -Path $builtDist -Destination (Join-Path $resolvedOutput "main.dist") -Recurse -Force

Write-Host "Build copied to: $(Join-Path $resolvedOutput 'main.dist')"
