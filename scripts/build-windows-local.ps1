# Local Windows release: replicates .github/workflows/release-windows.yml
# without downloading deps (assumes SWI-Prolog, ONNX runtime, whisper.dll,
# Velopack CLI are already on this machine).
#
# Usage:  pwsh -File scripts/build-windows-local.ps1 -Version 0.1.37 [-ReleasesRepo pegesund/spell_binaries] [-ReleaseChannel win]
#
# Produces dist/releases/win/Spell-Setup.exe + delta + RELEASES manifest.

param(
    [string]$Version = "0.1.37-local",
    [string]$OnnxDll = "",
    [string]$SwiplHome = "C:\Program Files\swipl",
    [string]$WhisperDir = "",
    [string]$ReleasesRepo = "pegesund/spell_binaries",
    [string]$ReleaseChannel = "win"
)

$ErrorActionPreference = "Stop"
$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
Set-Location $repoRoot
$env:SPELL_RELEASES_REPO = $ReleasesRepo
$env:SPELL_RELEASES_CHANNEL = $ReleaseChannel

function Resolve-OnnxDll {
    param([string]$RequestedPath)

    if ($RequestedPath -and (Test-Path $RequestedPath)) {
        return (Resolve-Path $RequestedPath).Path
    }

    $candidates = @(
        $env:ORT_DYLIB_PATH,
        (Join-Path $repoRoot "..\onnxruntime\onnxruntime-win-x64-1.24.4\lib\onnxruntime.dll"),
        (Join-Path $repoRoot "..\..\onnxruntime\onnxruntime-win-x64-1.24.4\lib\onnxruntime.dll"),
        "C:\onnxruntime\onnxruntime-win-x64-1.24.4\lib\onnxruntime.dll",
        "C:\onnxruntime\onnxruntime-win-x64-1.23.0\lib\onnxruntime.dll",
        "C:\Program Files\Microsoft Visual Studio\2022\Community\Common7\IDE\Extensions\Microsoft\CopilotUpgradeAgent\ServiceHub\Services\runtimes\win-x64\native\onnxruntime.dll",
        "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\Extensions\Microsoft\CopilotUpgradeAgent\ServiceHub\Services\runtimes\win-x64\native\onnxruntime.dll"
    ) | Where-Object { $_ -and $_.Trim() -ne "" }

    foreach ($candidate in $candidates) {
        if (Test-Path $candidate) {
            return (Resolve-Path $candidate).Path
        }
    }

    return $null
}

function Resolve-WhisperDll {
    param([string]$RequestedPath)

    $candidates = @()
    if ($RequestedPath -and $RequestedPath.Trim() -ne "") {
        if (Test-Path $RequestedPath -PathType Container) {
            $candidates += (Join-Path $RequestedPath "whisper.dll")
        } else {
            $candidates += $RequestedPath
        }
    }

    $candidates += @(
        (Join-Path $repoRoot "whisper-build\bin\Release\whisper.dll"),
        (Join-Path $repoRoot "..\whisper-build\bin\Release\whisper.dll"),
        (Join-Path $repoRoot "..\..\whisper-build\bin\Release\whisper.dll"),
        "C:\whisper-build\bin\Release\whisper.dll"
    )

    foreach ($candidate in $candidates) {
        if ($candidate -and (Test-Path $candidate)) {
            return (Resolve-Path $candidate).Path
        }
    }

    return $null
}

$resolvedOnnxDll = Resolve-OnnxDll -RequestedPath $OnnxDll
if (-not $resolvedOnnxDll) {
    Write-Error "Could not locate onnxruntime.dll automatically. Pass -OnnxDll <full-path>."
    exit 1
}

$resolvedWhisperDll = Resolve-WhisperDll -RequestedPath $WhisperDir
if (-not $resolvedWhisperDll) {
    Write-Error "Could not locate whisper.dll automatically. Pass -WhisperDir <dir-or-full-dll-path>."
    exit 1
}

Write-Host "=== Build spell.exe + native_bridge.exe ==="
$env:ORT_DYLIB_PATH = $resolvedOnnxDll
Write-Host "Using ONNX Runtime: $resolvedOnnxDll"
Write-Host "Using Whisper runtime: $resolvedWhisperDll"
Write-Host "Using release feed: $ReleasesRepo (channel: $ReleaseChannel)"
cargo build --release --bin spell --bin native_bridge
if ($LASTEXITCODE -ne 0) { Write-Error "cargo build failed"; exit 1 }

$bundleName = "Spell-windows-x64-$Version"
$bundleDir = "dist\$bundleName"
$frameworks = "$bundleDir\Frameworks"
$resSwipl = "$bundleDir\Resources\swipl"
$resAddin = "$bundleDir\Resources\word-addin"

if (Test-Path $bundleDir) { Remove-Item -Recurse -Force $bundleDir }
New-Item -Force -ItemType Directory -Path $frameworks | Out-Null
New-Item -Force -ItemType Directory -Path $resSwipl | Out-Null
New-Item -Force -ItemType Directory -Path $resAddin | Out-Null

Write-Host "=== Copy binaries ==="
Copy-Item "target\release\spell.exe" "$bundleDir\Spell.exe"
Copy-Item "target\release\native_bridge.exe" "$bundleDir\native_bridge.exe"

Write-Host "=== Copy SWI-Prolog DLLs ==="
Get-ChildItem "$SwiplHome\bin\*.dll" | ForEach-Object {
    Copy-Item $_.FullName "$frameworks\"
}

$requiredDeps = @("libswipl.dll", "libgmp-10.dll", "libwinpthread-1.dll",
                  "libgcc_s_seh-1.dll", "libstdc++-6.dll")
$fallbackDirs = @(
    "C:\Program Files\Git\mingw64\bin",
    "C:\msys64\mingw64\bin",
    "C:\msys64\ucrt64\bin",
    "C:\mingw64\bin"
)
foreach ($dep in $requiredDeps) {
    if (Test-Path "$frameworks\$dep") { continue }
    foreach ($dir in $fallbackDirs) {
        $c = Join-Path $dir $dep
        if (Test-Path $c) {
            Write-Host "  $dep <- $dir"
            Copy-Item $c "$frameworks\"
            break
        }
    }
    if (-not (Test-Path "$frameworks\$dep")) {
        Write-Error "Missing SWI dep: $dep"
        exit 1
    }
}

Write-Host "=== Copy SWI-Prolog home ==="
foreach ($item in @("ABI", "LICENSE", "README.md", "boot", "boot.prc",
                    "library", "app", "customize", "swipl.home")) {
    $src = "$SwiplHome\$item"
    if (Test-Path $src) { Copy-Item -Recurse $src "$resSwipl\" }
}
foreach ($item in @("boot.prc", "library", "boot")) {
    if (-not (Test-Path "$resSwipl\$item")) {
        Write-Error "Missing SWI home item: $item"
        exit 1
    }
}

Write-Host "=== Copy ONNX Runtime ==="
Copy-Item $resolvedOnnxDll "$frameworks\"

Write-Host "=== Copy MSVC runtime ==="
$vcDeps = @("vcruntime140.dll", "vcruntime140_1.dll", "msvcp140.dll",
            "concrt140.dll", "msvcp140_1.dll")
foreach ($dep in $vcDeps) {
    $src = "$env:SystemRoot\System32\$dep"
    if (Test-Path $src) {
        Copy-Item $src "$frameworks\"
    } else {
        Write-Host "  (optional) $dep not in System32"
    }
}

Write-Host "=== Copy Whisper DLLs ==="
$resolvedWhisperDir = Split-Path $resolvedWhisperDll -Parent
Get-ChildItem "$resolvedWhisperDir\*.dll" | ForEach-Object {
    Copy-Item $_.FullName "$frameworks\"
}

Write-Host "=== Copy fonts + Word add-in ==="
Copy-Item "fonts\OpenSans-Regular.ttf" "$bundleDir\Resources\"
foreach ($f in @("manifest.xml", "taskpane.html", "taskpane.js",
                 "commands.html", "commands.js", "fullchain.pem")) {
    if (Test-Path "word-addin\$f") {
        Copy-Item "word-addin\$f" "$resAddin\"
    }
}

$size = (Get-ChildItem $bundleDir -Recurse | Measure-Object -Property Length -Sum).Sum / 1MB
Write-Host ("Bundle size: {0:N1} MB" -f $size)

Write-Host "=== vpk pack ==="
$outputDir = "dist\releases\win"
New-Item -Force -ItemType Directory -Path $outputDir | Out-Null

& "$env:USERPROFILE\.dotnet\tools\vpk.exe" pack `
    --packId Spell `
    --packVersion $Version `
    --packDir $bundleDir `
    --mainExe "Spell.exe" `
    --channel $ReleaseChannel `
    --icon "assets\Spell.ico" `
    --packTitle "Spell" `
    --packAuthors "Cognio AS" `
    --outputDir $outputDir
if ($LASTEXITCODE -ne 0) { Write-Error "vpk pack failed"; exit 1 }

Write-Host ""
Write-Host "=== Output files ==="
Get-ChildItem $outputDir | ForEach-Object {
    $sz = "{0:N1} MB" -f ($_.Length / 1MB)
    Write-Host "  $($_.Name) ($sz)"
}
