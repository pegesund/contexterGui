param(
    [string]$OnnxVersion = "1.24.4",
    [string]$SwiplVersion = "9.2.9",
    [switch]$SkipSwipl,
    [switch]$SkipOnnx
)

$ErrorActionPreference = "Stop"

$repoRoot = Resolve-Path (Join-Path $PSScriptRoot "..")
$workspaceRoot = Resolve-Path (Join-Path $repoRoot "..\..")

function Download-File {
    param(
        [string[]]$Urls,
        [string]$OutFile,
        [int64]$MinBytes = 1MB
    )

    foreach ($url in $Urls) {
        Write-Host "Trying: $url"
        try {
            Invoke-WebRequest -Uri $url -OutFile $OutFile -MaximumRedirection 5 -UseBasicParsing
            if ((Get-Item $OutFile).Length -gt $MinBytes) {
                Write-Host "Downloaded $((Get-Item $OutFile).Length) bytes"
                return
            }
        } catch {
            Write-Host "Failed: $_"
        }
    }

    throw "Could not download $OutFile"
}

if (-not $SkipOnnx) {
    $onnxRoot = Join-Path $workspaceRoot "onnxruntime"
    $onnxDir = Join-Path $onnxRoot "onnxruntime-win-x64-$OnnxVersion"
    $onnxDll = Join-Path $onnxDir "lib\onnxruntime.dll"

    if (Test-Path $onnxDll) {
        Write-Host "ONNX Runtime already present: $onnxDll"
    } else {
        New-Item -Force -ItemType Directory -Path $onnxRoot | Out-Null
        $zip = Join-Path $env:TEMP "onnxruntime-win-x64-$OnnxVersion.zip"
        $url = "https://github.com/microsoft/onnxruntime/releases/download/v$OnnxVersion/onnxruntime-win-x64-$OnnxVersion.zip"
        Download-File -Urls @($url) -OutFile $zip
        Expand-Archive -Path $zip -DestinationPath $onnxRoot -Force
        if (-not (Test-Path $onnxDll)) {
            throw "ONNX Runtime extraction failed: $onnxDll not found"
        }
        Write-Host "ONNX Runtime ready: $onnxDll"
    }

    Write-Host "Dev ORT path:"
    Write-Host "  $onnxDll"
}

if (-not $SkipSwipl) {
    $swiplDll = "C:\Program Files\swipl\bin\libswipl.dll"

    if (Test-Path $swiplDll) {
        Write-Host "SWI-Prolog already present: $swiplDll"
    } else {
        $installer = Join-Path $env:TEMP "swipl-$SwiplVersion-1.x64.exe"
        $file = "swipl-$SwiplVersion-1.x64.exe"
        $urls = @(
            "https://www.swi-prolog.org/download/stable/bin/$file",
            "https://github.com/SWI-Prolog/swipl-devel/releases/download/V$SwiplVersion/$file"
        )
        Download-File -Urls $urls -OutFile $installer

        Write-Host "Installing SWI-Prolog silently. This may prompt for administrator permission."
        Start-Process -FilePath $installer -ArgumentList "/S" -Wait -NoNewWindow

        if (-not (Test-Path $swiplDll)) {
            throw "SWI-Prolog install did not produce $swiplDll. Set SPELL_SWIPL_DLL to your libswipl.dll path."
        }
        Write-Host "SWI-Prolog ready: $swiplDll"
    }
}

Write-Host ""
Write-Host "Windows dev runtime setup complete."
Write-Host "Run from contexterGui:"
Write-Host "  cargo run --bin acatts-rust"
