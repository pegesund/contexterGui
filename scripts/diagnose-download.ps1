# Spell download diagnostic
#
# Run on the machine where Spell's "Laster ned Bokmål..." screen fails with
# every entry showing "Feil". Tests the S3 endpoint Spell pulls models from
# (eu2.contabostorage.com) without needing Spell installed or running.
#
# Usage (open PowerShell, navigate to the folder with this script):
#   .\diagnose-download.ps1
#
# Writes a report to .\spell-diagnose.log next to the script. Send that
# file back to support.
#
# Tests, in order:
#   1. System clock vs UTC time servers
#       (AWS S3 rejects signed URLs if local clock is > 15min off)
#   2. DNS resolution of eu2.contabostorage.com
#       (some ISPs / AV products block Contabo)
#   3. TCP connection on port 443
#       (firewall reachability)
#   4. TLS handshake via Invoke-WebRequest (Windows cert store / SChannel)
#       (catches corporate AV that intercepts TLS with a custom CA)
#   5. Unsigned HEAD to check the endpoint responds
#   6. Presigned GET of one small file (sentence_split.pl, ~50 KB)
#       (this is what Spell does, end-to-end)
#   7. Same presigned GET via curl.exe if available
#       (rules out PowerShell vs curl HTTP-stack differences)
#
# Credentials are baked into the Spell binary too — they're public
# read-only keys for the 'spell' bucket. Including them here is the
# same exposure level.

$ErrorActionPreference = "Continue"
$logPath = Join-Path $PSScriptRoot "spell-diagnose.log"
Remove-Item $logPath -ErrorAction SilentlyContinue
"Spell download diagnostic — $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss zzz')" | Out-File $logPath
"=================================================================="  | Out-File $logPath -Append
"" | Out-File $logPath -Append

function Log($msg) {
    Write-Host $msg
    $msg | Out-File $logPath -Append
}

# ── Config (matches contexterGui/src/downloader.rs) ──────────────────
$Endpoint   = "https://eu2.contabostorage.com"
$Bucket     = "spell"
$Region     = "eu2"
$AccessKey  = "cd59e2c4bbbd7bd29951f126d87a096a"
$SecretKey  = "3f28f3941d0d20aaa829ef17c50fe4e7"
$TestKey    = "lang/nb/sentence_split.pl"      # small (~50KB), fast probe
$Host       = "eu2.contabostorage.com"

# ── 1. Clock check ───────────────────────────────────────────────────
Log "[1] System clock"
$localUtc = (Get-Date).ToUniversalTime()
Log "    Local UTC:  $($localUtc.ToString('yyyy-MM-ddTHH:mm:ssZ'))"
try {
    # Use the S3 endpoint's own Date response header as ground truth.
    # If S3 says we're > 15 min off, the presigned URL will be rejected
    # with HTTP 403 SignatureDoesNotMatch / RequestTimeTooSkewed.
    $head = Invoke-WebRequest -Uri "$Endpoint/" -Method Head -UseBasicParsing -TimeoutSec 10
    $serverDate = [DateTime]::Parse($head.Headers["Date"]).ToUniversalTime()
    Log "    Server UTC: $($serverDate.ToString('yyyy-MM-ddTHH:mm:ssZ'))"
    $skew = ($localUtc - $serverDate).TotalSeconds
    Log "    Skew: $([Math]::Round($skew)) seconds"
    if ([Math]::Abs($skew) -gt 600) {
        Log "    ⚠ CLOCK IS OFF BY > 10 MIN — this WILL cause AWS Sig V4 to reject downloads."
        Log "    Fix: Settings → Time & language → Date & time → Sync now"
    }
} catch {
    Log "    Could not fetch server time: $($_.Exception.Message)"
}
Log ""

# ── 2. DNS ───────────────────────────────────────────────────────────
Log "[2] DNS resolution"
try {
    $dns = Resolve-DnsName $Host -ErrorAction Stop
    foreach ($r in $dns) {
        if ($r.IPAddress) { Log "    $($r.Type) → $($r.IPAddress)" }
    }
} catch {
    Log "    ✗ DNS FAILED: $($_.Exception.Message)"
    Log "    Possible cause: ISP DNS blocks Contabo, or VPN/Pi-Hole filtering."
}
Log ""

# ── 3. TCP reachability ──────────────────────────────────────────────
Log "[3] TCP port 443 reachability"
try {
    $tcp = Test-NetConnection $Host -Port 443 -WarningAction SilentlyContinue
    Log "    TcpTestSucceeded: $($tcp.TcpTestSucceeded)"
    Log "    RemoteAddress:    $($tcp.RemoteAddress)"
    Log "    PingSucceeded:    $($tcp.PingSucceeded)"
    if (-not $tcp.TcpTestSucceeded) {
        Log "    ⚠ Cannot reach $Host on 443. Firewall / AV is blocking."
    }
} catch {
    Log "    Test-NetConnection failed: $($_.Exception.Message)"
}
Log ""

# ── 4. TLS via SChannel (Windows cert store) ─────────────────────────
Log "[4] TLS handshake (Windows cert store via SChannel)"
try {
    $resp = Invoke-WebRequest -Uri "$Endpoint/" -Method Head -UseBasicParsing -TimeoutSec 15
    Log "    HTTP Status:   $($resp.StatusCode)"
    Log "    Server:        $($resp.Headers['Server'])"
    Log "    ✓ TLS handshake OK (system cert store accepted endpoint)"
} catch {
    $we = $_.Exception
    Log "    ✗ FAILED: $($we.Message)"
    if ($we.InnerException) {
        Log "    Inner:    $($we.InnerException.Message)"
    }
    Log "    Possible cause: AV (Kaspersky/Norton/ESET/Bitdefender) is intercepting"
    Log "    TLS with its own cert and that cert isn't trusted by SChannel."
}
Log ""

# ── 5. AWS Sig V4 presigned URL ──────────────────────────────────────
# Same algorithm as contexterGui/src/downloader.rs::presign_url.
function Hmac-SHA256-Bytes($keyBytes, $dataString) {
    $hmac = New-Object System.Security.Cryptography.HMACSHA256
    $hmac.Key = $keyBytes
    $hmac.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($dataString))
}
function SHA256-Hex($s) {
    $sha = New-Object System.Security.Cryptography.SHA256Managed
    $h = $sha.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($s))
    -join ($h | ForEach-Object { "{0:x2}" -f $_ })
}
function AwsUriEncode($path) {
    # Matches aws_uri_encode_path in downloader.rs (preserves /).
    # Encodes everything except A-Z a-z 0-9 - _ . ~ /
    $sb = New-Object Text.StringBuilder
    foreach ($c in $path.ToCharArray()) {
        if (($c -ge 'A' -and $c -le 'Z') -or
            ($c -ge 'a' -and $c -le 'z') -or
            ($c -ge '0' -and $c -le '9') -or
            $c -in @('-','_','.','~','/')) {
            [void]$sb.Append($c)
        } else {
            [void]$sb.AppendFormat("%{0:X2}", [int]$c)
        }
    }
    $sb.ToString()
}

function Presign($key, $expiresSecs) {
    $now = (Get-Date).ToUniversalTime()
    $amzDate   = $now.ToString("yyyyMMddTHHmmssZ")
    $dateStamp = $now.ToString("yyyyMMdd")
    $canonicalUri = AwsUriEncode "/$Bucket/$key"
    $scope = "$dateStamp/$Region/s3/aws4_request"
    $credential = "$AccessKey/$scope"
    $credEncoded = $credential -replace '/', '%2F'
    $query = "X-Amz-Algorithm=AWS4-HMAC-SHA256&X-Amz-Credential=$credEncoded&X-Amz-Date=$amzDate&X-Amz-Expires=$expiresSecs&X-Amz-SignedHeaders=host"

    $canonicalRequest = "GET`n$canonicalUri`n$query`nhost:$Host`n`nhost`nUNSIGNED-PAYLOAD"
    $stringToSign = "AWS4-HMAC-SHA256`n$amzDate`n$scope`n$(SHA256-Hex $canonicalRequest)"

    $kDate    = Hmac-SHA256-Bytes ([System.Text.Encoding]::UTF8.GetBytes("AWS4$SecretKey")) $dateStamp
    $kRegion  = Hmac-SHA256-Bytes $kDate   $Region
    $kService = Hmac-SHA256-Bytes $kRegion "s3"
    $kSigning = Hmac-SHA256-Bytes $kService "aws4_request"
    $sigBytes = Hmac-SHA256-Bytes $kSigning $stringToSign
    $signature = -join ($sigBytes | ForEach-Object { "{0:x2}" -f $_ })

    "$Endpoint$canonicalUri" + "?" + $query + "&X-Amz-Signature=$signature"
}

Log "[5] Presigned GET via PowerShell (uses SChannel TLS — same as IE/Edge)"
$url = Presign $TestKey 600
Log "    Test key: $TestKey"
try {
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $resp = Invoke-WebRequest -Uri $url -UseBasicParsing -TimeoutSec 30
    $sw.Stop()
    Log "    HTTP Status:    $($resp.StatusCode)"
    Log "    Content-Length: $($resp.Headers['Content-Length']) bytes"
    Log "    Elapsed:        $($sw.ElapsedMilliseconds) ms"
    Log "    Body size:      $($resp.Content.Length) bytes"
    if ($resp.StatusCode -eq 200) {
        Log "    ✓ Download works via SChannel."
    }
} catch {
    Log "    ✗ FAILED: $($_.Exception.Message)"
    if ($_.Exception.Response) {
        $err = $_.Exception.Response
        Log "    HTTP Status: $([int]$err.StatusCode) $($err.StatusDescription)"
        try {
            $stream = $err.GetResponseStream()
            $reader = New-Object System.IO.StreamReader($stream)
            $body = $reader.ReadToEnd()
            Log "    Response body (first 500 chars):"
            Log "      $($body.Substring(0, [Math]::Min(500, $body.Length)))"
        } catch {}
    }
}
Log ""

# ── 6. Same GET via curl.exe (different HTTP stack, different cert path) ─
Log "[6] Presigned GET via curl.exe (if available)"
$curl = Get-Command curl.exe -ErrorAction SilentlyContinue
if ($curl) {
    $url2 = Presign $TestKey 600
    $tmp = [IO.Path]::GetTempFileName()
    try {
        $curlOut = & curl.exe -sS -w "%{http_code}|%{time_total}|%{size_download}" -o $tmp $url2 2>&1
        Log "    curl result: $curlOut"
        if (Test-Path $tmp) {
            $size = (Get-Item $tmp).Length
            Log "    Body written: $size bytes"
        }
    } finally {
        Remove-Item $tmp -ErrorAction SilentlyContinue
    }
} else {
    Log "    curl.exe not found in PATH — skipping."
}
Log ""

# ── 7. Environment summary ───────────────────────────────────────────
Log "[7] Environment"
Log "    Windows:        $((Get-CimInstance Win32_OperatingSystem).Caption) build $((Get-CimInstance Win32_OperatingSystem).BuildNumber)"
Log "    PowerShell:     $($PSVersionTable.PSVersion)"
Log "    Time zone:      $((Get-TimeZone).Id)"
Log "    NTP enabled:    $((w32tm /query /status 2>$null) | Select-String 'Source' | Out-String)".Trim()

# Active proxy?
$proxy = netsh winhttp show proxy 2>$null | Out-String
Log "    WinHTTP proxy:"
$proxy.Split("`n") | Where-Object { $_ -match '\S' } | ForEach-Object { Log "      $($_.Trim())" }

# AV products
Log "    AntiVirus:"
$av = Get-CimInstance -Namespace root/SecurityCenter2 -ClassName AntiVirusProduct -ErrorAction SilentlyContinue
if ($av) {
    foreach ($a in $av) { Log "      - $($a.displayName)" }
} else {
    Log "      (none detected via SecurityCenter2)"
}
Log ""

Log "Done. Report saved to:"
Log "    $logPath"
Log ""
Log "Please send this file back to support."
