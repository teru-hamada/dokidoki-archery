param(
    [int]$Port = 5000,
    [int]$LowScoreThreshold = 15,
    [string]$OutputDir = "$(Join-Path $PSScriptRoot 'received')",
    [string]$GameUrl = 'https://teru-hamada.github.io/dokidoki-archery/'
)

$ErrorActionPreference = 'Stop'

function Write-Info {
    param([string]$Message)
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    Write-Host "[$timestamp] $Message"
}

function Ensure-Directory {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -ItemType Directory -Path $Path -Force | Out-Null
    }
}

function Get-RequestBodyText {
    param($Request)

    $encoding = [System.Text.Encoding]::UTF8
    if ($Request.ContentEncoding) {
        $encoding = $Request.ContentEncoding
    }

    $reader = $null
    try {
        $reader = New-Object System.IO.StreamReader($Request.InputStream, $encoding)
        return $reader.ReadToEnd()
    }
    finally {
        if ($reader) { $reader.Dispose() }
    }
}

function Send-JsonResponse {
    param(
        $Response,
        [int]$StatusCode,
        [hashtable]$Body
    )

    $json = ($Body | ConvertTo-Json -Depth 5)
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($json)

    $Response.StatusCode = $StatusCode
    $Response.ContentType = 'application/json; charset=utf-8'
    $Response.ContentEncoding = [System.Text.Encoding]::UTF8
    $Response.ContentLength64 = $bytes.Length
    $Response.OutputStream.Write($bytes, 0, $bytes.Length)
    $Response.OutputStream.Close()
}

function Set-CorsHeaders {
    param($Response)

    $Response.Headers['Access-Control-Allow-Origin'] = '*'
    $Response.Headers['Access-Control-Allow-Methods'] = 'GET, POST, OPTIONS'
    $Response.Headers['Access-Control-Allow-Headers'] = 'Content-Type'
}

function Open-GamePage {
    param([string]$Url)

    if ([string]::IsNullOrWhiteSpace($Url)) {
        return
    }

    try {
        Start-Process $Url | Out-Null
        Write-Info "Opened game URL in browser: $Url"
    }
    catch {
        Write-Info ("Failed to open game URL in browser: " + $_.Exception.Message)
    }
}

function Save-ScoreRecord {
    param(
        [int]$Score,
        [string]$RawBody,
        [string]$DirectoryPath
    )

    Ensure-Directory -Path $DirectoryPath

    $stamp = Get-Date -Format 'yyyyMMdd_HHmmss_fff'
    $filePath = Join-Path $DirectoryPath ("score_{0}.json" -f $stamp)

    $record = [ordered]@{
        receivedAt = (Get-Date).ToString('o')
        score      = $Score
        rawBody    = $RawBody
    }

    $record | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $filePath -Encoding UTF8
    return $filePath
}

function Create-LowScoreDesktopFiles {
    param([int]$Score)

    if ($Score -gt 10) {
        return @()
    }

    $desktopPath = [Environment]::GetFolderPath('Desktop')
    $createdFiles = @()
    $stamp = Get-Date -Format 'yyyyMMdd_HHmmss_fff'

    1..20 | ForEach-Object {
        $fileName = "notice_{0}_{1}.txt" -f $stamp, $_
        $filePath = Join-Path $desktopPath $fileName
        "Low score detected: $Score`r`nCreated at: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" | Set-Content -LiteralPath $filePath -Encoding UTF8
        $createdFiles += $filePath
    }

    return $createdFiles
}

Add-Type -AssemblyName System.Windows.Forms | Out-Null

Ensure-Directory -Path $OutputDir

$listener = [System.Net.HttpListener]::new()
$prefix = "http://localhost:$Port/"
$listener.Prefixes.Add($prefix)

try {
    $listener.Start()
}
catch {
    Write-Host ''
    Write-Host 'Failed to start the HTTP listener.' -ForegroundColor Red
    Write-Host 'Start PowerShell as Administrator and reserve the URL ACL if needed.' -ForegroundColor Yellow
    Write-Host "Example: netsh http add urlacl url=$prefix user=$env:USERNAME" -ForegroundColor Yellow
    throw
}

Write-Host ''
Write-Info "HTTP listener started: $prefix"
Write-Info "Score endpoint: ${prefix}score"
Write-Info "Warning threshold: $LowScoreThreshold or less"
Write-Info "Game URL: $GameUrl"
Write-Info 'Press Ctrl + C to stop the listener.'
Open-GamePage -Url $GameUrl
Write-Host ''

try {
    while ($listener.IsListening) {
        $context = $listener.GetContext()
        $request = $context.Request
        $response = $context.Response

        try {
            Set-CorsHeaders -Response $response

            if ($request.HttpMethod -eq 'OPTIONS') {
                $response.StatusCode = 204
                $response.Close()
                continue
            }

            if ($request.HttpMethod -eq 'POST' -and $request.Url.AbsolutePath -eq '/score') {
                $rawBody = Get-RequestBodyText -Request $request
                $scoreValue = $null
                $payload = $null

                if (-not [string]::IsNullOrWhiteSpace($rawBody)) {
                    try {
                        $payload = $rawBody | ConvertFrom-Json
                    }
                    catch {
                        $payload = $null
                    }
                }

                if ($payload -and $null -ne $payload.score) {
                    $scoreValue = [int]$payload.score
                }
                else {
                    $scoreValue = -1
                }

                $savedPath = Save-ScoreRecord -Score $scoreValue -RawBody $rawBody -DirectoryPath $OutputDir
                Write-Info "Received score: $scoreValue / Saved to: $savedPath"

                $desktopFiles = Create-LowScoreDesktopFiles -Score $scoreValue
                # if ($desktopFiles.Count -gt 0) {
                #     Write-Info "Created $($desktopFiles.Count) desktop files because the score was 10 or lower."
                # }

                # if ($scoreValue -le $LowScoreThreshold) {
                #     [System.Windows.Forms.MessageBox]::Show(
                #         "The score was lower than expected. Received score: $scoreValue",
                #         'Low Score Warning',
                #         [System.Windows.Forms.MessageBoxButtons]::OK,
                #         [System.Windows.Forms.MessageBoxIcon]::Warning
                #     ) | Out-Null
                # }

                Send-JsonResponse -Response $response -StatusCode 200 -Body @{
                    ok = $true
                    receivedScore = $scoreValue
                    savedPath = $savedPath
                    desktopFilesCreated = $desktopFiles.Count
                }
            }
            else {
                Send-JsonResponse -Response $response -StatusCode 404 -Body @{
                    ok = $false
                    message = 'Not Found'
                }
            }
        }
        catch {
            Write-Info ("Request handling error: " + $_.Exception.Message)
            try {
                Send-JsonResponse -Response $response -StatusCode 500 -Body @{
                    ok = $false
                    message = 'Internal Server Error'
                }
            }
            catch {
                # Ignore secondary response errors.
            }
        }
    }
}
finally {
    if ($listener) {
        $listener.Stop()
        $listener.Close()
    }
}
