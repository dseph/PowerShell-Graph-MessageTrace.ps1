#PowerShell-Graph-MessageTrace.ps1
# This script demonstrates how to retrieve message trace data via Microsoft Graph API.
 
<#
.SYNOPSIS
  Retrieve message trace data from Microsoft Graph API. This script:
    1) Acquires an OAuth token (client credentials flow).    
    2) Calls Microsoft Graph /messageTraces endpoint with appropriate filters.
.NOTES
  - Requires Application permissions: MessageTrace.Read.All
    - Be sure to do an administrtor grant after setting the specifc permissions.
    - The /messageTraces endpoint supports various filters, e.g. by date range, sender/recipient, etc. Adjust the $filter variable as needed.
    - The response token need to include if permission was granted- check token_info.txt:  
      "roles": [
        "ExchangeMessageTrace.Read.All"
      ]
  
.LINKS:
  - Graph API reference for messageTraces: https://learn.microsoft.com/en-us/graph/api/messageTrace-list?view=graph-rest-1.0&tabs=http
  - https://learn.microsoft.com/en-us/exchange/monitoring/trace-an-email-message/graph-api-message-trace
  - https://techcommunity.microsoft.com/blog/exchange/message-trace-support-using-graph-api-is-now-in-public-preview/4488587

#>
 
# =========================
# CONFIG
# =========================
$tenantId     = "<tenant-id>"     # TODO: Update with your tenant ID
$clientId     = "<app-id>"        # TODO: Update with your app registration's client ID
$clientSecret = "<client-secret>" # TODO: Update with your app registration's client secret

$start = "2026-04-17T00:00:00Z"   # TODO: Update with your desired received start date/time (ISO 8601 format)
$end   = "2026-04-17T23:59:59Z"   # TODO: Update with your desired received end date/time (ISO 8601 format)

# Output files
$startlogpath = "c:\test\"  # TODO: Update with your desired output path
$startlogpath = ".\"
$jsonFile      = $startlogpath + "messageTrace_full.json"
$csvFile       = "messageTrace.csv"
$logFile       = "trace_log.txt"
$tokenInfoFile = "token_info.txt"

# =========================
# HELPER: BASE64 DECODE JWT
# =========================
function Decode-JWT {
    param([string]$token)

    $parts = $token.Split('.')
    $payload = $parts[1].Replace('-', '+').Replace('_', '/')

    switch ($payload.Length % 4) {
        2 { $payload += '==' }
        3 { $payload += '=' }
    }

    $bytes = [Convert]::FromBase64String($payload)
    return [System.Text.Encoding]::UTF8.GetString($bytes) | ConvertFrom-Json
}

# =========================
# LOG FUNCTION
# =========================
function Log {
    param([string]$msg)
    $time = (Get-Date).ToString("s")
    "$time - $msg" | Out-File -Append $logFile
    Write-Host $msg
}

# =========================
# GET TOKEN (App OAuth)
# =========================
Log "Requesting OAuth token..."

$tokenBody = @{
    client_id     = $clientId
    scope         = "https://graph.microsoft.com/.default"
    client_secret = $clientSecret
    grant_type    = "client_credentials"
}

$tokenResponse = Invoke-RestMethod -Method POST `
    -Uri "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token" `
    -Body $tokenBody `
    -ContentType "application/x-www-form-urlencoded"

$accessToken = $tokenResponse.access_token

# =========================
# DECODE TOKEN + LOG PERMISSIONS
# =========================
Log "Decoding access token..."

$decoded = Decode-JWT $accessToken

# Extract roles (APPLICATION PERMISSIONS)
$roles = $decoded.roles

Log "Permissions (roles) in token:"
$roles | ForEach-Object { Log "  $_" }

# Save token details
$decoded | ConvertTo-Json -Depth 5 | Out-File $tokenInfoFile

# =========================
# BUILD QUERY
# =========================
$uri = "https://graph.microsoft.com/v1.0/admin/exchange/tracing/messageTraces?`$filter=receivedDateTime ge $start and receivedDateTime le $end"

Log "Calling Graph API..."
Log $uri

$headers = @{
    Authorization       = "Bearer $accessToken"
    "client-request-id" = [guid]::NewGuid()
    "User-Agent"        = "Graph-MessageTrace-Sample"
}

# =========================
# CALL GRAPH
# =========================
try {
    $response = Invoke-RestMethod -Method GET -Uri $uri -Headers $headers

    if (-not $response.value) {
        Log "No results returned"
        return
    }

    $results = $response.value

    # =========================
    # WRITE FILES
    # =========================
    Log "Writing JSON output..."
    $results | ConvertTo-Json -Depth 10 | Out-File $jsonFile -Encoding utf8

    Log "Writing CSV output..."
    $results |
        Select-Object id, senderAddress, recipientAddress, subject, receivedDateTime |
        Export-Csv $csvFile -NoTypeInformation -Encoding UTF8

    Log "Completed successfully"
}
catch {
    Log "ERROR: $($_.Exception.Message)"

    if ($_.Exception.Response) {
        $reader = New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream())
        $body = $reader.ReadToEnd()

        Log "Response Body:"
        Log $body
    }
}
 
     
