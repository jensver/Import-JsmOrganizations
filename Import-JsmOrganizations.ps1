<#
.SYNOPSIS
Imports Jira Service Management organizations from a CSV file and optionally attaches them to a Service Desk.

.DESCRIPTION
- Reads a CSV with a single column (default header: Organization)
- Creates organizations in JSM (idempotent: existing orgs are returned)
- Optionally attaches each org to a specific Service Desk (Service Desk ID, not projectId)
- Includes retry/backoff for transient throttling (429) and 5xx responses
- Produces an end-of-run error summary and can export errors to CSV
- Returns exit code 1 if any errors occurred (CI-friendly)

.REQUIREMENTS
PowerShell 5.1+ (Windows PowerShell) or PowerShell 7+

.AUTH
Atlassian Cloud: Basic Auth (email + API token)
#>

[CmdletBinding()]
param(
  [Parameter(Mandatory = $true)]
  [ValidateNotNullOrEmpty()]
  [string]$BaseUrl,

  [Parameter(Mandatory = $true)]
  [ValidateNotNullOrEmpty()]
  [string]$Email,

  [Parameter(Mandatory = $true)]
  [ValidateNotNullOrEmpty()]
  [string]$ApiToken,

  [Parameter(Mandatory = $true)]
  [ValidateNotNullOrEmpty()]
  [string]$CsvPath,

  [Parameter(Mandatory = $false)]
  [ValidateNotNullOrEmpty()]
  [string]$ColumnName = "Organization",

  [Parameter(Mandatory = $false)]
  [int]$ServiceDeskId,

  [Parameter(Mandatory = $false)]
  [int]$RowDelayMs = 300,

  [Parameter(Mandatory = $false)]
  [int]$MaxRetries = 6,

  [Parameter(Mandatory = $false)]
  [int]$MaxBackoffSec = 15,

  [Parameter(Mandatory = $false)]
  [switch]$ExportErrors,

  [Parameter(Mandatory = $false)]
  [string]$ErrorExportCsv = ".\jsm_org_import_errors.csv",

  [Parameter(Mandatory = $false)]
  [switch]$AttachOnly
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Get-HttpStatusCode($err) {
  try {
    if ($err.Exception.Response -and $err.Exception.Response.StatusCode) {
      return [int]$err.Exception.Response.StatusCode
    }
    return $null
  } catch { return $null }
}

function Get-HttpErrorBody($err) {
  try {
    $resp = $err.Exception.Response
    if (-not $resp) { return $null }
    $stream = $resp.GetResponseStream()
    if (-not $stream) { return $null }
    $reader = New-Object System.IO.StreamReader($stream)
    $reader.ReadToEnd()
  } catch { $null }
}

function Get-RetryAfterSeconds($err) {
  try {
    $resp = $err.Exception.Response
    if (-not $resp) { return $null }
    $retryAfter = $resp.Headers["Retry-After"]
    if ([string]::IsNullOrWhiteSpace($retryAfter)) { return $null }
    $sec = 0
    if ([int]::TryParse($retryAfter, [ref]$sec)) { return $sec }
    return $null
  } catch { $null }
}

function New-AuthHeaders {
  param([string]$Email, [string]$ApiToken)

  $basic = [Convert]::ToBase64String(
    [Text.Encoding]::ASCII.GetBytes("$($Email):$($ApiToken)")
  )

  return @{
    Authorization  = "Basic $basic"
    Accept         = "application/json"
    "Content-Type" = "application/json; charset=utf-8"
  }
}

function Invoke-WithRetry {
  param(
    [Parameter(Mandatory=$true)]
    [scriptblock]$Action,
    [int]$MaxRetries = 5,
    [int]$MaxBackoffSec = 10,
    [string]$OpName = "API call"
  )

  for ($i = 1; $i -le $MaxRetries; $i++) {
    try {
      return & $Action
    } catch {
      $status = Get-HttpStatusCode $_
      $retryAfter = Get-RetryAfterSeconds $_
      $backoff = if ($retryAfter) { $retryAfter } else { [Math]::Min([Math]::Pow(2, $i), $MaxBackoffSec) }

      $isTransient =
        ($status -eq 429) -or
        ($status -ge 500 -and $status -le 599)

      if ($isTransient -and $i -lt $MaxRetries) {
        Write-Warning "$OpName transient failure (HTTP $status). Retry $i/$MaxRetries in ${backoff}s..."
        Start-Sleep -Seconds $backoff
        continue
      }

      throw
    }
  }
}

function Normalize-BaseUrl([string]$u) {
  if ($u.EndsWith("/")) { return $u.TrimEnd("/") }
  return $u
}

$BaseUrl = Normalize-BaseUrl $BaseUrl
$Headers = New-AuthHeaders -Email $Email -ApiToken $ApiToken

if (-not (Test-Path -LiteralPath $CsvPath)) {
  throw "CSV file not found: $CsvPath"
}

$errors = @()

# Optional pre-flight check for service desk access (if ServiceDeskId provided)
if ($PSBoundParameters.ContainsKey("ServiceDeskId") -and $ServiceDeskId -gt 0) {
  try {
    $sd = Invoke-RestMethod -Method Get -Uri "$BaseUrl/rest/servicedeskapi/servicedesk/$ServiceDeskId" -Headers $Headers
    Write-Host "Service Desk $ServiceDeskId accessible: $($sd.projectName) [$($sd.projectKey)]"
  } catch {
    $errors += [pscustomobject]@{
      Timestamp    = (Get-Date).ToString("s")
      Organization = ""
      Step         = "PreFlight"
      HttpStatus   = (Get-HttpStatusCode $_)
      Status       = $_.Exception.Message
      Details      = (Get-HttpErrorBody $_)
    }
    throw "Cannot access Service Desk $ServiceDeskId (permissions or wrong ID)."
  }
} elseif (-not $AttachOnly) {
  # no-op
} elseif ($AttachOnly -and (-not ($PSBoundParameters.ContainsKey("ServiceDeskId") -and $ServiceDeskId -gt 0))) {
  throw "AttachOnly requires -ServiceDeskId."
}

# Import CSV
$rows = Import-Csv -Path $CsvPath -Encoding utf8
if (-not $rows -or $rows.Count -eq 0) { throw "CSV appears empty or unreadable: $CsvPath" }

# If ColumnName doesn't exist, auto-detect first column
$firstRowProps = $rows[0].PSObject.Properties.Name
if (-not ($firstRowProps -contains $ColumnName)) {
  $auto = $firstRowProps | Select-Object -First 1
  Write-Warning "Column '$ColumnName' not found. Auto-detected '$auto' instead."
  $ColumnName = $auto
}

Write-Host "Using column header: '$ColumnName'"

foreach ($row in $rows) {
  $orgName = $row.$ColumnName
  if ($null -ne $orgName) { $orgName = $orgName.Trim() }

  if ([string]::IsNullOrWhiteSpace($orgName)) { continue }

  Write-Host "Processing org: $orgName"

  $orgId = $null

  if (-not $AttachOnly) {
    # 1) Create org
    $createJson  = @{ name = $orgName } | ConvertTo-Json -Compress
    $createBytes = [System.Text.Encoding]::UTF8.GetBytes($createJson)

    try {
      $org = Invoke-WithRetry -MaxRetries $MaxRetries -MaxBackoffSec $MaxBackoffSec -OpName "Create org" -Action {
        Invoke-RestMethod -Method Post -Uri "$BaseUrl/rest/servicedeskapi/organization" -Headers $Headers -Body $createBytes
      }
      $orgId = [int]$org.id
      Write-Host "  ✔ Org ID: $orgId"
    } catch {
      $errors += [pscustomobject]@{
        Timestamp    = (Get-Date).ToString("s")
        Organization = $orgName
        Step         = "Create"
        HttpStatus   = (Get-HttpStatusCode $_)
        Status       = $_.Exception.Message
        Details      = (Get-HttpErrorBody $_)
      }
      Write-Warning "  ✖ Create failed"
      if ($RowDelayMs -gt 0) { Start-Sleep -Milliseconds $RowDelayMs }
      continue
    }
  } else {
    # AttachOnly mode needs orgId. We try to find it via search/list (paged).
    # For large instances, this can be slower; recommended to run full create+attach when possible.
    try {
      $start = 0
      $limit = 50
      $found = $null

      while ($true) {
        $page = Invoke-RestMethod -Method Get -Uri "$BaseUrl/rest/servicedeskapi/organization?start=$start&limit=$limit" -Headers $Headers
        foreach ($v in $page.values) {
          if ($v.name -eq $orgName) { $found = $v; break }
        }
        if ($found) { break }
        if ($page.isLastPage -eq $true) { break }
        $start += $limit
      }

      if (-not $found) {
        throw "Organization not found by exact name. Run without -AttachOnly to create it first."
      }

      $orgId = [int]$found.id
      Write-Host "  ✔ Found Org ID: $orgId"
    } catch {
      $errors += [pscustomobject]@{
        Timestamp    = (Get-Date).ToString("s")
        Organization = $orgName
        Step         = "Lookup"
        HttpStatus   = (Get-HttpStatusCode $_)
        Status       = $_.Exception.Message
        Details      = (Get-HttpErrorBody $_)
      }
      Write-Warning "  ✖ Lookup failed"
      if ($RowDelayMs -gt 0) { Start-Sleep -Milliseconds $RowDelayMs }
      continue
    }
  }

  # 2) Attach to Service Desk (optional)
  if ($PSBoundParameters.ContainsKey("ServiceDeskId") -and $ServiceDeskId -gt 0) {
    $attachJson  = @{ organizationId = $orgId } | ConvertTo-Json -Compress
    $attachBytes = [System.Text.Encoding]::UTF8.GetBytes($attachJson)

    try {
      Invoke-WithRetry -MaxRetries $MaxRetries -MaxBackoffSec $MaxBackoffSec -OpName "Attach org to SD $ServiceDeskId" -Action {
        Invoke-RestMethod -Method Post -Uri "$BaseUrl/rest/servicedeskapi/servicedesk/$ServiceDeskId/organization" -Headers $Headers -Body $attachBytes
      }
      Write-Host "  ✔ Attached to SD $ServiceDeskId"
    } catch {
      $errors += [pscustomobject]@{
        Timestamp    = (Get-Date).ToString("s")
        Organization = $orgName
        Step         = "Attach"
        HttpStatus   = (Get-HttpStatusCode $_)
        Status       = $_.Exception.Message
        Details      = (Get-HttpErrorBody $_)
      }
      Write-Warning "  ✖ Attach failed"
    }
  }

  if ($RowDelayMs -gt 0) { Start-Sleep -Milliseconds $RowDelayMs }
}

# =========================
# FINAL ERROR SUMMARY
# =========================
Write-Host ""
Write-Host "========== ERROR SUMMARY ==========" -ForegroundColor Yellow

if ($errors.Count -eq 0) {
  Write-Host "✔ No errors encountered 🎉" -ForegroundColor Green
  exit 0
}

$errors | Sort-Object Step, Organization | Format-Table -AutoSize

if ($ExportErrors) {
  try {
    $errors | Export-Csv -Path $ErrorExportCsv -NoTypeInformation -Encoding utf8
    Write-Host ""
    Write-Host "Errors exported to: $ErrorExportCsv" -ForegroundColor Yellow
  } catch {
    Write-Warning "Failed to export errors to CSV: $ErrorExportCsv"
  }
}

exit 1
