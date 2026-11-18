<#

    ##################################################################
    # Creating an App in your Office 365 tenancy to access Graph API #
    ##################################################################

    1. Go to https://portal.azure.com/#blade/Microsoft_AAD_RegisteredApps/ApplicationsListBlade
    2. Click 'New registration'
    3. Give it a name and select 'Accounts in this organizational directory only (<TenantName> only - Single tenant)' and click 'Register'
    4. On the 'Overview' page click 'Add a Redirect URI'
    5. Now click 'Add a platform', then 'Web'
    6. Now enter 'http://localhost' in 'Redirect URL and 'https://localhost' in 'Logout URI' and click 'Configure'
    7. Now go to 'API permissions'
    8. Click 'Add a permission' and select 'Microsoft Graph'
    9. Click 'Delegated permissions'
    10. Search for 'group.readwrite.all' and tick the box, then click 'Add permission'
    11. Now click 'Grant admin consent for <TenantName>' and click 'Yes'
    12. Now go to 'Manifest'
    13. Replace "allowPublicClient": null with "allowPublicClient": true
    14. Click 'Save'.
    15. Go to 'Overview'
    16. Copy the 'Application (client) ID' into the $ClientId below.
    17. Go to 'Certificates & secrets'
    18. Create 'New client secret'
    19. Give the secret a name and select an expiry date.
    20. Copy the 'Value' into the $ClientSecret below, IMPORTANT once you click off the value you cannot view it again, make sure you copy it right away.
    21. Finally enter your Tenant Name in the $TenantName at top. E.g. $TenantName = "<tenantname>.onmicrosoft.com"
    
    ##############################################################################
    # Finding the Group Id of the Office 365 group you want to import the ICS to #
    ##############################################################################

    1. Go to https://portal.azure.com/#blade/Microsoft_AAD_IAM/GroupsManagementMenuBlade/AllGroups
    2. Find the Office 365 group you wish to import the calendar to
    3. In the 'Overview' of the group copy the 'Object Id'
    4. Paste the object ID into $GroupId below.

    ##########
    # Credit #
    ##########

    # Guide to authenticating the API with powershell 
    # https://www.thelazyadministrator.com/2019/07/22/connect-and-navigate-the-microsoft-graph-api-with-powershell/
#>

# Import-GroupCalendar.ps1
# Run in Windows PowerShell.

$ErrorActionPreference = "Stop"

# ---- Configure these values ----
$ics              = "C:\path\to\your\file.ics"
$TenantId         = "<your-tenant.onmicrosoft.com or your-domain.xyz>"
$ClientId         = "<your-client-id>"
$GroupId          = "<your-group-id>"
$RedirectUri      = "http://localhost"
$CalendarTimeZone = "W. Europe Standard Time" # Obtain the timezone for your usage by utilizing Get-TimeZone on PowerShell
# --------------------------------

foreach ($value in @($TenantId, $ClientId, $GroupId)) {
    if ([string]::IsNullOrWhiteSpace($value) -or $value -like "<*") {
        throw "Set TenantId, ClientId, and GroupId at the top of the script."
    }
}
if (-not (Test-Path -LiteralPath $ics)) {
    throw "ICS file not found: $ics"
}

Add-Type -AssemblyName System.Web
Add-Type -AssemblyName System.Windows.Forms

# Prompt for the client secret rather than storing it in this script.
$secureSecret = Read-Host "Enter the app's client secret" -AsSecureString
$secretPtr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($secureSecret)
try {
    $ClientSecret = [Runtime.InteropServices.Marshal]::PtrToStringBSTR($secretPtr)
}
finally {
    [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($secretPtr)
}

# Read UTF-8 (with or without BOM); use Windows-1252 if the file is not valid UTF-8.
$bytes = [System.IO.File]::ReadAllBytes($ics)
if ($bytes.Length -ge 3 -and
    $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) {
    $icsText = [System.Text.Encoding]::UTF8.GetString($bytes, 3, $bytes.Length - 3)
} else {
    $strictUtf8 = New-Object System.Text.UTF8Encoding($false, $true)
    try {
        $icsText = $strictUtf8.GetString($bytes)
    } catch {
        $icsText = [System.Text.Encoding]::GetEncoding(1252).GetString($bytes)
    }
}

# Unfold ICS continuation lines.
$lines = [System.Collections.Generic.List[string]]::new()
foreach ($line in ($icsText -split "`r?`n")) {
    if (($line.StartsWith(" ") -or $line.StartsWith("`t")) -and $lines.Count -gt 0) {
        $lines[$lines.Count - 1] += $line.Substring(1)
    } else {
        $lines.Add($line)
    }
}

function Convert-IcsDate {
    param(
        [Parameter(Mandatory)]
        [string]$Value
    )

    # A trailing Z means the timestamp is already UTC.
    if ($Value -match '^\d{8}T\d{6}Z$') {
        $parsed = [datetime]::ParseExact(
            $Value,
            "yyyyMMdd'T'HHmmss'Z'",
            [Globalization.CultureInfo]::InvariantCulture,
            [Globalization.DateTimeStyles]::AssumeUniversal
        )
        return [pscustomobject]@{
            DateTime = $parsed.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss")
            TimeZone = "UTC"
        }
    }

    # No Z means this is a local calendar time (or a date-only value).
    if ($Value -match '^\d{8}$') {
        $format = "yyyyMMdd"
    } elseif ($Value -match '^\d{8}T\d{6}$') {
        $format = "yyyyMMdd'T'HHmmss"
    } else {
        throw "Unsupported ICS date value: $Value"
    }

    $parsed = [datetime]::ParseExact(
        $Value,
        $format,
        [Globalization.CultureInfo]::InvariantCulture
    )

    return [pscustomobject]@{
        DateTime = $parsed.ToString("yyyy-MM-ddTHH:mm:ss")
        TimeZone = $CalendarTimeZone
    }
}

# Parse VEVENTs.
$events = [System.Collections.Generic.List[object]]::new()
$current = $null

foreach ($line in $lines) {
    if ($line -eq "BEGIN:VEVENT") {
        $current = @{}
        continue
    }

    if ($line -eq "END:VEVENT") {
        if ($current -and $current["SUMMARY"] -and $current["DTSTART"] -and $current["DTEND"]) {
            $startValue = Convert-IcsDate $current["DTSTART"]
            $endValue   = Convert-IcsDate $current["DTEND"]

            $events.Add([pscustomobject]@{
                Subject   = $current["SUMMARY"]
                Body      = $current["DESCRIPTION"]
                Start     = $startValue.DateTime
                StartZone = $startValue.TimeZone
                End       = $endValue.DateTime
                EndZone   = $endValue.TimeZone
                Location  = $current["LOCATION"]
            })
        }
        $current = $null
        continue
    }

    if ($null -eq $current) { continue }

    # Split at the first colon only; values can themselves contain colons.
    $colon = $line.IndexOf(":")
    if ($colon -lt 1) { continue }

    $property = $line.Substring(0, $colon)
    $value = $line.Substring($colon + 1)
    $name = $property.Split(";")[0].ToUpperInvariant()

    if ($name -in @("SUMMARY", "DESCRIPTION", "LOCATION", "DTSTART", "DTEND")) {
        $value = $value -replace '\\n', "`n" `
                        -replace '\\N', "`n" `
                        -replace '\\,', "," `
                        -replace '\\;', ";" `
                        -replace '\\\\', '\'
        $current[$name] = $value
    }
}

if ($events.Count -eq 0) {
    throw "No usable VEVENT entries found in the ICS file."
}

# Sign in and obtain an authorization code.
$tenantEncoded   = [System.Web.HttpUtility]::UrlEncode($TenantId)
$clientEncoded   = [System.Web.HttpUtility]::UrlEncode($ClientId)
$redirectEncoded = [System.Web.HttpUtility]::UrlEncode($RedirectUri)
$resourceEncoded = [System.Web.HttpUtility]::UrlEncode("https://graph.microsoft.com")

$authUrl = "https://login.microsoftonline.com/$tenantEncoded/oauth2/authorize" +
           "?response_type=code" +
           "&redirect_uri=$redirectEncoded" +
           "&client_id=$clientEncoded" +
           "&resource=$resourceEncoded" +
           "&prompt=select_account"

$form = New-Object System.Windows.Forms.Form
$form.Text = "Sign in to Microsoft 365"
$form.Width = 520
$form.Height = 720

$browser = New-Object System.Windows.Forms.WebBrowser
$browser.Dock = [System.Windows.Forms.DockStyle]::Fill
$browser.ScriptErrorsSuppressed = $true
$form.Controls.Add($browser)

$script:redirectedUri = $null
$browser.add_Navigated({
    param($sender, $e)
    if ($e.Url.AbsoluteUri.StartsWith($RedirectUri, [StringComparison]::OrdinalIgnoreCase)) {
        $script:redirectedUri = $e.Url.AbsoluteUri
        $form.Close()
    }
})

$form.add_Shown({
    $form.Activate()
    $browser.Navigate($authUrl)
})

[void]$form.ShowDialog()

if (-not $script:redirectedUri) {
    throw "Sign-in did not return to $RedirectUri. Check the app's registered redirect URI."
}

$returnUri = [uri]$script:redirectedUri
$query = [System.Web.HttpUtility]::ParseQueryString($returnUri.Query)

if ($query["error"]) {
    throw "Sign-in failed: $($query['error']) $($query['error_description'])"
}
$authCode = $query["code"]
if (-not $authCode) {
    throw "No authorization code was returned. Redirect URL: $script:redirectedUri"
}

# Exchange the code for a Microsoft Graph access token.
$tokenBody = @{
    grant_type    = "authorization_code"
    redirect_uri  = $RedirectUri
    client_id     = $ClientId
    client_secret = $ClientSecret
    code          = $authCode
    resource      = "https://graph.microsoft.com"
}

$tokenResponse = Invoke-RestMethod `
    -Uri "https://login.microsoftonline.com/$tenantEncoded/oauth2/token" `
    -Method Post `
    -ContentType "application/x-www-form-urlencoded" `
    -Body $tokenBody

if (-not $tokenResponse.access_token) {
    throw "Token response did not contain an access token."
}

# Create events in the Microsoft 365 Group calendar.
$apiUrl = "https://graph.microsoft.com/v1.0/groups/$GroupId/calendar/events"
$headers = @{ Authorization = "Bearer $($tokenResponse.access_token)" }

foreach ($event in $events) {
    $payload = @{
        subject = $event.Subject
        body = @{
            contentType = "text"
            content     = $event.Body
        }
        start = @{
            dateTime = $event.Start
            timeZone = $event.StartZone
        }
        end = @{
            dateTime = $event.End
            timeZone = $event.EndZone
        }
        location = @{
            displayName = $event.Location
        }
    }

Write-Host "Subject before upload: $($event.Subject)"
$json = $payload | ConvertTo-Json -Depth 5 -Compress
$utf8Body = [System.Text.Encoding]::UTF8.GetBytes($json)

Invoke-RestMethod `
    -Uri $apiUrl `
    -Headers $headers `
    -Method Post `
    -Body $utf8Body `
    -ContentType "application/json; charset=utf-8" | Out-Null

    Write-Host "Imported: $($event.Subject)"
}

Write-Host "Done. Imported $($events.Count) event(s)."
