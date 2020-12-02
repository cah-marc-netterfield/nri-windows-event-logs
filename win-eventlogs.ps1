###
# This PowerShell script grabs the Windows event logs and posts them to New Relic Logs.
# v2.2 -  Updated to use the get-winevent command to allow collecting extended logs.
#         Updated payload to more closely match the Infra agent based log forwarder.
#
# The following command is required for testing:
#
# [System.Diagnostics.EventLog]::CreateEventSource("New Relic","Application")
###

###
# Parameters (a.k.a. Command Line Arguments)
# Usage: -LogName "LogName"
# Required - throw if missing
###

param (
    [string]$LogName=$(throw "-LogName is mandatory"),
	[string[]]$ExclLevel = "",
	[string[]]$ExclEventID = ""
)

# App Version (so we know if someone's running an older release of this integration)
$appVersion = "2.2"

# New Relic Logs endpoint
$nrLogEndpoint = "https://log-api.newrelic.com/log/v1"

###
# Get license key from infra config file
###

((get-Content -Path ('c:\program files\New Relic\newrelic-infra\newrelic-infra.yml') ) | Select-String  -Pattern 'license_key:' ) -imatch "^license_key:\s*(?<key>\w*)$" | Out-Null
$LicenseKey = $matches['key']

###
# Logic to handle getting new log entries by saving current date to file
# to use as styarttime argument of Get-WinEvent in next pull. On first run we use current date.
# On subsequent runs it will use last date written to file.
#
# Uses LogName param to create timestamp for each LogName with slashes removed
###

$LAST_PULL_TIMESTAMP_FILE = "./last-pull-timestamp-$($LogName.replace('\',' ').replace('/',' ')).txt"

###
# If timestamp file exists, use it; otherwise,
# set timestamp to 15 minutes ago to pull some data on
# first run.
###

if(Test-Path $LAST_PULL_TIMESTAMP_FILE -PathType Leaf) {

    $timestamp = Get-Content -Path $LAST_PULL_TIMESTAMP_FILE -Encoding String | Out-String;
    $timestamp = [DateTime] $timestamp.ToString();

}else{

    $timestamp = (Get-Date).AddMinutes(-15);

}

###
# Write current timestamp to file to pull on next run.
###

Set-Content -Path $LAST_PULL_TIMESTAMP_FILE -Value (Get-Date -Format o)

###
# Pull events using StartTime with timestamp.  Use PowerShell Calculated Properties to rename some properties to expected values.
# For some reason, the -or operator doesn't work as expected in a single Where-Object cmdlet, so I had to use two.
###

$events = Get-WinEvent -filterhashtable @{
    Logname=$LogName
    StartTime=$timestamp
} | Where-Object {$_.LevelDisplayName -notin $ExclLevel} | Where-Object {$_.ID -notin $ExclEventID} | 
Select-Object ProcessID, UserID, 
@{Name = "hostname"; Expression ={ $env:COMPUTERNAME.ToLower() }},
@{Name = "event_type"; Expression ={'Windows Event Logs'}},
@{Name = "message"; Expression ={$_.message}},
@{Name = "ComputerName"; Expression ={$_.MachineName}},
@{Name = "Channel"; Expression ={$_.LogName}},
@{Name = "EventCategory"; Expression ={$_.level}},
@{Name = "WinEventType"; Expression ={$_.LevelDisplayName}},
@{Name = "EventID"; Expression ={$_.id}},
@{Name = "SourceName"; Expression ={$_.ProviderName}},
@{Name = "TimeGenerated"; Expression ={($_.TimeCreated).datetime }},
@{Name = "TimeWritten"; Expression ={( get-date ).DateTime }},
@{Name = "RecordNumber"; Expression ={ $_.RecordID }}

###
# ConvertTo-Json with -Compress argument required in order for Logs to consume.
###

$logPayload = @($events) | ConvertTo-Json -Compress
Write-Output $logPayload

# Set the headers expected by Logs
$headers = @{'Content-Type' = 'application/json'; 'X-License-Key' = $LicenseKey}

# Do the post to the Logs API
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$response = Invoke-WebRequest $nrLogEndpoint -Headers $headers -Method 'POST' -Body $logPayload

# Create a structure readable by NR Infra with metadata about this transaction for meta-Logging / debugging purposes
$nrMetaData = @{'event_type' = 'winEventLogAgt'; 'appVersion' = $appVersion; 'excludedEvents' = $ExclEventID; 'excludedLevels' = $ExclLevel; 'hostTime' = Get-Date -UFormat %s -Millisecond 0; 'pullAfter' = $timestamp.ToString(); 'logName' = $LogName; 'eventCount' = $events.Count; 'nrLogsResponse' = $response.StatusCode}

$metaPayload = @{
    name = "com.newrelic.windows.eventlog"
    integration_version = "0.2.1"
    protocol_version = 1
	  metrics = @($nrMetaData)
    inventory = @{}
    events = @()
} | ConvertTo-Json -Compress
Write-Output $metaPayload

#Write-Output "Response" $response.StatusCode
