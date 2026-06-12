<#
.SYNOPSIS
    Prints Veeam Backup & Replication session log records from the last 24 hours.

.DESCRIPTION
    Uses the Veeam Backup PowerShell module/snap-in to enumerate recent sessions for:
      - all backup jobs returned by Get-VBRJob
      - VBR sessions returned by Get-VBRSession, when available
      - backup sessions returned by Get-VBRBackupSession
      - internal/core sessions returned through Veeam.Backup.Core.CBackupSession, when available

    This intentionally writes to standard output only. Redirect stdout if you want a file.

.PARAMETER Hours
    Time window (in hours) to collect sessions/log records from. Default is 24.

.PARAMETER Json
    Emit one JSON object per line (JSON Lines) for machine-readable output.

.EXAMPLE
    .\Veeam_Collector.ps1

.EXAMPLE
    .\Veeam_Collector.ps1 -Hours 48

.EXAMPLE
    .\Veeam_Collector.ps1 -Json

.NOTES
    Usage notes:
      - Run this script in PowerShell on a Veeam Backup & Replication server or a host
        with the Veeam console/PowerShell components installed.
      - The script tries Veeam.Backup.PowerShell first, then VeeamPSSnapIn fallback.
      - It attempts to include internal/background sessions (for example SOBR/capacity-tier
        offload) via broader session cmdlets and core backup session fallback.

    PowerShell version requirements:
      - PowerShell 7.0 or later: the modern Veeam.Backup.PowerShell module is loaded.
      - Windows PowerShell 5.1 (Desktop edition): the modern module manifest declares a
        minimum PS version of 7.0 and cannot be loaded by PS 5.1. The script catches that
        failure and automatically falls back to the legacy VeeamPSSnapIn snap-in.
        Ensure VeeamPSSnapIn is registered (it is included with Veeam Backup & Replication
        console components) when running under Windows PowerShell 5.1.

    Run requirements:
    Run in an elevated PowerShell session on the Veeam Backup & Replication server or a host
    with the Veeam console/PowerShell components installed.
#>

[CmdletBinding()]
param(
    [ValidateRange(1, 8760)]
    [int]$Hours = 24,

    # Emit JSON Lines instead of readable text. Useful for later ingestion/parsing.
    [switch]$Json
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$script:Cutoff = (Get-Date).AddHours(-[Math]::Abs($Hours))
$script:SeenSessions = New-Object 'System.Collections.Generic.HashSet[string]'

function Import-VeeamPowerShell {
    [CmdletBinding()]
    param()

    $loaded = $false

    # Try the modern Veeam.Backup.PowerShell module first.
    # On Windows PowerShell 5.1 the module manifest may declare a minimum PS version of
    # 7.0, which causes Import-Module to throw.  Catch that failure and fall through to
    # the legacy VeeamPSSnapIn snap-in below.
    if (Get-Module -ListAvailable -Name 'Veeam.Backup.PowerShell' -ErrorAction SilentlyContinue) {
        try {
            Import-Module 'Veeam.Backup.PowerShell' -ErrorAction Stop
            $loaded = $true
        }
        catch {
            Write-Warning (("Could not import Veeam.Backup.PowerShell module: {0}  " +
                "Falling back to VeeamPSSnapIn (required on Windows PowerShell 5.1).") `
                -f $_.Exception.Message)
        }
    }

    if (-not $loaded) {
        $snapIn = Get-PSSnapin -Registered -Name 'VeeamPSSnapIn' -ErrorAction SilentlyContinue
        if ($snapIn) {
            Add-PSSnapin 'VeeamPSSnapIn' -ErrorAction Stop
            $loaded = $true
        }
    }

    if (-not $loaded) {
        throw (
            'Unable to load Veeam PowerShell components. ' +
            'PowerShell 7.0 or later can import the modern Veeam.Backup.PowerShell module. ' +
            'Windows PowerShell 5.1 requires the legacy VeeamPSSnapIn snap-in to be registered ' +
            '(included with the Veeam Backup & Replication console/PowerShell components). ' +
            'Install the Veeam console components and re-run this script on a VBR server or console machine.'
        )
    }
}

function Get-PropertyValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [object]$InputObject,
        [Parameter(Mandatory)] [string[]]$Names
    )

    foreach ($name in $Names) {
        $property = $InputObject.PSObject.Properties[$name]
        if ($null -ne $property -and $null -ne $property.Value) {
            return $property.Value
        }
    }

    return $null
}

function Get-ObjectIdentity {
    [CmdletBinding()]
    param([Parameter(Mandatory)] [object]$InputObject)

    $id = Get-PropertyValue -InputObject $InputObject -Names @('Id', 'Uid', 'SessionId')
    if ($null -ne $id) { return [string]$id }

    $name = Get-PropertyValue -InputObject $InputObject -Names @('Name', 'JobName')
    $start = Get-SessionStartTime -Session $InputObject
    return ('{0}|{1}|{2}' -f $InputObject.GetType().FullName, $name, $start)
}

function Get-SessionName {
    [CmdletBinding()]
    param([Parameter(Mandatory)] [object]$Session)

    $name = Get-PropertyValue -InputObject $Session -Names @('Name', 'JobName', 'SessionName')
    if ($null -ne $name) { return [string]$name }
    return '<unnamed>'
}

function Get-SessionType {
    [CmdletBinding()]
    param([Parameter(Mandatory)] [object]$Session)

    $type = Get-PropertyValue -InputObject $Session -Names @('JobType', 'SessionType', 'Type', 'Operation')
    if ($null -ne $type) { return [string]$type }
    return $Session.GetType().Name
}

function Get-SessionState {
    [CmdletBinding()]
    param([Parameter(Mandatory)] [object]$Session)

    $state = Get-PropertyValue -InputObject $Session -Names @('State', 'Status', 'Result')
    if ($null -ne $state) { return [string]$state }
    return '<unknown>'
}

function Get-SessionStartTime {
    [CmdletBinding()]
    param([Parameter(Mandatory)] [object]$Session)

    $value = Get-PropertyValue -InputObject $Session -Names @(
        'CreationTime',
        'CreationTimeLocal',
        'CreationTimeUTC',
        'StartTime',
        'StartTimeLocal',
        'StartTimeUTC'
    )

    if ($null -eq $value) { return $null }
    return [datetime]$value
}

function Get-SessionEndTime {
    [CmdletBinding()]
    param([Parameter(Mandatory)] [object]$Session)

    $value = Get-PropertyValue -InputObject $Session -Names @(
        'EndTime',
        'EndTimeLocal',
        'EndTimeUTC',
        'StopTime',
        'StopTimeLocal',
        'StopTimeUTC'
    )

    if ($null -eq $value) { return $null }
    return [datetime]$value
}

function Test-SessionInWindow {
    [CmdletBinding()]
    param([Parameter(Mandatory)] [object]$Session)

    $start = Get-SessionStartTime -Session $Session
    $end = Get-SessionEndTime -Session $Session

    # Include sessions that started in the window, ended in the window, or are still running.
    if ($null -ne $start -and $start -ge $script:Cutoff) { return $true }
    if ($null -ne $end -and $end -ge $script:Cutoff) { return $true }
    if ($null -ne $start -and $null -eq $end) { return $true }

    # If time metadata is not exposed, include it rather than silently losing possible internal sessions.
    return ($null -eq $start -and $null -eq $end)
}

function Get-SessionLogRecords {
    [CmdletBinding()]
    param([Parameter(Mandatory)] [object]$Session)

    $logger = Get-PropertyValue -InputObject $Session -Names @('Logger')
    if ($null -eq $logger) { return @() }

    try {
        $log = $logger.GetLog()
        if ($null -eq $log) { return @() }

        $records = Get-PropertyValue -InputObject $log -Names @('UpdatedRecords', 'Records')
        if ($null -ne $records) { return @($records) }
    }
    catch {
        return @()
    }

    return @()
}

function Get-RecordTime {
    [CmdletBinding()]
    param([Parameter(Mandatory)] [object]$Record)

    $value = Get-PropertyValue -InputObject $Record -Names @(
        'StartTime',
        'StartTimeLocal',
        'StartTimeUTC',
        'UpdateTime',
        'UpdateTimeLocal',
        'UpdateTimeUTC',
        'CreationTime',
        'Time'
    )

    if ($null -eq $value) { return $null }
    return [datetime]$value
}

function Get-RecordTitle {
    [CmdletBinding()]
    param([Parameter(Mandatory)] [object]$Record)

    $title = Get-PropertyValue -InputObject $Record -Names @('Title', 'Name', 'Text', 'Message')
    if ($null -ne $title) { return [string]$title }
    return '<no title>'
}

function Get-RecordDescription {
    [CmdletBinding()]
    param([Parameter(Mandatory)] [object]$Record)

    $description = Get-PropertyValue -InputObject $Record -Names @('Description', 'Details', 'FullMessage')
    if ($null -ne $description) { return [string]$description }
    return ''
}

function Get-RecordStatus {
    [CmdletBinding()]
    param([Parameter(Mandatory)] [object]$Record)

    $status = Get-PropertyValue -InputObject $Record -Names @('Status', 'Result', 'State')
    if ($null -ne $status) { return [string]$status }
    return '<unknown>'
}

function Write-LogEntry {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [object]$Session,
        [Parameter(Mandatory)] [object]$Record,
        [Parameter(Mandatory)] [string]$Source
    )

    $entry = [ordered]@{
        collected_at   = (Get-Date).ToString('o')
        source         = $Source
        session_id     = Get-ObjectIdentity -InputObject $Session
        session_name   = Get-SessionName -Session $Session
        session_type   = Get-SessionType -Session $Session
        session_state  = Get-SessionState -Session $Session
        session_start  = $(if ($null -ne (Get-SessionStartTime -Session $Session)) { (Get-SessionStartTime -Session $Session).ToString('o') } else { $null })
        session_end    = $(if ($null -ne (Get-SessionEndTime -Session $Session)) { (Get-SessionEndTime -Session $Session).ToString('o') } else { $null })
        record_time    = $(if ($null -ne (Get-RecordTime -Record $Record)) { (Get-RecordTime -Record $Record).ToString('o') } else { $null })
        record_status  = Get-RecordStatus -Record $Record
        record_title   = Get-RecordTitle -Record $Record
        record_details = Get-RecordDescription -Record $Record
    }

    if ($Json) {
        [pscustomobject]$entry | ConvertTo-Json -Compress -Depth 6
        return
    }

    Write-Output ('[{0}] [{1}] [{2}] {3}' -f $entry.record_time, $entry.source, $entry.record_status, $entry.record_title)
    Write-Output ('  Collected: {0}' -f $entry.collected_at)
    Write-Output ('  SessionId: {0}' -f $entry.session_id)
    Write-Output ('  Session : {0}' -f $entry.session_name)
    Write-Output ('  Type    : {0}' -f $entry.session_type)
    Write-Output ('  State   : {0}' -f $entry.session_state)
    Write-Output ('  Start   : {0}' -f $entry.session_start)
    Write-Output ('  End     : {0}' -f $entry.session_end)
    if (-not [string]::IsNullOrWhiteSpace($entry.record_details)) {
        Write-Output ('  Details : {0}' -f $entry.record_details)
    }
    Write-Output ''
}

function Write-SessionLogs {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [object]$Session,
        [Parameter(Mandatory)] [string]$Source
    )

    if (-not (Test-SessionInWindow -Session $Session)) { return }

    $sessionId = Get-ObjectIdentity -InputObject $Session
    if (-not $script:SeenSessions.Add($sessionId)) { return }

    $records = Get-SessionLogRecords -Session $Session

    if ($records.Count -eq 0) {
        $synthetic = [pscustomobject]@{
            StartTime   = Get-SessionStartTime -Session $Session
            Status      = Get-SessionState -Session $Session
            Title       = 'No detailed logger records exposed for this session.'
            Description = ''
        }
        Write-LogEntry -Session $Session -Record $synthetic -Source $Source
        return
    }

    foreach ($record in $records) {
        $recordTime = Get-RecordTime -Record $record
        if ($null -eq $recordTime -or $recordTime -ge $script:Cutoff) {
            Write-LogEntry -Session $Session -Record $record -Source $Source
        }
    }
}

function Get-RecentSessionsFromCmdlet {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string]$CmdletName,
        [Parameter(Mandatory)] [string]$Source
    )

    if (-not (Get-Command -Name $CmdletName -ErrorAction SilentlyContinue)) { return }

    try {
        & $CmdletName -ErrorAction Stop | ForEach-Object {
            Write-SessionLogs -Session $_ -Source $Source
        }
    }
    catch {
        Write-Warning ('Unable to collect sessions via {0}: {1}' -f $CmdletName, $_.Exception.Message)
    }
}

function Get-RecentJobSessions {
    [CmdletBinding()]
    param()

    if (-not (Get-Command -Name 'Get-VBRJob' -ErrorAction SilentlyContinue)) { return }

    try {
        $jobs = Get-VBRJob -ErrorAction Stop
        foreach ($job in $jobs) {
            try {
                $sessions = New-Object 'System.Collections.Generic.List[object]'

                if ($job.PSObject.Methods['GetSessions']) {
                    foreach ($session in @($job.GetSessions())) {
                        if ($null -ne $session) { [void]$sessions.Add($session) }
                    }
                }

                if ($job.PSObject.Methods['FindLastSessions']) {
                    foreach ($session in @($job.FindLastSessions())) {
                        if ($null -ne $session) { [void]$sessions.Add($session) }
                    }
                }

                if ($job.PSObject.Methods['FindLastSession']) {
                    $lastSession = $job.FindLastSession()
                    if ($null -ne $lastSession) { [void]$sessions.Add($lastSession) }
                }

                foreach ($session in $sessions) {
                    Write-SessionLogs -Session $session -Source 'Get-VBRJob'
                }
            }
            catch {
                Write-Warning ('Unable to collect sessions for job "{0}": {1}' -f $job.Name, $_.Exception.Message)
            }
        }
    }
    catch {
        Write-Warning ('Unable to enumerate jobs via Get-VBRJob: {0}' -f $_.Exception.Message)
    }
}

function Get-RecentCoreBackupSessions {
    [CmdletBinding()]
    param()

    try {
        $type = [Veeam.Backup.Core.CBackupSession]
    }
    catch {
        return
    }

    try {
        $sessions = @($type::GetAll())
        foreach ($session in $sessions) {
            Write-SessionLogs -Session $session -Source 'Veeam.Backup.Core.CBackupSession'
        }
    }
    catch {
        Write-Warning ('Unable to collect core backup sessions: {0}' -f $_.Exception.Message)
    }
}

function Write-CollectorHeader {
    [CmdletBinding()]
    param()

    if ($Json) { return }

    Write-Output 'Veeam Log Collector'
    Write-Output ('Window       : {0:o} through {1:o}' -f $script:Cutoff, (Get-Date))
    Write-Output ('Host         : {0}' -f $env:COMPUTERNAME)
    Write-Output ('PowerShell   : {0}' -f $PSVersionTable.PSVersion)
    Write-Output ''
}

Import-VeeamPowerShell
Write-CollectorHeader

# 1. Job-centric collection: all configured backup jobs and their recent sessions.
Get-RecentJobSessions

# 2. Public Veeam session cmdlets. These often include backup copy, replica, tape, agent,
#    restore, and infrastructure/internal sessions depending on VBR version.
Get-RecentSessionsFromCmdlet -CmdletName 'Get-VBRSession' -Source 'Get-VBRSession'
Get-RecentSessionsFromCmdlet -CmdletName 'Get-VBRBackupSession' -Source 'Get-VBRBackupSession'

# 3. Core session fallback. This is useful for internal processes such as SOBR/capacity-tier
#    offload sessions that may not be attached to a user-created backup job object.
Get-RecentCoreBackupSessions

if (-not $Json) {
    Write-Output ('Completed. Sessions examined/emitted: {0}' -f $script:SeenSessions.Count)
}
