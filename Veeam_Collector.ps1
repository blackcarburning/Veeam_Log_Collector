<#
.SYNOPSIS
    Collects Veeam Backup & Replication session log records and exports a Veeam
    log bundle for the last N hours.

.DESCRIPTION
    Uses the Veeam Backup PowerShell module/snap-in to:
      1. Enumerate recent sessions for all job types and emit log records:
           - Regular backup jobs via Get-VBRJob + Get-VBRSession -Job (per-job,
             to avoid interactive prompts on Veeam versions where Get-VBRSession
             requires a -Job argument).
           - Computer/agent backup jobs via Get-VBRComputerBackupJob (when available),
             avoiding the deprecation warning emitted by Get-VBRJob for those jobs.
           - Broad VBR session cmdlets (Get-VBRBackupSession, etc.).
           - Internal/background sessions via Veeam.Backup.Core.CBackupSession when
             available (covers SOBR/capacity-tier offload and similar).
      2. Call Export-VBRLogs to produce an archive/bundle of Veeam log files.

    Progress and status messages are printed throughout so you can follow along.
    In default (human-readable) mode these go directly to the console (Write-Host /
    non-success stream) and never contaminate function return values.
    In -Json mode they are sent to the Warning stream so that standard output
    remains pure JSON Lines.

.PARAMETER Hours
    Time window (in hours) to collect sessions/log records from and to export
    logs for. Default is 24 (last 24 hours).

.PARAMETER OutputPath
    Directory where the exported Veeam log bundle will be written.  If omitted,
    defaults to E:\VEEAM_LOGS\COLLECTOR.  The directory is created if it does not
    exist.  After a successful export, items in the collector output directory that
    are older than 2 days are automatically removed to keep the folder tidy.

.PARAMETER Json
    Emit one JSON object per line (JSON Lines) for machine-readable output.
    Session log records and the final export summary are emitted as JSON objects.
    Progress/status messages are sent to the Warning stream so that standard
    output remains parseable JSON Lines.

.EXAMPLE
    .\Veeam_Collector.ps1

    Collects session logs and exports Veeam logs for the last 24 hours.

.EXAMPLE
    .\Veeam_Collector.ps1 -Hours 48 -OutputPath C:\Temp\VeeamLogs

    Collects and exports logs for the last 48 hours to C:\Temp\VeeamLogs.

.EXAMPLE
    .\Veeam_Collector.ps1 -Json

    Emits JSON Lines on stdout (session records + export summary).

.NOTES
    Usage notes:
      - Run this script with PowerShell 7 on a Veeam Backup & Replication server or
        a host with the Veeam console/PowerShell components installed.
      - The script tries Veeam.Backup.PowerShell first, then VeeamPSSnapIn fallback.
      - Human-readable mode prints timestamped progress to the console (Write-Host).
        Progress messages are never written to the success/output stream so they
        cannot contaminate function return values.
      - -Json mode routes all progress/status messages to the Warning stream; standard
        output contains only JSON Lines records suitable for piping or log ingestion.

    Get-VBRSession note:
      - On some Veeam versions Get-VBRSession requires a -Job value and will prompt
        interactively when called without arguments. This script avoids that entirely:
        when Get-VBRSession exposes a -Job parameter, it is called as
        Get-VBRSession -Job <job> for each regular job and is never called generically.

    Computer/agent backup jobs:
      - Get-VBRComputerBackupJob is used when available so that Get-VBRJob is not
        asked to enumerate agent/computer jobs (which triggers a deprecation warning).

    PowerShell version requirements:
      - PowerShell 7.0 or later: the modern Veeam.Backup.PowerShell module is loaded.
      - Windows PowerShell 5.1 (Desktop edition): the modern module manifest declares a
        minimum PS version of 7.0 and cannot be loaded by PS 5.1. The script catches that
        failure and automatically falls back to the legacy VeeamPSSnapIn snap-in.
        If neither loading path succeeds under PS 5.1, the error message includes the
        exact command to re-run using pwsh.exe (PowerShell 7).

    Run requirements:
      Run in an elevated PowerShell session on the Veeam Backup & Replication server
      or a host with the Veeam console/PowerShell components installed.
#>

[CmdletBinding()]
param(
    [ValidateRange(1, 8760)]
    [int]$Hours = 24,

    # Directory where the exported log bundle will be written.
    # Defaults to a timestamped sub-directory under the system temp folder.
    [string]$OutputPath = '',

    # Emit JSON Lines instead of readable text. Progress goes to Warning stream.
    [switch]$Json
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$script:EndTime      = Get-Date
$script:StartTime    = $script:EndTime.AddHours(-[Math]::Abs($Hours))
$script:Cutoff       = $script:StartTime
$script:SeenSessions = New-Object 'System.Collections.Generic.HashSet[string]'
$script:EmittedCount = 0

# ---------------------------------------------------------------------------
# Write-ProgressMessage
#   Timestamped status/progress output visible to the operator at all times.
#   IMPORTANT: Uses Write-Host (human-readable mode) or Write-Warning (Json mode)
#   so it never writes to the success/output stream (stream 1).  This prevents
#   progress text from being captured when functions are called in assignments.
# ---------------------------------------------------------------------------
function Write-ProgressMessage {
    [CmdletBinding()]
    param([Parameter(Mandatory)] [string]$Message)

    $line = '[{0:yyyy-MM-dd HH:mm:ss}] {1}' -f (Get-Date), $Message

    if ($Json) {
        Write-Warning $line
    } else {
        Write-Host $line
    }
}

# ---------------------------------------------------------------------------
# Test-CmdletHasParameter
#   Returns $true when the named cmdlet exposes the named parameter.
# ---------------------------------------------------------------------------
function Test-CmdletHasParameter {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string]$CmdletName,
        [Parameter(Mandatory)] [string]$ParameterName
    )

    $cmd = Get-Command -Name $CmdletName -ErrorAction SilentlyContinue
    if ($null -eq $cmd) { return $false }

    return $cmd.Parameters.ContainsKey($ParameterName)
}

# ---------------------------------------------------------------------------
# Test-CmdletCanInvokeWithoutArguments
#   Returns $true when the named cmdlet has at least one parameter set with no
#   mandatory parameters (i.e. can be safely called with no arguments).
# ---------------------------------------------------------------------------
function Test-CmdletCanInvokeWithoutArguments {
    [CmdletBinding()]
    param([Parameter(Mandatory)] [string]$CmdletName)

    $cmd = Get-Command -Name $CmdletName -ErrorAction SilentlyContinue
    if ($null -eq $cmd) { return $false }
    if ($cmd.ParameterSets.Count -eq 0) { return $true }

    foreach ($paramSet in $cmd.ParameterSets) {
        $mandatoryParameters = @($paramSet.Parameters | Where-Object { $_.IsMandatory })
        if ($mandatoryParameters.Count -eq 0) { return $true }
    }

    return $false
}

# ---------------------------------------------------------------------------
# Import-VeeamPowerShell
#   Loads the Veeam Backup PowerShell module or legacy snap-in.
#   Enhanced guidance when running under Windows PowerShell 5.1.
# ---------------------------------------------------------------------------
function Import-VeeamPowerShell {
    [CmdletBinding()]
    param()

    $currentPowerShellEdition = if ($PSVersionTable.PSEdition) { $PSVersionTable.PSEdition } else { 'Desktop' }
    Write-ProgressMessage ('PowerShell {0} {1} on host: {2}' -f $currentPowerShellEdition, $PSVersionTable.PSVersion, $env:COMPUTERNAME)

    $loaded = $false

    # Try the modern Veeam.Backup.PowerShell module first.
    # On Windows PowerShell 5.1 the module manifest may declare a minimum PS version of
    # 7.0, which causes Import-Module to throw.  Catch that and fall through to the
    # legacy VeeamPSSnapIn snap-in below.
    Write-ProgressMessage 'Attempting to load modern module: Veeam.Backup.PowerShell ...'
    if (Get-Module -ListAvailable -Name 'Veeam.Backup.PowerShell' -ErrorAction SilentlyContinue) {
        try {
            Import-Module 'Veeam.Backup.PowerShell' -ErrorAction Stop
            $loaded = $true
            Write-ProgressMessage 'Modern module Veeam.Backup.PowerShell loaded successfully.'
        }
        catch {
            Write-ProgressMessage ('  Modern module load failed: {0}' -f $_.Exception.Message)
            Write-Warning (('Could not import Veeam.Backup.PowerShell module: {0}  ' +
                'Falling back to VeeamPSSnapIn (required on Windows PowerShell 5.1).') `
                -f $_.Exception.Message)
        }
    } else {
        Write-ProgressMessage '  Module Veeam.Backup.PowerShell not found in module path.'
    }

    if (-not $loaded) {
        Write-ProgressMessage 'Attempting to load legacy snap-in: VeeamPSSnapIn ...'
        $snapIn = Get-PSSnapin -Registered -Name 'VeeamPSSnapIn' -ErrorAction SilentlyContinue
        if ($snapIn) {
            Add-PSSnapin 'VeeamPSSnapIn' -ErrorAction Stop
            $loaded = $true
            Write-ProgressMessage 'Legacy snap-in VeeamPSSnapIn loaded successfully.'
        } else {
            Write-ProgressMessage '  Snap-in VeeamPSSnapIn not found or not registered.'
        }
    }

    if (-not $loaded) {
        # Build an actionable error message.  When running under Windows PowerShell 5.1,
        # detect whether pwsh.exe is available and suggest the exact re-launch command.
        $isDesktopEdition = ($PSVersionTable.PSEdition -eq 'Desktop') -or
                            ($PSVersionTable.PSVersion.Major -le 5)

        $pwshSuggestion = ''
        if ($isDesktopEdition) {
            $pwshExe = Get-Command -Name 'pwsh.exe' -ErrorAction SilentlyContinue
            if ($null -ne $pwshExe) {
                $pwshSuggestion = (
                    '  PowerShell 7 (pwsh.exe) was found at: {0}{1}' +
                    '  Re-run the script with PowerShell 7:{1}' +
                    '    pwsh.exe -ExecutionPolicy Bypass -File "{2}"'
                ) -f $pwshExe.Source, [Environment]::NewLine, $PSCommandPath
            } else {
                $pwshSuggestion = (
                    '  You are running Windows PowerShell {0}. ' +
                    'Install PowerShell 7 from https://aka.ms/powershell and re-run:{1}' +
                    '    pwsh.exe -ExecutionPolicy Bypass -File "{2}"'
                ) -f $PSVersionTable.PSVersion, [Environment]::NewLine, $PSCommandPath
            }
        }

        $errorMessage = (
            'Unable to load Veeam PowerShell components. ' +
            'PowerShell 7.0 or later can import the modern Veeam.Backup.PowerShell module. ' +
            'Windows PowerShell 5.1 requires the legacy VeeamPSSnapIn snap-in to be registered ' +
            '(included with the Veeam Backup & Replication console/PowerShell components). ' +
            'Install the Veeam console components and re-run this script on a VBR server or console machine.'
        )

        if ($pwshSuggestion -ne '') {
            $errorMessage = $errorMessage + [Environment]::NewLine + $pwshSuggestion
        }

        throw $errorMessage
    }
}

# ---------------------------------------------------------------------------
# Get-PropertyValue
#   Returns the value of the first matching property name found on an object.
# ---------------------------------------------------------------------------
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
        'CreationTime', 'CreationTimeLocal', 'CreationTimeUTC',
        'StartTime', 'StartTimeLocal', 'StartTimeUTC'
    )

    if ($null -eq $value) { return $null }
    return [datetime]$value
}

function Get-SessionEndTime {
    [CmdletBinding()]
    param([Parameter(Mandatory)] [object]$Session)

    $value = Get-PropertyValue -InputObject $Session -Names @(
        'EndTime', 'EndTimeLocal', 'EndTimeUTC',
        'StopTime', 'StopTimeLocal', 'StopTimeUTC'
    )

    if ($null -eq $value) { return $null }
    return [datetime]$value
}

function Test-SessionInWindow {
    [CmdletBinding()]
    param([Parameter(Mandatory)] [object]$Session)

    $start = Get-SessionStartTime -Session $Session
    $end   = Get-SessionEndTime   -Session $Session

    if ($null -ne $start -and $start -ge $script:Cutoff) { return $true }
    if ($null -ne $end   -and $end   -ge $script:Cutoff) { return $true }
    if ($null -ne $start -and $null -eq $end)            { return $true }

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
        'StartTime', 'StartTimeLocal', 'StartTimeUTC',
        'UpdateTime', 'UpdateTimeLocal', 'UpdateTimeUTC',
        'CreationTime', 'Time'
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

    $sessionStart = Get-SessionStartTime -Session $Session
    $sessionEnd   = Get-SessionEndTime   -Session $Session
    $recordTime   = Get-RecordTime       -Record  $Record

    $entry = [ordered]@{
        collected_at   = (Get-Date).ToString('o')
        source         = $Source
        session_id     = Get-ObjectIdentity -InputObject $Session
        session_name   = Get-SessionName    -Session $Session
        session_type   = Get-SessionType    -Session $Session
        session_state  = Get-SessionState   -Session $Session
        session_start  = $(if ($null -ne $sessionStart) { $sessionStart.ToString('o') } else { $null })
        session_end    = $(if ($null -ne $sessionEnd)   { $sessionEnd.ToString('o')   } else { $null })
        record_time    = $(if ($null -ne $recordTime)   { $recordTime.ToString('o')   } else { $null })
        record_status  = Get-RecordStatus      -Record $Record
        record_title   = Get-RecordTitle       -Record $Record
        record_details = Get-RecordDescription -Record $Record
    }

    if ($Json) {
        [pscustomobject]$entry | ConvertTo-Json -Compress -Depth 6
        $script:EmittedCount++
        return
    }

    Write-Output ('[{0}] [{1}] [{2}] {3}' -f $entry.record_time, $entry.source, $entry.record_status, $entry.record_title)
    Write-Output ('  Collected: {0}' -f $entry.collected_at)
    Write-Output ('  SessionId: {0}' -f $entry.session_id)
    Write-Output ('  Session  : {0}' -f $entry.session_name)
    Write-Output ('  Type     : {0}' -f $entry.session_type)
    Write-Output ('  State    : {0}' -f $entry.session_state)
    Write-Output ('  Start    : {0}' -f $entry.session_start)
    Write-Output ('  End      : {0}' -f $entry.session_end)
    if (-not [string]::IsNullOrWhiteSpace($entry.record_details)) {
        Write-Output ('  Details  : {0}' -f $entry.record_details)
    }
    Write-Output ''
    $script:EmittedCount++
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

    # Force the result into an array before checking .Count.
    # PowerShell may unwrap a single log record into a scalar; under Set-StrictMode
    # reading .Count from a scalar throws "The property 'Count' cannot be found".
    $records = @(Get-SessionLogRecords -Session $Session)

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

# ---------------------------------------------------------------------------
# Invoke-JobSessionMethods
#   Collects sessions for a single job object using available methods/cmdlets.
#   Handles both regular VBR jobs and computer backup jobs defensively.
# ---------------------------------------------------------------------------
function Invoke-JobSessionMethods {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [object]$Job,
        [Parameter(Mandatory)] [string]$Source,
        [bool]$UseVBRSessionPerJob = $false
    )

    $sessions = New-Object 'System.Collections.Generic.List[object]'

    if ($Job.PSObject.Methods['GetSessions']) {
        Write-ProgressMessage '    Method: GetSessions'
        try {
            foreach ($s in @($Job.GetSessions())) {
                if ($null -ne $s) { [void]$sessions.Add($s) }
            }
        } catch {
            Write-Warning ('    GetSessions failed: {0}' -f $_.Exception.Message)
        }
    }

    if ($Job.PSObject.Methods['FindLastSessions']) {
        Write-ProgressMessage '    Method: FindLastSessions'
        try {
            foreach ($s in @($Job.FindLastSessions())) {
                if ($null -ne $s) { [void]$sessions.Add($s) }
            }
        } catch {
            Write-Warning ('    FindLastSessions failed: {0}' -f $_.Exception.Message)
        }
    }

    if ($Job.PSObject.Methods['FindLastSession']) {
        Write-ProgressMessage '    Method: FindLastSession'
        try {
            $lastSession = $Job.FindLastSession()
            if ($null -ne $lastSession) { [void]$sessions.Add($lastSession) }
        } catch {
            Write-Warning ('    FindLastSession failed: {0}' -f $_.Exception.Message)
        }
    }

    if ($UseVBRSessionPerJob) {
        try {
            Write-ProgressMessage ('    Method: Get-VBRSession -Job "{0}"' -f $Job.Name)
            $vbrSessions = @(Get-VBRSession -Job $Job -ErrorAction Stop)
            Write-ProgressMessage ('    Get-VBRSession -Job returned {0} session(s).' -f $vbrSessions.Count)
            foreach ($s in $vbrSessions) {
                if ($null -ne $s) { [void]$sessions.Add($s) }
            }
        } catch {
            Write-Warning ('    Get-VBRSession -Job "{0}" failed: {1}' -f $Job.Name, $_.Exception.Message)
        }
    }

    Write-ProgressMessage ('    Processing {0} session(s) for this job ...' -f $sessions.Count)
    foreach ($session in $sessions) {
        Write-SessionLogs -Session $session -Source $Source
    }
}

# ---------------------------------------------------------------------------
# Get-RecentJobSessions
#   Phase 1 — Regular VBR jobs via Get-VBRJob.
# ---------------------------------------------------------------------------
function Get-RecentJobSessions {
    [CmdletBinding()]
    param()

    Write-ProgressMessage 'Phase 1 — Job-centric collection (Get-VBRJob).'

    if (-not (Get-Command -Name 'Get-VBRJob' -ErrorAction SilentlyContinue)) {
        Write-ProgressMessage '  Get-VBRJob not available. Skipping job-centric collection.'
        return
    }

    $vbrSessionAvail       = $null -ne (Get-Command -Name 'Get-VBRSession' -ErrorAction SilentlyContinue)
    $vbrSessionHasJobParam = $vbrSessionAvail -and (Test-CmdletHasParameter -CmdletName 'Get-VBRSession' -ParameterName 'Job')

    if ($vbrSessionAvail) {
        if ($vbrSessionHasJobParam) {
            Write-ProgressMessage '  Get-VBRSession exposes -Job; will call Get-VBRSession -Job for each job.'
        } else {
            Write-ProgressMessage '  Get-VBRSession does not expose -Job; it may be called generically in Phase 3 if safe.'
        }
    } else {
        Write-ProgressMessage '  Get-VBRSession not available on this system.'
    }

    try {
        $jobs = @(Get-VBRJob -ErrorAction Stop -WarningAction SilentlyContinue)
        Write-ProgressMessage ('  Found {0} job(s).' -f $jobs.Count)

        $jobIndex = 0
        foreach ($job in $jobs) {
            $jobIndex++
            Write-ProgressMessage ('  Job {0}/{1}: {2}' -f $jobIndex, $jobs.Count, $job.Name)

            try {
                Invoke-JobSessionMethods -Job $job -Source 'Get-VBRJob' -UseVBRSessionPerJob $vbrSessionHasJobParam
            } catch {
                Write-Warning ('Unable to collect sessions for job "{0}": {1}' -f $job.Name, $_.Exception.Message)
            }
        }
    } catch {
        Write-Warning ('Unable to enumerate jobs via Get-VBRJob: {0}' -f $_.Exception.Message)
    }
}

# ---------------------------------------------------------------------------
# Get-RecentComputerBackupJobSessions
#   Phase 2 — Computer/agent backup jobs via Get-VBRComputerBackupJob.
#   Using the dedicated cmdlet avoids the deprecation warning that Get-VBRJob
#   emits for computer backup jobs on newer Veeam versions.
# ---------------------------------------------------------------------------
function Get-RecentComputerBackupJobSessions {
    [CmdletBinding()]
    param()

    Write-ProgressMessage 'Phase 2 — Computer/agent backup jobs (Get-VBRComputerBackupJob).'

    if (-not (Get-Command -Name 'Get-VBRComputerBackupJob' -ErrorAction SilentlyContinue)) {
        Write-ProgressMessage '  Get-VBRComputerBackupJob not available on this system. Skipping.'
        return
    }

    Write-ProgressMessage '  Get-VBRComputerBackupJob is available.'

    try {
        $computerJobs = @(Get-VBRComputerBackupJob -ErrorAction Stop -WarningAction SilentlyContinue)
        Write-ProgressMessage ('  Found {0} computer backup job(s).' -f $computerJobs.Count)

        $jobIndex = 0
        foreach ($job in $computerJobs) {
            $jobIndex++
            $jobName = if ($null -ne $job.PSObject.Properties['Name']) { $job.Name } else { '<unnamed>' }
            Write-ProgressMessage ('  Computer job {0}/{1}: {2}' -f $jobIndex, $computerJobs.Count, $jobName)

            try {
                Invoke-JobSessionMethods -Job $job -Source 'Get-VBRComputerBackupJob' -UseVBRSessionPerJob $false
            } catch {
                Write-Warning ('Unable to collect sessions for computer backup job "{0}": {1}' -f $jobName, $_.Exception.Message)
            }
        }
    } catch {
        Write-Warning ('Unable to enumerate computer backup jobs via Get-VBRComputerBackupJob: {0}' -f $_.Exception.Message)
    }
}

# ---------------------------------------------------------------------------
# Get-RecentSessionsFromCmdlet
#   Runs a session-enumeration cmdlet without arguments (when safe) and
#   processes any sessions returned.
# ---------------------------------------------------------------------------
function Get-RecentSessionsFromCmdlet {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string]$CmdletName,
        [Parameter(Mandatory)] [string]$Source
    )

    Write-ProgressMessage ('Checking cmdlet: {0}' -f $CmdletName)

    if (-not (Get-Command -Name $CmdletName -ErrorAction SilentlyContinue)) {
        Write-ProgressMessage ('  Skipped: {0} is not available on this system.' -f $CmdletName)
        return
    }

    if ($CmdletName -eq 'Get-VBRSession' -and (Test-CmdletHasParameter -CmdletName $CmdletName -ParameterName 'Job')) {
        Write-ProgressMessage (
            '  Skipped: {0} exposes -Job and will be collected per-job as Get-VBRSession -Job <job> ' +
            'to avoid interactive prompts.' -f $CmdletName
        )
        return
    }

    if (-not (Test-CmdletCanInvokeWithoutArguments -CmdletName $CmdletName)) {
        Write-ProgressMessage ('  Skipped: {0} has mandatory parameter(s) and cannot be safely called without arguments.' -f $CmdletName)
        return
    }

    Write-ProgressMessage ('  Enumerating sessions via {0} ...' -f $CmdletName)
    $before = $script:SeenSessions.Count

    try {
        & $CmdletName -ErrorAction Stop | ForEach-Object {
            Write-SessionLogs -Session $_ -Source $Source
        }
    } catch {
        Write-Warning ('Unable to collect sessions via {0}: {1}' -f $CmdletName, $_.Exception.Message)
    }

    Write-ProgressMessage ('  {0}: {1} new unique session(s) processed.' -f $CmdletName, ($script:SeenSessions.Count - $before))
}

# ---------------------------------------------------------------------------
# Get-RecentCoreBackupSessions
#   Phase 4 — Internal/background sessions via Veeam.Backup.Core.CBackupSession.
#   Covers SOBR/capacity-tier offload and similar background sessions that may
#   not be visible through the public PowerShell cmdlets.
# ---------------------------------------------------------------------------
function Get-RecentCoreBackupSessions {
    [CmdletBinding()]
    param()

    Write-ProgressMessage 'Phase 4 — Core backup session fallback (Veeam.Backup.Core.CBackupSession).'

    try {
        $type = [Veeam.Backup.Core.CBackupSession]
    } catch {
        Write-ProgressMessage '  Type Veeam.Backup.Core.CBackupSession not available. Skipping.'
        return
    }

    $before = $script:SeenSessions.Count
    Write-ProgressMessage '  Enumerating all core backup sessions ...'

    try {
        $coreSessions = @($type::GetAll())
        Write-ProgressMessage ('  Core sessions returned: {0}.' -f $coreSessions.Count)
        foreach ($session in $coreSessions) {
            Write-SessionLogs -Session $session -Source 'Veeam.Backup.Core.CBackupSession'
        }
    } catch {
        Write-Warning ('Unable to collect core backup sessions: {0}' -f $_.Exception.Message)
    }

    Write-ProgressMessage ('  Phase 4 complete: {0} new unique session(s) processed.' -f ($script:SeenSessions.Count - $before))
}

# ---------------------------------------------------------------------------
# Resolve-ExportOutputPath
#   Returns the resolved output directory as a scalar [string].
#   Creates the directory when it does not exist.
#   NOTE: Write-ProgressMessage uses Write-Host so progress messages do NOT
#   appear on the success stream and cannot contaminate this function's return value.
# ---------------------------------------------------------------------------
function Resolve-ExportOutputPath {
    [CmdletBinding()]
    param([string]$RequestedPath)

    if ([string]::IsNullOrWhiteSpace($RequestedPath)) {
        $RequestedPath = 'E:\VEEAM_LOGS\COLLECTOR'
    }

    if (-not (Test-Path -LiteralPath $RequestedPath -PathType Container)) {
        Write-ProgressMessage ('Creating output directory: {0}' -f $RequestedPath)
        [void](New-Item -ItemType Directory -Path $RequestedPath -Force -ErrorAction Stop)
    }

    # Explicit [string] cast ensures exactly one scalar string is returned,
    # even if Resolve-Path returns a provider-qualified path object.
    $resolved = [string](Resolve-Path -LiteralPath $RequestedPath).ProviderPath
    if ([string]::IsNullOrWhiteSpace($resolved)) {
        throw ('Resolve-ExportOutputPath: resolved path is null or empty for input: {0}' -f $RequestedPath)
    }
    return $resolved
}

# ---------------------------------------------------------------------------
# Invoke-VBRLogsExport
#   Validates that Export-VBRLogs is available, introspects its parameters at
#   runtime, and calls it with the appropriate time-window and path bindings
#   for the installed Veeam PowerShell version.
# ---------------------------------------------------------------------------
function Invoke-VBRLogsExport {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [datetime]$StartTime,
        [Parameter(Mandatory)] [datetime]$EndTime,
        [Parameter(Mandatory)] [string]$ResolvedOutputPath
    )

    $exportCmd = Get-Command -Name 'Export-VBRLogs' -ErrorAction SilentlyContinue
    if ($null -eq $exportCmd) {
        throw (
            'Export-VBRLogs is not available after loading Veeam PowerShell components. ' +
            'This cmdlet is included in Veeam Backup & Replication PowerShell starting with ' +
            'version 11.  Ensure that the installed Veeam console/PowerShell components ' +
            'include Export-VBRLogs, then re-run this script.'
        )
    }

    $availableParams = $exportCmd.Parameters.Keys
    Write-ProgressMessage ('Export-VBRLogs parameters available: {0}' -f ($availableParams -join ', '))

    $exportParams = @{}

    # --- Resolve output/path parameter ---
    # FolderPath is listed first because that is the parameter name used by the
    # Veeam 12.x Export-VBRLogs cmdlet.  The remaining names cover older/alternative
    # Veeam PowerShell versions.
    $pathParamCandidates = @('FolderPath', 'Path', 'Folder', 'OutputPath', 'TargetPath',
                              'DestinationPath', 'FilePath', 'ExportPath', 'Target',
                              'Destination', 'Directory')
    $pathParam = $pathParamCandidates | Where-Object { $availableParams -contains $_ } |
                 Select-Object -First 1
    if ($null -ne $pathParam) {
        Write-ProgressMessage ('  Binding output path via -{0}' -f $pathParam)
        $exportParams[$pathParam] = $ResolvedOutputPath
    } else {
        # Last-resort: try positional only when the first positional parameter of
        # Export-VBRLogs is confirmed to be a string/path type (i.e. position 0).
        $firstPositional = $exportCmd.Parameters.Values |
            Where-Object { $_.Attributes | Where-Object { $_ -is [System.Management.Automation.ParameterAttribute] -and $_.Position -eq 0 } } |
            Select-Object -First 1

        if ($null -ne $firstPositional) {
            Write-ProgressMessage ('  No recognised named path parameter found; passing output path positionally via position-0 parameter ''{0}''.' -f $firstPositional.Name)
            $exportParams['PositionalPath'] = $ResolvedOutputPath
        } else {
            throw (
                'Export-VBRLogs: no recognised output-path parameter was found and no position-0 ' +
                'parameter exists to accept the path positionally. ' +
                'Available parameters: ' + ($availableParams -join ', ')
            )
        }
    }

    # --- Resolve time-window parameters ---
    $durationHours = [int][Math]::Ceiling(($EndTime - $StartTime).TotalHours)

    $fromParamCandidates = @('From', 'StartTime', 'StartDate', 'Since', 'After',
                              'FromDate', 'Start', 'DateFrom', 'BeginTime', 'Begin')
    $toParamCandidates   = @('To', 'EndTime', 'EndDate', 'Until', 'Before',
                              'ToDate', 'End', 'DateTo', 'StopTime', 'Finish')

    $fromParam = $fromParamCandidates | Where-Object { $availableParams -contains $_ } |
                 Select-Object -First 1
    $toParam   = $toParamCandidates   | Where-Object { $availableParams -contains $_ } |
                 Select-Object -First 1

    if ($null -ne $fromParam -and $null -ne $toParam) {
        Write-ProgressMessage ('  Binding time window via -{0} / -{1}' -f $fromParam, $toParam)
        $exportParams[$fromParam] = $StartTime
        $exportParams[$toParam]   = $EndTime
    } elseif ($null -ne $fromParam) {
        Write-ProgressMessage ('  Binding start time via -{0} (no matching end-time parameter found)' -f $fromParam)
        $exportParams[$fromParam] = $StartTime
    } else {
        $durationParamCandidates = @('Last', 'Hours', 'LastHours', 'Duration',
                                      'TimeSpanHours', 'Period', 'HoursBack')
        $durationParam = $durationParamCandidates | Where-Object { $availableParams -contains $_ } |
                         Select-Object -First 1

        if ($null -ne $durationParam) {
            Write-ProgressMessage ('  Binding duration via -{0} {1}' -f $durationParam, $durationHours)
            $exportParams[$durationParam] = $durationHours
        } else {
            Write-ProgressMessage (
                '  No time-window parameter (From/To, StartTime/EndTime, Last/Hours, etc.) ' +
                'recognised in this version of Export-VBRLogs. ' +
                'Calling without a time-window filter; the export will cover the full log history.'
            )
        }
    }

    Write-ProgressMessage 'Calling Export-VBRLogs ...'

    try {
        if ($exportParams.ContainsKey('PositionalPath')) {
            $posPath = $exportParams['PositionalPath']
            $exportParams.Remove('PositionalPath')

            if ($exportParams.Count -gt 0) {
                $result = Export-VBRLogs $posPath @exportParams -ErrorAction Stop
            } else {
                $result = Export-VBRLogs $posPath -ErrorAction Stop
            }
        } else {
            $result = Export-VBRLogs @exportParams -ErrorAction Stop
        }
    } catch {
        throw ('Export-VBRLogs failed: {0}' -f $_.Exception.Message)
    }

    Write-ProgressMessage 'Export-VBRLogs completed successfully.'
    return $result
}

# ---------------------------------------------------------------------------
# Collect-ExportedPaths
#   Inspects the return value of Export-VBRLogs and returns a flat string array
#   of file/folder paths.  Falls back to the resolved output directory when the
#   cmdlet returns nothing.
# ---------------------------------------------------------------------------
function Collect-ExportedPaths {
    [CmdletBinding()]
    param(
        [AllowNull()] [object]$ExportResult,
        [Parameter(Mandatory)] [string]$ResolvedOutputPath
    )

    $paths = New-Object 'System.Collections.Generic.List[string]'

    if ($null -ne $ExportResult) {
        $resultItems = @($ExportResult)
        foreach ($item in $resultItems) {
            if ($item -is [string] -and -not [string]::IsNullOrWhiteSpace($item)) {
                $paths.Add($item)
                continue
            }
            foreach ($propName in @('FullName', 'Path', 'FilePath', 'FileName', 'Name')) {
                $val = $item.PSObject.Properties[$propName]
                if ($null -ne $val -and -not [string]::IsNullOrWhiteSpace([string]$val.Value)) {
                    $paths.Add([string]$val.Value)
                    break
                }
            }
        }
    }

    if ($paths.Count -eq 0) {
        $paths.Add($ResolvedOutputPath)
    }

    return $paths.ToArray()
}

# ---------------------------------------------------------------------------
# Remove-OldCollectorExports
#   Deletes files and sub-directories under $CollectorPath whose LastWriteTime
#   is older than $RetentionDays days.  Errors are treated as warnings so that
#   a successful collection/export is not rolled back by a cleanup failure.
#   Progress messages respect -Json mode (non-success stream only).
# ---------------------------------------------------------------------------
function Remove-OldCollectorExports {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string]$CollectorPath,
        [int]$RetentionDays = 2
    )

    if (-not (Test-Path -LiteralPath $CollectorPath -PathType Container)) {
        Write-ProgressMessage ('  Cleanup skipped: collector path does not exist: {0}' -f $CollectorPath)
        return
    }

    $cutoff   = (Get-Date).AddDays(-$RetentionDays)
    $removed  = 0
    $retained = 0

    Write-ProgressMessage ('  Cleanup: removing items in ''{0}'' older than {1} day(s) (cutoff: {2:o}).' `
        -f $CollectorPath, $RetentionDays, $cutoff)

    try {
        $children = Get-ChildItem -LiteralPath $CollectorPath -Force -ErrorAction Stop
    } catch {
        Write-Warning ('Cleanup: could not enumerate ''{0}'': {1}' -f $CollectorPath, $_.Exception.Message)
        return
    }

    foreach ($child in $children) {
        if ($child.LastWriteTime -lt $cutoff) {
            try {
                Remove-Item -LiteralPath $child.FullName -Recurse -Force -ErrorAction Stop
                $removed++
            } catch {
                Write-Warning ('Cleanup: failed to remove ''{0}'': {1}' -f $child.FullName, $_.Exception.Message)
                $retained++
            }
        } else {
            $retained++
        }
    }

    Write-ProgressMessage ('  Cleanup complete: {0} item(s) removed, {1} item(s) retained/skipped.' `
        -f $removed, $retained)
}


function Write-CollectorHeader {
    [CmdletBinding()]
    param()

    if ($Json) { return }

    Write-Output '============================================================'
    Write-Output 'Veeam Log Collector'
    Write-Output ('Window     : last {0} hour(s)  ({1:o} to {2:o})' -f $Hours, $script:StartTime, $script:EndTime)
    Write-Output ('Host       : {0}' -f $env:COMPUTERNAME)
    $currentPowerShellEdition = if ($PSVersionTable.PSEdition) { $PSVersionTable.PSEdition } else { 'Desktop' }
    Write-Output ('PowerShell : {0} {1}' -f $currentPowerShellEdition, $PSVersionTable.PSVersion)
    Write-Output '============================================================'
    Write-Output ''
}

# ===========================================================================
# Main
# ===========================================================================

Write-ProgressMessage ('Veeam Log Collector starting. Window: last {0} hour(s) ({1:o} to {2:o}).' `
    -f $Hours, $script:StartTime, $script:EndTime)

Import-VeeamPowerShell
Write-CollectorHeader

# ---------------------------------------------------------------------------
# Session log collection (Phases 1–4)
# ---------------------------------------------------------------------------
Get-RecentJobSessions

Get-RecentComputerBackupJobSessions

Write-ProgressMessage 'Phase 3 — Broad session cmdlet enumeration.'
Get-RecentSessionsFromCmdlet -CmdletName 'Get-VBRSession'       -Source 'Get-VBRSession'
Get-RecentSessionsFromCmdlet -CmdletName 'Get-VBRBackupSession' -Source 'Get-VBRBackupSession'

Get-RecentCoreBackupSessions

Write-ProgressMessage ('Session collection complete. Unique sessions: {0}  Records emitted: {1}.' `
    -f $script:SeenSessions.Count, $script:EmittedCount)

# ---------------------------------------------------------------------------
# Export-VBRLogs — create the Veeam diagnostic log bundle
# ---------------------------------------------------------------------------
Write-ProgressMessage 'Starting Veeam log bundle export (Export-VBRLogs) ...'

$resolvedOutputPath = [string](Resolve-ExportOutputPath -RequestedPath $OutputPath)
if ([string]::IsNullOrWhiteSpace($resolvedOutputPath)) {
    throw 'Failed to resolve a valid output path for Export-VBRLogs.'
}
Write-ProgressMessage ('Output path : {0}' -f $resolvedOutputPath)

$exportResult = Invoke-VBRLogsExport `
    -StartTime          $script:StartTime `
    -EndTime            $script:EndTime `
    -ResolvedOutputPath $resolvedOutputPath

$exportedPaths = Collect-ExportedPaths -ExportResult $exportResult -ResolvedOutputPath $resolvedOutputPath

# ---------------------------------------------------------------------------
# Final output
# ---------------------------------------------------------------------------
if ($Json) {
    $summary = [ordered]@{
        record_type = 'export_summary'
        status      = 'success'
        hours       = $Hours
        start_time  = $script:StartTime.ToString('o')
        end_time    = $script:EndTime.ToString('o')
        output_path = $resolvedOutputPath
        exported    = @($exportedPaths)
        sessions_seen    = $script:SeenSessions.Count
        records_emitted  = $script:EmittedCount
    }
    [pscustomobject]$summary | ConvertTo-Json -Compress -Depth 4
} else {
    Write-Output ''
    Write-Output '------------------------------------------------------------'
    Write-Output 'Collection and export complete.'
    Write-Output ('Window          : last {0} hour(s)  ({1:o}  to  {2:o})' -f $Hours, $script:StartTime, $script:EndTime)
    Write-Output ('Sessions seen   : {0}' -f $script:SeenSessions.Count)
    Write-Output ('Records emitted : {0}' -f $script:EmittedCount)
    Write-Output ''
    Write-Output 'Exported log bundle location(s):'
    foreach ($p in $exportedPaths) {
        Write-Output ('  {0}' -f $p)
    }
    Write-Output '------------------------------------------------------------'
}

# ---------------------------------------------------------------------------
# Post-export cleanup: remove collector exports older than 2 days.
# This runs after the final summary so that even in -Json mode the summary is
# already emitted before any cleanup messages appear on non-success streams.
# ---------------------------------------------------------------------------
Remove-OldCollectorExports -CollectorPath $resolvedOutputPath -RetentionDays 2
