<#
.SYNOPSIS
    Prints Veeam Backup & Replication session log records from the last N hours.

.DESCRIPTION
    Uses the Veeam Backup PowerShell module/snap-in to enumerate recent sessions for:
      - all backup jobs returned by Get-VBRJob
      - VBR sessions returned by Get-VBRSession (called per-job when -Job is mandatory,
        otherwise called generically)
      - backup sessions returned by Get-VBRBackupSession
      - internal/core sessions returned through Veeam.Backup.Core.CBackupSession, when available

    Progress and status messages are printed throughout so you can see what the script is
    doing while it runs.  In default (human-readable) mode these go to standard output
    interleaved with log records.  In -Json mode they are sent to the Warning stream so
    that standard output remains pure JSON Lines.

    This intentionally writes log records to standard output only.
    Redirect stdout if you want a file.

.PARAMETER Hours
    Time window (in hours) to collect sessions/log records from. Default is 24.

.PARAMETER Json
    Emit one JSON object per line (JSON Lines) for machine-readable output.
    Progress/status messages are sent to the Warning stream in this mode so that
    standard output remains parseable JSON Lines.

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
      - Human-readable mode (default, -Json not specified) prints timestamped progress and
        status messages to standard output so you can follow along in real time.
      - -Json mode routes all progress/status messages to the Warning stream; standard
        output contains only JSON Lines records suitable for piping or log ingestion.

    Get-VBRSession note:
      - On some Veeam versions Get-VBRSession requires a mandatory -Job parameter and will
        prompt interactively when called without arguments.  This script detects that
        condition automatically: if -Job is mandatory the cmdlet is called per-job inside
        the job-enumeration phase; if no mandatory Job parameter exists it is called
        generically as before.

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

    # Emit JSON Lines instead of readable text. Progress goes to Warning stream in this mode.
    [switch]$Json
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$script:Cutoff        = (Get-Date).AddHours(-[Math]::Abs($Hours))
$script:SeenSessions  = New-Object 'System.Collections.Generic.HashSet[string]'
$script:EmittedCount  = 0

# ---------------------------------------------------------------------------
# Write-ProgressMessage
#   Timestamped status/progress output visible to the operator at all times.
#   - Human-readable mode: goes to standard output (stream 1).
#   - Json mode: goes to the Warning stream (stream 3) so stdout stays pure JSON.
# ---------------------------------------------------------------------------
function Write-ProgressMessage {
    [CmdletBinding()]
    param([Parameter(Mandatory)] [string]$Message)

    $line = '[{0:yyyy-MM-dd HH:mm:ss}] {1}' -f (Get-Date), $Message

    if ($Json) {
        Write-Warning $line
    } else {
        Write-Output $line
    }
}

# ---------------------------------------------------------------------------
# Test-CmdletHasMandatoryJobParam
#   Returns $true when every parameter set of the named cmdlet declares -Job
#   as mandatory, meaning the cmdlet cannot be safely called without a -Job
#   argument (doing so would trigger an interactive "Supply values" prompt).
# ---------------------------------------------------------------------------
function Test-CmdletHasMandatoryJobParam {
    [CmdletBinding()]
    param([Parameter(Mandatory)] [string]$CmdletName)

    $cmd = Get-Command -Name $CmdletName -ErrorAction SilentlyContinue
    if ($null -eq $cmd -or $cmd.ParameterSets.Count -eq 0) { return $false }

    # Walk every parameter set.  If we find even one set where Job is NOT
    # mandatory, the cmdlet can be invoked without arguments — return $false.
    foreach ($paramSet in $cmd.ParameterSets) {
        $jobMandatoryInSet = $false
        foreach ($param in $paramSet.Parameters) {
            if ($param.Name -eq 'Job' -and $param.IsMandatory) {
                $jobMandatoryInSet = $true
                break
            }
        }
        if (-not $jobMandatoryInSet) { return $false }
    }

    # Every parameter set requires -Job.
    return $true
}

function Import-VeeamPowerShell {
    [CmdletBinding()]
    param()

    Write-ProgressMessage "PowerShell $($PSVersionTable.PSEdition) $($PSVersionTable.PSVersion) on host: $env:COMPUTERNAME"

    $loaded = $false

    # Try the modern Veeam.Backup.PowerShell module first.
    # On Windows PowerShell 5.1 the module manifest may declare a minimum PS version of
    # 7.0, which causes Import-Module to throw.  Catch that failure and fall through to
    # the legacy VeeamPSSnapIn snap-in below.
    Write-ProgressMessage 'Attempting to load modern module: Veeam.Backup.PowerShell ...'
    if (Get-Module -ListAvailable -Name 'Veeam.Backup.PowerShell' -ErrorAction SilentlyContinue) {
        try {
            Import-Module 'Veeam.Backup.PowerShell' -ErrorAction Stop
            $loaded = $true
            Write-ProgressMessage 'Modern module Veeam.Backup.PowerShell loaded successfully.'
        }
        catch {
            Write-ProgressMessage ('  Modern module load failed: {0}' -f $_.Exception.Message)
            Write-Warning (("Could not import Veeam.Backup.PowerShell module: {0}  " +
                "Falling back to VeeamPSSnapIn (required on Windows PowerShell 5.1).") `
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
        $script:EmittedCount++
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

    Write-ProgressMessage ('Checking cmdlet: {0}' -f $CmdletName)

    if (-not (Get-Command -Name $CmdletName -ErrorAction SilentlyContinue)) {
        Write-ProgressMessage ('  Skipped: {0} is not available on this system.' -f $CmdletName)
        return
    }

    # Guard against cmdlets that require a mandatory -Job parameter.  Calling such a
    # cmdlet without arguments causes an interactive "Supply values for the following
    # parameters: Job:" prompt that hangs unattended execution.
    if (Test-CmdletHasMandatoryJobParam -CmdletName $CmdletName) {
        Write-ProgressMessage ('  Skipped: {0} requires a mandatory -Job parameter. ' +
            'Sessions are collected per-job in Phase 1 instead.' -f $CmdletName)
        return
    }

    Write-ProgressMessage ('  Enumerating sessions via {0} ...' -f $CmdletName)
    $before = $script:SeenSessions.Count

    try {
        & $CmdletName -ErrorAction Stop | ForEach-Object {
            Write-SessionLogs -Session $_ -Source $Source
        }
    }
    catch {
        Write-Warning ('Unable to collect sessions via {0}: {1}' -f $CmdletName, $_.Exception.Message)
    }

    Write-ProgressMessage ('  {0}: {1} new unique session(s) processed.' -f $CmdletName, ($script:SeenSessions.Count - $before))
}

function Get-RecentJobSessions {
    [CmdletBinding()]
    param()

    Write-ProgressMessage 'Phase 1 — Job-centric collection (Get-VBRJob).'

    if (-not (Get-Command -Name 'Get-VBRJob' -ErrorAction SilentlyContinue)) {
        Write-ProgressMessage '  Get-VBRJob not available. Skipping job-centric collection.'
        return
    }

    # Determine up-front whether Get-VBRSession requires a mandatory -Job argument.
    # If it does, we call it per-job here rather than generically in Phase 2.
    $vbrSessionAvail    = $null -ne (Get-Command -Name 'Get-VBRSession' -ErrorAction SilentlyContinue)
    $vbrSessionNeedsJob = $vbrSessionAvail -and (Test-CmdletHasMandatoryJobParam -CmdletName 'Get-VBRSession')

    if ($vbrSessionAvail) {
        if ($vbrSessionNeedsJob) {
            Write-ProgressMessage '  Get-VBRSession requires -Job; will be called per-job inside this phase.'
        } else {
            Write-ProgressMessage '  Get-VBRSession does not require -Job; it will be called generically in Phase 2.'
        }
    } else {
        Write-ProgressMessage '  Get-VBRSession not available on this system.'
    }

    try {
        $jobs = @(Get-VBRJob -ErrorAction Stop)
        Write-ProgressMessage ('  Found {0} job(s).' -f $jobs.Count)

        $jobIndex = 0
        foreach ($job in $jobs) {
            $jobIndex++
            Write-ProgressMessage ('  Job {0}/{1}: {2}' -f $jobIndex, $jobs.Count, $job.Name)

            try {
                $sessions = New-Object 'System.Collections.Generic.List[object]'

                if ($job.PSObject.Methods['GetSessions']) {
                    Write-ProgressMessage '    Method: GetSessions'
                    foreach ($session in @($job.GetSessions())) {
                        if ($null -ne $session) { [void]$sessions.Add($session) }
                    }
                }

                if ($job.PSObject.Methods['FindLastSessions']) {
                    Write-ProgressMessage '    Method: FindLastSessions'
                    foreach ($session in @($job.FindLastSessions())) {
                        if ($null -ne $session) { [void]$sessions.Add($session) }
                    }
                }

                if ($job.PSObject.Methods['FindLastSession']) {
                    Write-ProgressMessage '    Method: FindLastSession'
                    $lastSession = $job.FindLastSession()
                    if ($null -ne $lastSession) { [void]$sessions.Add($lastSession) }
                }

                # If Get-VBRSession requires -Job, call it per-job here to avoid an
                # interactive parameter prompt when called with no arguments.
                if ($vbrSessionNeedsJob) {
                    try {
                        Write-ProgressMessage '    Method: Get-VBRSession -Job'
                        $vbrSessions = @(Get-VBRSession -Job $job -ErrorAction Stop)
                        Write-ProgressMessage ('    Get-VBRSession -Job returned {0} session(s).' -f $vbrSessions.Count)
                        foreach ($session in $vbrSessions) {
                            if ($null -ne $session) { [void]$sessions.Add($session) }
                        }
                    }
                    catch {
                        Write-Warning ('Unable to call Get-VBRSession for job "{0}": {1}' -f $job.Name, $_.Exception.Message)
                    }
                }

                Write-ProgressMessage ('    Processing {0} session(s) for this job ...' -f $sessions.Count)
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

    Write-ProgressMessage 'Phase 3 — Core backup session fallback (Veeam.Backup.Core.CBackupSession).'

    try {
        $type = [Veeam.Backup.Core.CBackupSession]
    }
    catch {
        Write-ProgressMessage '  Type Veeam.Backup.Core.CBackupSession not available. Skipping.'
        return
    }

    $before = $script:SeenSessions.Count
    Write-ProgressMessage '  Enumerating all core backup sessions ...'

    try {
        $sessions = @($type::GetAll())
        Write-ProgressMessage ('  Core sessions returned: {0}.' -f $sessions.Count)
        foreach ($session in $sessions) {
            Write-SessionLogs -Session $session -Source 'Veeam.Backup.Core.CBackupSession'
        }
    }
    catch {
        Write-Warning ('Unable to collect core backup sessions: {0}' -f $_.Exception.Message)
    }

    Write-ProgressMessage ('  Phase 3 complete: {0} new unique session(s) processed.' -f ($script:SeenSessions.Count - $before))
}

function Write-CollectorHeader {
    [CmdletBinding()]
    param()

    if ($Json) { return }

    Write-Output '============================================================'
    Write-Output 'Veeam Log Collector'
    Write-Output ('Window     : last {0} hour(s)  (since {1:o})' -f $Hours, $script:Cutoff)
    Write-Output ('Host       : {0}' -f $env:COMPUTERNAME)
    Write-Output ('PowerShell : {0} {1}' -f $PSVersionTable.PSEdition, $PSVersionTable.PSVersion)
    Write-Output '============================================================'
    Write-Output ''
}

Write-ProgressMessage ('Veeam Log Collector starting. Window: last {0} hour(s) (since {1:o}).' -f $Hours, $script:Cutoff)
Import-VeeamPowerShell
Write-CollectorHeader

# 1. Job-centric collection: all configured backup jobs and their recent sessions.
#    Also calls Get-VBRSession -Job per-job when that cmdlet requires -Job.
Get-RecentJobSessions

# 2. Public Veeam session cmdlets. These often include backup copy, replica, tape, agent,
#    restore, and infrastructure/internal sessions depending on VBR version.
#    Each cmdlet is inspected for mandatory parameters before calling; any cmdlet that
#    would require interactive input is skipped automatically.
Write-ProgressMessage 'Phase 2 — Broad session cmdlet enumeration.'
Get-RecentSessionsFromCmdlet -CmdletName 'Get-VBRSession' -Source 'Get-VBRSession'
Get-RecentSessionsFromCmdlet -CmdletName 'Get-VBRBackupSession' -Source 'Get-VBRBackupSession'

# 3. Core session fallback. This is useful for internal processes such as SOBR/capacity-tier
#    offload sessions that may not be attached to a user-created backup job object.
Get-RecentCoreBackupSessions

Write-ProgressMessage ('Collection complete. Unique sessions: {0}  Records emitted: {1}.' -f $script:SeenSessions.Count, $script:EmittedCount)
