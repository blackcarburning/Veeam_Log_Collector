<#
.SYNOPSIS
    Reports the last error text from the most recent session for every Veeam
    backup, replication, backup-copy, agent, and SOBR offload job.

.DESCRIPTION
    Uses the Veeam Backup PowerShell module/snap-in to enumerate all jobs of
    interest (backup, replication, backup copy, computer/agent jobs, and SOBR
    capacity-tier offload sessions) and, for each job, find the most recent
    session within the last N hours.  For that session it extracts the last
    error/warning text using a defensive, multi-fallback approach:

      1. $session.GetLastError() — primary documented API.
      2. $session.GetTaskSessions() — per-task details for failed/warning tasks.
      3. Logger records — only EFailed/EWarning entries (never the full log).

    The result is a compact, LLM-friendly report showing each job's status and
    last error text.  No log bundles are created; no Export-VBRLogs calls are
    made.

    Progress messages are printed throughout.  In default (human-readable) mode
    they go to the console (Write-Host).  In -Json mode they go to the Warning
    stream so that standard output remains pure JSON.

.PARAMETER Hours
    Time window in hours.  Only sessions whose end or start time falls within the
    last N hours are considered.  Default is 24.  Valid range: 1-8760.

.PARAMETER Json
    When set, emit results as a single JSON array on stdout.  Progress/status
    messages are sent to the Warning stream so stdout stays valid JSON.

.PARAMETER OnlyFailures
    When set, only include jobs whose most recent session result is Failed or
    Warning.  Successful/skipped jobs are omitted from the report.

.EXAMPLE
    .\Veeam_Collector.ps1

    Lists every backup/replication/offload job's most recent session in the last
    24 hours along with its status and any last error text.

.EXAMPLE
    .\Veeam_Collector.ps1 -Hours 48 -OnlyFailures

    Shows only jobs with a Failed or Warning last session in the last 48 hours.

.EXAMPLE
    .\Veeam_Collector.ps1 -Json

    Emits a JSON array on stdout suitable for piping to an LLM or jq.
    Progress messages appear on the Warning stream only.

.NOTES
    Usage notes:
      - Run this script with PowerShell 7 on a Veeam Backup & Replication server
        or a host with the Veeam console/PowerShell components installed.
      - The script tries Veeam.Backup.PowerShell first, then VeeamPSSnapIn fallback.
      - Human-readable mode prints timestamped progress to the console (Write-Host).
        Progress messages are never written to the success/output stream so they
        cannot contaminate function return values.
      - -Json mode routes all progress/status messages to the Warning stream; stdout
        contains only a single JSON array suitable for parsing with ConvertFrom-Json
        or jq.

    Computer/agent backup jobs:
      - Get-VBRComputerBackupJob is used when available so that Get-VBRJob is not
        asked to enumerate agent/computer jobs (which triggers a deprecation warning).

    SOBR capacity-tier offload:
      - Get-VBRCapacityTierSyncSession is used when available.  Absence of this
        cmdlet is handled gracefully.

    PowerShell version requirements:
      - PowerShell 7.0 or later: the modern Veeam.Backup.PowerShell module is loaded.
      - Windows PowerShell 5.1 (Desktop edition): the script catches any version
        mismatch and falls back to the legacy VeeamPSSnapIn snap-in.
        If neither path succeeds, the error message includes the exact command to
        re-run using pwsh.exe (PowerShell 7).

    Run requirements:
      Run in an elevated PowerShell session on the Veeam Backup & Replication server
      or a host with the Veeam console/PowerShell components installed.
#>

[CmdletBinding()]
param(
    [ValidateRange(1, 8760)]
    [int]$Hours = 24,

    # Emit a JSON array on stdout. Progress goes to Warning stream.
    [switch]$Json,

    # Only include jobs with Failed or Warning last session.
    [switch]$OnlyFailures
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Get-SortableTicks {
    [CmdletBinding()]
    param([object]$Value)

    if ($null -eq $Value) { return [long]0 }
    try {
        if ($Value -is [datetime]) { return [long]$Value.ToUniversalTime().Ticks }
        $s = [string]$Value
        if ([string]::IsNullOrWhiteSpace($s)) { return [long]0 }
        $parsed = [datetime]::Parse($s, [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::RoundtripKind)
        return [long]$parsed.ToUniversalTime().Ticks
    } catch {
        return [long]0
    }
}

$script:EndTime      = Get-Date
$script:StartTime    = $script:EndTime.AddHours(-[Math]::Abs($Hours))
$script:Cutoff       = $script:StartTime
$script:SeenSessions = New-Object 'System.Collections.Generic.HashSet[string]'

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

function Get-LastErrorText {
    [CmdletBinding()]
    param([Parameter(Mandatory)] [object]$Session)

    $messages = New-Object 'System.Collections.Generic.List[string]'

    # --- Approach 1: $session.GetLastError() ---
    if ($Session.PSObject.Methods['GetLastError']) {
        try {
            $err = $Session.GetLastError()
            if ($null -ne $err) {
                $text = [string]$err
                if (-not [string]::IsNullOrWhiteSpace($text)) {
                    return $text.Trim()
                }
            }
        } catch { }
    }

    # --- Approach 2: task sessions ---
    if ($Session.PSObject.Methods['GetTaskSessions']) {
        try {
            $tasks = @($Session.GetTaskSessions())
            foreach ($task in $tasks) {
                $taskResult = ''
                $taskResultProp = $task.PSObject.Properties['Result']
                if ($null -ne $taskResultProp) { $taskResult = [string]$taskResultProp.Value }
                $taskStateProp = $task.PSObject.Properties['State']
                if ($null -ne $taskStateProp -and [string]::IsNullOrWhiteSpace($taskResult)) {
                    $taskResult = [string]$taskStateProp.Value
                }

                $isBad = $taskResult -imatch 'Failed|Warning|Error'
                if (-not $isBad) { continue }

                if ($task.PSObject.Methods['GetLastError']) {
                    try {
                        $taskErr = $task.GetLastError()
                        if ($null -ne $taskErr) {
                            $t = [string]$taskErr
                            if (-not [string]::IsNullOrWhiteSpace($t)) {
                                [void]$messages.Add($t.Trim())
                                continue
                            }
                        }
                    } catch { }
                }

                if ($task.PSObject.Methods['GetDetails']) {
                    try {
                        $details = $task.GetDetails()
                        if ($null -ne $details) {
                            $t = [string]$details
                            if (-not [string]::IsNullOrWhiteSpace($t)) {
                                [void]$messages.Add($t.Trim())
                                continue
                            }
                        }
                    } catch { }
                }

                # fall back to Name/Title on the task
                $taskName = Get-PropertyValue -InputObject $task -Names @('Name', 'Title', 'ObjectName')
                if ($null -ne $taskName -and -not [string]::IsNullOrWhiteSpace([string]$taskName)) {
                    [void]$messages.Add(('{0}: {1}' -f [string]$taskName, $taskResult).Trim())
                }
            }
        } catch { }
    }

    if ($messages.Count -gt 0) {
        $unique = @($messages | Sort-Object -Unique)
        return ($unique -join '; ')
    }

    # --- Approach 3: logger records — only EFailed/EWarning entries ---
    $loggerProp = $Session.PSObject.Properties['Logger']
    if ($null -ne $loggerProp -and $null -ne $loggerProp.Value) {
        try {
            $log = $loggerProp.Value.GetLog()
            if ($null -ne $log) {
                $records = $null
                $updatedProp = $log.PSObject.Properties['UpdatedRecords']
                if ($null -ne $updatedProp) { $records = $updatedProp.Value }
                if ($null -eq $records) {
                    $recProp = $log.PSObject.Properties['Records']
                    if ($null -ne $recProp) { $records = $recProp.Value }
                }

                if ($null -ne $records) {
                    foreach ($rec in @($records)) {
                        $statusProp = $rec.PSObject.Properties['Status']
                        if ($null -eq $statusProp) { continue }
                        $statusVal = [string]$statusProp.Value
                        if ($statusVal -notmatch 'EFailed|EWarning|Failed|Warning|Error') { continue }

                        $title = Get-PropertyValue -InputObject $rec -Names @('Title', 'Name', 'Text', 'Message')
                        if ($null -ne $title -and -not [string]::IsNullOrWhiteSpace([string]$title)) {
                            [void]$messages.Add([string]$title.Trim())
                        }
                    }
                }
            }
        } catch { }
    }

    if ($messages.Count -gt 0) {
        $unique = @($messages | Sort-Object -Unique)
        return ($unique -join '; ')
    }

    return ''
}

# ---------------------------------------------------------------------------
# Get-ResultSeverityOrder
#   Returns a sort key for a result/status string: 0=Failed, 1=Warning, 2=other.
# ---------------------------------------------------------------------------
function Get-ResultSeverityOrder {
    [CmdletBinding()]
    param([string]$Result)
    if ($Result -imatch 'Failed|Fail|Error') { return [int]0 }
    if ($Result -imatch 'Warning|Warn')      { return [int]1 }
    return [int]2
}

# ---------------------------------------------------------------------------
# Build-JobReport
#   Given a session and metadata, builds one report [pscustomobject].
# ---------------------------------------------------------------------------
function Build-JobReport {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [object]$Session,
        [string]$JobName   = '',
        [string]$JobType   = '',
        [string]$Source    = ''
    )

    $name    = if ($JobName) { $JobName } else { Get-SessionName -Session $Session }
    $type    = if ($JobType) { $JobType } else { Get-SessionType -Session $Session }
    $result  = Get-SessionState    -Session $Session
    $start   = Get-SessionStartTime -Session $Session
    $end     = Get-SessionEndTime   -Session $Session
    $lastErr = Get-LastErrorText    -Session $Session

    return [pscustomobject][ordered]@{
        job_name   = $name
        job_type   = $type
        result     = $result
        start_time = if ($null -ne $start) { $start.ToString('o') } else { $null }
        end_time   = if ($null -ne $end)   { $end.ToString('o')   } else { $null }
        last_error = $lastErr
        source     = $Source
    }
}

# ---------------------------------------------------------------------------
# Add-JobReportFromJob
#   Finds the most recent in-window session for a job and appends a report
#   entry to $Results (passed by [ref]).  De-duplicates via $SeenSessions.
# ---------------------------------------------------------------------------
function Add-JobReportFromJob {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [object]$Job,
        [Parameter(Mandatory)] [string]$Source,
        [Parameter(Mandatory)] [ref]$Results
    )

    $jobName = if ($null -ne $Job.PSObject.Properties['Name']) { [string]$Job.Name } else { '<unnamed>' }
    $jobType = Get-SessionType -Session $Job

    # Collect candidate sessions from the job object.
    $candidates = New-Object 'System.Collections.Generic.List[object]'

    if ($Job.PSObject.Methods['FindLastSession']) {
        try {
            $s = $Job.FindLastSession()
            if ($null -ne $s) { [void]$candidates.Add($s) }
        } catch { }
    }

    if ($Job.PSObject.Methods['FindLastSessions']) {
        try {
            foreach ($s in @($Job.FindLastSessions())) {
                if ($null -ne $s) { [void]$candidates.Add($s) }
            }
        } catch { }
    }

    if ($Job.PSObject.Methods['GetSessions']) {
        try {
            foreach ($s in @($Job.GetSessions())) {
                if ($null -ne $s) { [void]$candidates.Add($s) }
            }
        } catch { }
    }

    # Also try Get-VBRBackupSession filtered by job when available.
    if (Get-Command -Name 'Get-VBRBackupSession' -ErrorAction SilentlyContinue) {
        try {
            $bsSessions = @(Get-VBRBackupSession -Job $Job -ErrorAction Stop)
            foreach ($s in $bsSessions) {
                if ($null -ne $s) { [void]$candidates.Add($s) }
            }
        } catch { }
    }

    # Filter to window and pick the most recent.
    $inWindow = @($candidates | Where-Object { $null -ne $_ -and (Test-SessionInWindow -Session $_) })
    if ($inWindow.Count -eq 0) { return }

    # Sort by end time descending, then start time descending; pick first.
    $sorted = @($inWindow | Sort-Object -Property {
        $e = Get-SessionEndTime   -Session $_
        $s = Get-SessionStartTime -Session $_
        if ($null -ne $e) { Get-SortableTicks -Value $e }
        elseif ($null -ne $s) { Get-SortableTicks -Value $s }
        else { [long]0 }
    } -Descending)

    $session = $sorted[0]

    $sessionId = Get-ObjectIdentity -InputObject $session
    if (-not $script:SeenSessions.Add($sessionId)) { return }

    $report = Build-JobReport -Session $session -JobName $jobName -JobType $jobType -Source $Source
    [void]$Results.Value.Add($report)
}

# ---------------------------------------------------------------------------
# Write-TextReport
#   Emits one compact block per report entry to stdout.
# ---------------------------------------------------------------------------
function Write-TextReport {
    [CmdletBinding()]
    param([Parameter(Mandatory)] [object[]]$Reports)

    foreach ($r in $Reports) {
        Write-Output ('Job      : {0}' -f $r.job_name)
        Write-Output ('Type     : {0}' -f $r.job_type)
        Write-Output ('Result   : {0}' -f $r.result)
        Write-Output ('End Time : {0}' -f $(if ($null -ne $r.end_time) { $r.end_time } else { '(running/unknown)' }))
        if (-not [string]::IsNullOrWhiteSpace($r.last_error)) {
            Write-Output ('Error    : {0}' -f $r.last_error)
        }
        Write-Output ''
    }
}

# ---------------------------------------------------------------------------
# Write-CollectorHeader
# ---------------------------------------------------------------------------
function Write-CollectorHeader {
    [CmdletBinding()]
    param()

    if ($Json) { return }

    Write-Output '============================================================'
    Write-Output 'Veeam Last-Error Report'
    Write-Output ('Window     : last {0} hour(s)  ({1:o} to {2:o})' -f $Hours, $script:StartTime, $script:EndTime)
    Write-Output ('Host       : {0}' -f $env:COMPUTERNAME)
    $ed = if ($PSVersionTable.PSEdition) { $PSVersionTable.PSEdition } else { 'Desktop' }
    Write-Output ('PowerShell : {0} {1}' -f $ed, $PSVersionTable.PSVersion)
    Write-Output '============================================================'
    Write-Output ''
}

# ===========================================================================
# Main
# ===========================================================================

Write-ProgressMessage ('Veeam Last-Error Report starting. Window: last {0} hour(s) ({1:o} to {2:o}).' `
    -f $Hours, $script:StartTime, $script:EndTime)

Import-VeeamPowerShell
Write-CollectorHeader

$allReports = New-Object 'System.Collections.Generic.List[object]'

# ---------------------------------------------------------------------------
# Phase 1 — Regular VBR jobs via Get-VBRJob
# ---------------------------------------------------------------------------
Write-ProgressMessage 'Phase 1 — Enumerating regular jobs (Get-VBRJob).'
if (Get-Command -Name 'Get-VBRJob' -ErrorAction SilentlyContinue) {
    try {
        $vbrJobs = @(Get-VBRJob -ErrorAction Stop -WarningAction SilentlyContinue)
        Write-ProgressMessage ('  Found {0} job(s) via Get-VBRJob.' -f $vbrJobs.Count)
        $idx = 0
        foreach ($job in $vbrJobs) {
            $idx++
            $jn = if ($null -ne $job.PSObject.Properties['Name']) { $job.Name } else { '<unnamed>' }
            Write-ProgressMessage ('  Job {0}/{1}: {2}' -f $idx, $vbrJobs.Count, $jn)
            try {
                Add-JobReportFromJob -Job $job -Source 'Get-VBRJob' -Results ([ref]$allReports)
            } catch {
                Write-Warning ('  Unable to process job "{0}": {1}' -f $jn, $_.Exception.Message)
            }
        }
    } catch {
        Write-Warning ('Unable to enumerate jobs via Get-VBRJob: {0}' -f $_.Exception.Message)
    }
} else {
    Write-ProgressMessage '  Get-VBRJob not available. Skipping.'
}

# ---------------------------------------------------------------------------
# Phase 2 — Computer/agent backup jobs via Get-VBRComputerBackupJob
# ---------------------------------------------------------------------------
Write-ProgressMessage 'Phase 2 — Computer/agent backup jobs (Get-VBRComputerBackupJob).'
if (Get-Command -Name 'Get-VBRComputerBackupJob' -ErrorAction SilentlyContinue) {
    try {
        $computerJobs = @(Get-VBRComputerBackupJob -ErrorAction Stop -WarningAction SilentlyContinue)
        Write-ProgressMessage ('  Found {0} computer backup job(s).' -f $computerJobs.Count)
        $idx = 0
        foreach ($job in $computerJobs) {
            $idx++
            $jn = if ($null -ne $job.PSObject.Properties['Name']) { $job.Name } else { '<unnamed>' }
            Write-ProgressMessage ('  Computer job {0}/{1}: {2}' -f $idx, $computerJobs.Count, $jn)
            try {
                Add-JobReportFromJob -Job $job -Source 'Get-VBRComputerBackupJob' -Results ([ref]$allReports)
            } catch {
                Write-Warning ('  Unable to process computer backup job "{0}": {1}' -f $jn, $_.Exception.Message)
            }
        }
    } catch {
        Write-Warning ('Unable to enumerate computer backup jobs: {0}' -f $_.Exception.Message)
    }
} else {
    Write-ProgressMessage '  Get-VBRComputerBackupJob not available. Skipping.'
}

# ---------------------------------------------------------------------------
# Phase 3 — SOBR capacity-tier offload sessions
# ---------------------------------------------------------------------------
Write-ProgressMessage 'Phase 3 — SOBR capacity-tier offload (Get-VBRCapacityTierSyncSession).'
if (Get-Command -Name 'Get-VBRCapacityTierSyncSession' -ErrorAction SilentlyContinue) {
    try {
        $sobrSessions = @(Get-VBRCapacityTierSyncSession -ErrorAction Stop)
        Write-ProgressMessage ('  Found {0} capacity-tier session(s).' -f $sobrSessions.Count)
        $inWindow = @($sobrSessions | Where-Object { Test-SessionInWindow -Session $_ })
        Write-ProgressMessage ('  {0} session(s) within window.' -f $inWindow.Count)

        # For SOBR sessions there is no parent job object — report per session.
        # Group by name to pick the most-recent per named offload job.
        $grouped = @{}
        foreach ($s in $inWindow) {
            $sName = Get-SessionName -Session $s
            $sEnd  = Get-SessionEndTime -Session $s
            $sTime = if ($null -ne $sEnd) { Get-SortableTicks -Value $sEnd } else {
                $st = Get-SessionStartTime -Session $s
                if ($null -ne $st) { Get-SortableTicks -Value $st } else { [long]0 }
            }
            if (-not $grouped.ContainsKey($sName)) {
                $grouped[$sName] = @{ Session = $s; Time = $sTime }
            } elseif ($sTime -gt $grouped[$sName].Time) {
                $grouped[$sName] = @{ Session = $s; Time = $sTime }
            }
        }

        foreach ($entry in $grouped.Values) {
            $s = $entry.Session
            $sessionId = Get-ObjectIdentity -InputObject $s
            if (-not $script:SeenSessions.Add($sessionId)) { continue }
            $report = Build-JobReport -Session $s -JobType 'CapacityTierSync' -Source 'Get-VBRCapacityTierSyncSession'
            [void]$allReports.Add($report)
        }
    } catch {
        Write-Warning ('Unable to enumerate capacity-tier sessions: {0}' -f $_.Exception.Message)
    }
} else {
    Write-ProgressMessage '  Get-VBRCapacityTierSyncSession not available. Skipping.'
}

Write-ProgressMessage ('Enumeration complete. Total report entries before filtering: {0}.' -f $allReports.Count)

# ---------------------------------------------------------------------------
# Apply -OnlyFailures filter
# ---------------------------------------------------------------------------
$filtered = if ($OnlyFailures) {
    @($allReports | Where-Object { $_.result -imatch 'Failed|Warning|Warn|Error' })
} else {
    @($allReports)
}

# ---------------------------------------------------------------------------
# Sort: Failed first (0), Warning (1), other (2); then end_time descending.
# ---------------------------------------------------------------------------
$sorted = @($filtered | Sort-Object -Property @(
    @{ Expression = { [int](Get-ResultSeverityOrder -Result $_.result) }; Descending = $false },
    @{ Expression = { Get-SortableTicks -Value $_.end_time }; Descending = $true }
))

# ---------------------------------------------------------------------------
# Compute summary counts
# ---------------------------------------------------------------------------
$totalJobs   = $sorted.Count
$failedCount = @($sorted | Where-Object { $_.result -imatch 'Failed|Fail' }).Count
$warnCount   = @($sorted | Where-Object { $_.result -imatch 'Warning|Warn' }).Count
$successCount= @($sorted | Where-Object { $_.result -imatch 'Success' }).Count
$withError   = @($sorted | Where-Object { -not [string]::IsNullOrWhiteSpace($_.last_error) }).Count

# ---------------------------------------------------------------------------
# Output
# ---------------------------------------------------------------------------
if ($Json) {
    $sorted | ConvertTo-Json -Depth 6
    Write-Warning ('Summary: jobs={0} failed={1} warning={2} success={3} with_error={4}' `
        -f $totalJobs, $failedCount, $warnCount, $successCount, $withError)
} else {
    if ($sorted.Count -eq 0) {
        if ($OnlyFailures) {
            Write-Output 'No Failed or Warning sessions found in the specified window.'
        } else {
            Write-Output 'No sessions found in the specified window.'
        }
    } else {
        Write-TextReport -Reports $sorted
    }

    Write-Output '------------------------------------------------------------'
    Write-Output ('Window   : last {0} hour(s)  ({1:o}  to  {2:o})' -f $Hours, $script:StartTime, $script:EndTime)
    Write-Output ('Jobs     : {0}  (Failed: {1}  Warning: {2}  Success: {3}  WithError: {4})' `
        -f $totalJobs, $failedCount, $warnCount, $successCount, $withError)
    Write-Output '------------------------------------------------------------'
}
