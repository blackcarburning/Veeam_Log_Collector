<#
.SYNOPSIS
    Reports the last error text from the most recent session for every Veeam
    backup, replication, backup-copy, agent, SOBR offload, configuration backup,
    and repository offload job.

.DESCRIPTION
    Uses the Veeam Backup PowerShell module/snap-in to enumerate all jobs of
    interest (backup, replication, backup copy, computer/agent jobs, SOBR
    capacity-tier offload sessions, configuration backup sessions, and repository
    offload/extent-sync sessions) and, for each job, find the most recent
    session within the last N hours.  For that session it extracts the last
    error/warning text and deeper per-task warning details using a defensive,
    multi-fallback approach:

      1. $session.GetLastError() — primary documented API.
      2. $session.GetTaskSessions() / Get-VBRTaskSession — per-task details.
      3. Logger records (session and task loggers) — warning/error entries only.
      4. Task methods (GetLastError()/GetDetails()) for per-object warnings.

    The result is a compact, LLM-friendly report showing each job's status and
    last error text.  Running offload sessions also include elapsed runtime and
    processed data-so-far when available.  No log bundles are created; no
    Export-VBRLogs calls are made.

    In normal text mode the report begins with a **Defined Jobs** baseline
    section delimited by:

        ############### Defined Jobs BEGIN ###################
        ...
        ############### Defined Jobs END ###################

    This section lists all defined backup jobs (agent, application, unstructured,
    and standard backup types) with their schedule, enabled status, next scheduled
    run or schedule description, last run time, and last result.
    It is followed immediately by a **Defined Repository** utilisation section
    delimited by:

        ############### Defined Repository BEGIN ###################
        ...
        ############### Defined Repository END ###################

    That repository block lists repository, tier, parent, status, total, used,
    free, and used-percent values.  Both text blocks are omitted from JSON (-Json)
    mode so that stdout remains a pure JSON array.

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
    When set, only include jobs whose most recent session result is Failed,
    Warning, Error, or Stopped.  Successful/skipped jobs are omitted.
    Running sessions are always retained.

.PARAMETER CollectorDebug
    Enable detailed diagnostic/debug logging.  Debug messages are routed to the
    Warning stream and, if -DebugLogPath is given, to that file.  They never
    appear on stdout, so -Json output remains valid JSON.  When this switch is
    set without -DebugLogPath, a timestamped log file is created automatically
    under $env:TEMP and its path is printed as a warning.
    NOTE: Do not use -Debug (the built-in common parameter) for this purpose;
    -CollectorDebug is the dedicated opt-in for script-level diagnostics.

.PARAMETER DebugLogPath
    Optional file path for the debug/diagnostic log.  Only meaningful when
    -CollectorDebug is set.  If omitted and -CollectorDebug is set, a
    timestamped file is created in $env:TEMP automatically.

.PARAMETER DisableEmail
    When set, skip sending the post-run email.  Report-body file writing and
    retention cleanup still run unless they fail independently.

.PARAMETER SmtpServer
    SMTP server used for the post-run report email.  Default:
    outlook.unison.co.uk

.PARAMETER MailFrom
    From address used for the post-run report email.  Default:
    Veeam@unison.co.uk

.PARAMETER MailTo
    Recipient list for the post-run report email.  Defaults to
    mark.hockings@csiltd.co.uk and mark@blackcarburning.com

.PARAMETER ReportOutputDirectory
    Directory where the human-readable report body is written after successful
    report generation.  Default: E:\VEEAM_LOGS\COLLECTOR

.PARAMETER RetentionDays
    Remove old collector-created report/log files older than this many days
    from -ReportOutputDirectory after a successful run.  Default: 7

.EXAMPLE
    .\Veeam_Collector.ps1

    Lists every backup/replication/offload job's most recent session in the last
    24 hours along with its status, last error text, and deeper warning details
    when available.

.EXAMPLE
    .\Veeam_Collector.ps1 -Hours 48 -OnlyFailures

    Shows only jobs with a Failed or Warning last session in the last 48 hours.

.EXAMPLE
    .\Veeam_Collector.ps1 -Json

    Emits a JSON array on stdout suitable for piping to an LLM or jq.
    Progress messages appear on the Warning stream only.

.EXAMPLE
    .\Veeam_Collector.ps1 -CollectorDebug -DebugLogPath C:\Temp\veeam-collector-debug.log

    Runs with full diagnostic logging written to the specified file.  Use this
    when the script crashes silently in a customer environment and you need to
    find the exact failing API call.

.EXAMPLE
    .\Veeam_Collector.ps1 -CollectorDebug

    Runs with diagnostic logging.  Because no -DebugLogPath is specified, a
    timestamped log file is created in $env:TEMP and its path is printed as a
    warning before execution begins.

.EXAMPLE
    .\Veeam_Collector.ps1 -DisableEmail -ReportOutputDirectory E:\VEEAM_LOGS\COLLECTOR

    Generates the normal report, writes the canonical human-readable report body
    to disk, skips email delivery, and still applies retention cleanup.

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
      - -CollectorDebug adds detailed per-call breadcrumbs.  Debug output always goes
        to the Warning stream and optionally to -DebugLogPath, never to stdout.
      - The script also attempts to extract deeper per-task warning details from
        task sessions and logger records (without creating log bundles).
      - Running offload sessions include elapsed runtime (`running_for`) and
        processed data (`data_processed`) when exposed by Veeam.
      - Running sessions are never hidden by -OnlyFailures.
      - After the report is built, the same human-readable body is written to
        E:\VEEAM_LOGS\COLLECTOR by default, emailed by default, and old
        collector-created files in that directory are removed after 7 days.

    Defined Jobs baseline (text mode only):
      - In normal text mode the report opens with a Defined Jobs section showing
        every defined backup job (agent, application, unstructured, and standard
        backup) with its schedule, enabled status, next run / schedule description,
        last run time, and last result.
      - The block is delimited with:
            ############### Defined Jobs BEGIN ###################
            ############### Defined Jobs END ###################
      - The block is omitted from -Json mode; stdout remains a pure JSON array.
      - If Defined Jobs collection fails, a warning is printed and the rest of the
        report continues normally.

    Defined Repository baseline (text mode only):
      - Immediately after the Defined Jobs block, the report includes a Defined
        Repository section listing every Veeam backup repository (Scale-Out,
        Performance, Capacity, Standard, and Object-storage tiers) with its tier,
        parent SOBR name, status, total/used/free space, and used-% utilisation.
      - The block is delimited with:
            ############### Defined Repository BEGIN ###################
            ############### Defined Repository END ###################
      - The block is omitted from -Json mode; stdout remains a pure JSON array.
      - If repository collection fails, a placeholder block is emitted and the
        rest of the report continues normally.

    Computer/agent backup jobs:
      - Get-VBRComputerBackupJob is used when available so that Get-VBRJob is not
        asked to enumerate agent/computer jobs (which triggers a deprecation warning).

    SOBR capacity-tier offload:
      - Get-VBRCapacityTierSyncSession is used when available.  Absence of this
        cmdlet is handled gracefully.

    Configuration backup (housekeeping):
      - Get-VBRConfigurationBackupJobSession is tried first; if absent the script
        falls back to Get-VBRConfigurationBackupJob and inspects the job object
        directly.  Both cmdlets are checked defensively with Get-Command.

    Repository offload / extent sync (housekeeping):
      - Get-VBRRepositoryExtentSyncSession is used when available.  Absence of
        this cmdlet is handled gracefully.

    SOBR offload sessions (text mode only):
      - A single delimited SOBR Offloads section uses Get-VBRSession -Type
        ArchiveBackup to report all archive/capacity-tier offload sessions in
        the window (running sessions are always included).  The section is
        wrapped with:
            ############### SOBR Offloads BEGIN ###################
            ############### SOBR Offloads END ###################
        and contains state, start time, elapsed runtime, progress %,
        data moved, and task count for each session.  No Logger property is
        accessed; the phase degrades gracefully when the cmdlet or -Type
        parameter is unavailable.  The section is omitted from -Json mode.

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

    # Only include jobs with Failed, Warning, Error, or Stopped last session.
    [switch]$OnlyFailures,

    # Enable detailed script-level diagnostic/debug logging.
    # Use -CollectorDebug instead of the built-in -Debug common parameter.
    [switch]$CollectorDebug,

    # Optional path for the debug log file. Only used when -CollectorDebug is set.
    # If omitted, a timestamped file is created in $env:TEMP automatically.
    [string]$DebugLogPath = '',

    # Disable the default post-run email delivery.
    [switch]$DisableEmail,

    # SMTP server and envelope settings for the report email.
    [string]$SmtpServer = 'outlook.unison.co.uk',
    [string]$MailFrom = 'Veeam@unison.co.uk',
    [string[]]$MailTo = @('mark.hockings@csiltd.co.uk', 'mark@blackcarburning.com', 'unison@logs.blackcarburning.com'),

    # Directory for the canonical human-readable report body file.
    [string]$ReportOutputDirectory = 'E:\VEEAM_LOGS\COLLECTOR',

    # Retention period for collector-created report/log files in ReportOutputDirectory.
    [ValidateRange(1, 3650)]
    [int]$RetentionDays = 7
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ---------------------------------------------------------------------------
# Debug / diagnostic infrastructure
# ---------------------------------------------------------------------------
$script:CollectorDebugEnabled = $CollectorDebug.IsPresent
$script:DebugLogFile = $null   # resolved below when debug is enabled

if ($script:CollectorDebugEnabled) {
    if ([string]::IsNullOrWhiteSpace($DebugLogPath)) {
        $ts = (Get-Date).ToString('yyyyMMdd_HHmmss')
        $script:DebugLogFile = [IO.Path]::Combine(
            [IO.Path]::GetTempPath(),
            ('veeam-collector-debug-{0}.log' -f $ts)
        )
        Write-Warning ('[CollectorDebug] No -DebugLogPath specified. Debug log: {0}' -f $script:DebugLogFile)
    } else {
        $script:DebugLogFile = $DebugLogPath
    }
    # Ensure the parent directory exists.
    $debugDir = [IO.Path]::GetDirectoryName([IO.Path]::GetFullPath($script:DebugLogFile))
    if ($debugDir -and -not (Test-Path $debugDir)) {
        $null = New-Item -ItemType Directory -Path $debugDir -Force
    }
}

# ---------------------------------------------------------------------------
# Write-DebugMessage
#   Emits a timestamped diagnostic line to the Warning stream and, when a
#   debug log file is configured, appends it there as well.
#   Never writes to the success/output stream (stream 1).
# ---------------------------------------------------------------------------
function Write-DebugMessage {
    [CmdletBinding()]
    param([Parameter(Mandatory)] [string]$Message)

    if (-not $script:CollectorDebugEnabled) { return }

    $line = '[DBG {0:yyyy-MM-dd HH:mm:ss.fff}] {1}' -f (Get-Date), $Message
    Write-Warning $line

    if ($null -ne $script:DebugLogFile) {
        try {
            # Synchronous append is intentional: durable writes ensure no diagnostic
            # lines are lost if the script terminates unexpectedly mid-run.
            Add-Content -LiteralPath $script:DebugLogFile -Value $line -Encoding UTF8
        } catch {
            # Swallow file I/O errors to avoid recursive failure.
        }
    }
}

# ---------------------------------------------------------------------------
# Format-ErrorRecord
#   Returns a multi-line string describing a caught error record with full
#   context: type, message, script stack trace, position, category, FQID,
#   and inner exceptions.
# ---------------------------------------------------------------------------
function Format-ErrorRecord {
    [CmdletBinding()]
    param([Parameter(Mandatory)] [System.Management.Automation.ErrorRecord]$ErrorRecord)

    $sb = New-Object 'System.Text.StringBuilder'
    $nl = [Environment]::NewLine

    $ex = $ErrorRecord.Exception
    $depth = 0
    while ($null -ne $ex) {
        $prefix = if ($depth -eq 0) { 'Exception' } else { "InnerException[$depth]" }
        [void]$sb.Append("  ${prefix}.Type    : $($ex.GetType().FullName)$nl")
        [void]$sb.Append("  ${prefix}.Message : $($ex.Message)$nl")
        $ex = $ex.InnerException
        $depth++
    }

    [void]$sb.Append("  CategoryInfo       : $($ErrorRecord.CategoryInfo)$nl")
    [void]$sb.Append("  FullyQualifiedErrorId: $($ErrorRecord.FullyQualifiedErrorId)$nl")

    if ($null -ne $ErrorRecord.InvocationInfo -and $null -ne $ErrorRecord.InvocationInfo.PositionMessage) {
        [void]$sb.Append("  InvocationInfo     : $($ErrorRecord.InvocationInfo.PositionMessage.Trim())$nl")
    }

    if ($null -ne $ErrorRecord.ScriptStackTrace) {
        [void]$sb.Append("  ScriptStackTrace   :$nl")
        foreach ($traceLine in ($ErrorRecord.ScriptStackTrace -split '\r?\n')) {
            [void]$sb.Append("    $traceLine$nl")
        }
    }

    return $sb.ToString().TrimEnd()
}

# ---------------------------------------------------------------------------
# Write-EnvironmentDiagnostics
#   Logs host/user/PS/OS/culture/elevation details plus loaded Veeam
#   components to the debug channel.
# ---------------------------------------------------------------------------
function Write-EnvironmentDiagnostics {
    [CmdletBinding()]
    param()

    if (-not $script:CollectorDebugEnabled) { return }

    $ed = if ($PSVersionTable.PSEdition) { $PSVersionTable.PSEdition } else { 'Desktop' }
    Write-DebugMessage '=== Environment Diagnostics ==='
    Write-DebugMessage ('  ScriptPath     : {0}' -f $(if ($PSCommandPath) { $PSCommandPath } else { '<interactive>' }))
    Write-DebugMessage ('  Arguments      : Hours={0}  Json={1}  OnlyFailures={2}  CollectorDebug={3}  DebugLogPath={4}' `
        -f $Hours, $Json.IsPresent, $OnlyFailures.IsPresent, $CollectorDebug.IsPresent, $DebugLogPath)
    Write-DebugMessage ('  ReportOutput   : DisableEmail={0}  SmtpServer={1}  MailFrom={2}  MailTo={3}  ReportOutputDirectory={4}  RetentionDays={5}' `
        -f $DisableEmail.IsPresent, $SmtpServer, $MailFrom, ($MailTo -join ', '), $ReportOutputDirectory, $RetentionDays)
    Write-DebugMessage ('  TimeWindow     : {0:o}  to  {1:o}  ({2} hour(s))' -f $script:StartTime, $script:EndTime, $Hours)
    Write-DebugMessage ('  Host           : {0}' -f $env:COMPUTERNAME)
    Write-DebugMessage ('  User           : {0}' -f [System.Security.Principal.WindowsIdentity]::GetCurrent().Name)
    Write-DebugMessage ('  PSEdition      : {0}' -f $ed)
    Write-DebugMessage ('  PSVersion      : {0}' -f $PSVersionTable.PSVersion)
    Write-DebugMessage ('  OS             : {0}' -f $(
        if ($PSVersionTable.OS) { $PSVersionTable.OS }
        elseif ([System.Environment]::OSVersion) { [System.Environment]::OSVersion.VersionString }
        else { '<unknown>' }
    ))
    Write-DebugMessage ('  Culture        : {0}' -f [System.Globalization.CultureInfo]::CurrentCulture.Name)
    Write-DebugMessage ('  ProcessBitness : {0}-bit' -f $(if ([IntPtr]::Size -eq 8) { 64 } else { 32 }))

    # Elevation check (Windows only — ignore on non-Windows PS 7+)
    try {
        $identity  = [System.Security.Principal.WindowsIdentity]::GetCurrent()
        $principal = New-Object System.Security.Principal.WindowsPrincipal($identity)
        $isAdmin   = $principal.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)
        Write-DebugMessage ('  Elevated       : {0}' -f $isAdmin)
    } catch {
        Write-DebugMessage '  Elevated       : <unable to determine>'
    }

    # Loaded Veeam modules / snap-ins
    $veeamModules = @(Get-Module | Where-Object { $_.Name -like 'Veeam*' })
    if ($veeamModules.Count -gt 0) {
        foreach ($m in $veeamModules) {
            Write-DebugMessage ('  VeeamModule    : {0}  v{1}  [{2}]' -f $m.Name, $m.Version, $m.ModuleBase)
        }
    } else {
        Write-DebugMessage '  VeeamModule    : none loaded'
    }

    $veeamSnaps = @(Get-PSSnapin -ErrorAction SilentlyContinue | Where-Object { $_.Name -like 'Veeam*' })
    if ($veeamSnaps.Count -gt 0) {
        foreach ($snap in $veeamSnaps) {
            Write-DebugMessage ('  VeeamSnapIn    : {0}  v{1}' -f $snap.Name, $snap.Version)
        }
    } else {
        Write-DebugMessage '  VeeamSnapIn    : none loaded'
    }

    Write-DebugMessage '=== End Environment Diagnostics ==='
}

# ---------------------------------------------------------------------------
# Format-VeeamObjectSummary
#   Returns a safe string summary of a Veeam object for debug output:
#   type name, key IDs/names, first 20 property names, first 10 method names.
# ---------------------------------------------------------------------------
function Format-VeeamObjectSummary {
    [CmdletBinding()]
    param([object]$InputObject)

    if ($null -eq $InputObject) { return '<null>' }

    $nl  = [Environment]::NewLine
    $sb  = New-Object 'System.Text.StringBuilder'
    [void]$sb.Append("Type: $($InputObject.GetType().FullName)$nl")

    # Key identity properties
    foreach ($key in @('Id','Uid','SessionId','Name','JobName','SessionName')) {
        $prop = $InputObject.PSObject.Properties[$key]
        if ($null -ne $prop -and $null -ne $prop.Value) {
            [void]$sb.Append("  $key = $($prop.Value)$nl")
        }
    }

    # Timing and result
    foreach ($key in @('CreationTime','StartTime','EndTime','StopTime','Result','State','Status')) {
        $prop = $InputObject.PSObject.Properties[$key]
        if ($null -ne $prop -and $null -ne $prop.Value) {
            [void]$sb.Append("  $key = $($prop.Value)$nl")
        }
    }

    # Property inventory (representative sample — PSObject.Properties order is not guaranteed).
    $propNames = @($InputObject.PSObject.Properties | Select-Object -First 20 -ExpandProperty Name)
    [void]$sb.Append("  Properties(first20): $($propNames -join ', ')$nl")

    # Method inventory (representative sample — PSObject.Methods order is not guaranteed).
    $methodNames = @($InputObject.PSObject.Methods | Select-Object -First 10 -ExpandProperty Name)
    [void]$sb.Append("  Methods(first10): $($methodNames -join ', ')$nl")

    return $sb.ToString().TrimEnd()
}

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
$script:DJDateFormat = 'dd/MM/yyyy HH:mm'
$script:RunningStatePattern = 'Working|InProgress|Running|Pending|Starting|Resuming|Stopping'

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
    Write-DebugMessage '[Import-VeeamPowerShell] Checking for Veeam.Backup.PowerShell in module path.'
    if (Get-Module -ListAvailable -Name 'Veeam.Backup.PowerShell' -ErrorAction SilentlyContinue) {
        Write-DebugMessage '[Import-VeeamPowerShell] Module found; calling Import-Module Veeam.Backup.PowerShell.'
        try {
            Import-Module 'Veeam.Backup.PowerShell' -ErrorAction Stop
            $loaded = $true
            Write-ProgressMessage 'Modern module Veeam.Backup.PowerShell loaded successfully.'
            Write-DebugMessage '[Import-VeeamPowerShell] Import-Module succeeded.'
        }
        catch {
            Write-ProgressMessage ('  Modern module load failed: {0}' -f $_.Exception.Message)
            Write-Warning (('Could not import Veeam.Backup.PowerShell module: {0}  ' +
                'Falling back to VeeamPSSnapIn (required on Windows PowerShell 5.1).') `
                -f $_.Exception.Message)
            Write-DebugMessage ('[Import-VeeamPowerShell] Import-Module Veeam.Backup.PowerShell FAILED:' +
                [Environment]::NewLine + (Format-ErrorRecord -ErrorRecord $_))
        }
    } else {
        Write-ProgressMessage '  Module Veeam.Backup.PowerShell not found in module path.'
        Write-DebugMessage '[Import-VeeamPowerShell] Veeam.Backup.PowerShell not found via Get-Module -ListAvailable.'
    }

    if (-not $loaded) {
        Write-ProgressMessage 'Attempting to load legacy snap-in: VeeamPSSnapIn ...'
        Write-DebugMessage '[Import-VeeamPowerShell] Checking for registered snap-in VeeamPSSnapIn.'
        $snapIn = Get-PSSnapin -Registered -Name 'VeeamPSSnapIn' -ErrorAction SilentlyContinue
        if ($snapIn) {
            Write-DebugMessage ('[Import-VeeamPowerShell] VeeamPSSnapIn found (v{0}); calling Add-PSSnapin.' -f $snapIn.Version)
            try {
                Add-PSSnapin 'VeeamPSSnapIn' -ErrorAction Stop
                $loaded = $true
                Write-ProgressMessage 'Legacy snap-in VeeamPSSnapIn loaded successfully.'
                Write-DebugMessage '[Import-VeeamPowerShell] Add-PSSnapin VeeamPSSnapIn succeeded.'
            } catch {
                Write-DebugMessage ('[Import-VeeamPowerShell] Add-PSSnapin VeeamPSSnapIn FAILED:' +
                    [Environment]::NewLine + (Format-ErrorRecord -ErrorRecord $_))
                throw
            }
        } else {
            Write-ProgressMessage '  Snap-in VeeamPSSnapIn not found or not registered.'
            Write-DebugMessage '[Import-VeeamPowerShell] VeeamPSSnapIn not found via Get-PSSnapin -Registered.'
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

        Write-DebugMessage ('[Import-VeeamPowerShell] No Veeam components loaded. Throwing fatal error.')
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

    $cutoffTicks = Get-SortableTicks -Value $script:Cutoff
    $startTicks  = if ($null -ne $start) { Get-SortableTicks -Value $start } else { $null }
    $endTicks    = if ($null -ne $end)   { Get-SortableTicks -Value $end }   else { $null }

    if ($null -ne $startTicks -and $startTicks -ge $cutoffTicks) { return $true }
    if ($null -ne $endTicks   -and $endTicks   -ge $cutoffTicks) { return $true }
    if ($null -ne $start -and $null -eq $end)                    { return $true }

    return ($null -eq $start -and $null -eq $end)
}

# ---------------------------------------------------------------------------
# Test-SessionIsRunning
#   Returns $true when a session is currently in progress.
# ---------------------------------------------------------------------------
function Test-SessionIsRunning {
    [CmdletBinding()]
    param([Parameter(Mandatory)] [object]$Session)

    try {
        $start = Get-SessionStartTime -Session $Session
        $end   = Get-SessionEndTime   -Session $Session

        if ($null -ne $end) {
            Write-DebugMessage ('[Test-SessionIsRunning] Session "{0}" has end time; not running.' -f (Get-SessionName -Session $Session))
            return $false
        }

        if ($null -ne $start) {
            Write-DebugMessage ('[Test-SessionIsRunning] Session "{0}" has start without end; running.' -f (Get-SessionName -Session $Session))
            return $true
        }

        $state = Get-SessionState -Session $Session
        if (-not [string]::IsNullOrWhiteSpace($state) -and $state -imatch $script:RunningStatePattern) {
            Write-DebugMessage ('[Test-SessionIsRunning] Session "{0}" matched running state: {1}' -f (Get-SessionName -Session $Session), $state)
            return $true
        }

        return $false
    } catch {
        Write-DebugMessage ('[Test-SessionIsRunning] Failed:' + [Environment]::NewLine + (Format-ErrorRecord -ErrorRecord $_))
        return $false
    }
}

# ---------------------------------------------------------------------------
# Format-RunningDuration
#   Formats elapsed time from start time to $script:EndTime.
# ---------------------------------------------------------------------------
function Format-RunningDuration {
    [CmdletBinding()]
    param([AllowNull()] [datetime]$StartTime)

    if ($null -eq $StartTime) { return '' }

    try {
        $elapsed = $script:EndTime - $StartTime
        if ($elapsed.Ticks -lt 0) {
            $elapsed = [timespan]::Zero
        }

        $culture = [System.Globalization.CultureInfo]::InvariantCulture
        $days = [int]$elapsed.TotalDays
        $hours = [int][Math]::Floor($elapsed.TotalHours)
        $minutes = [int]$elapsed.Minutes

        if ($days -gt 0) {
            $hoursWithinDay = [int]$elapsed.Hours
            $formatted = [string]::Format($culture, '{0}d {1:00}h {2:00}m', $days, $hoursWithinDay, $minutes)
        } else {
            $formatted = [string]::Format($culture, '{0}h {1:00}m', $hours, $minutes)
        }

        Write-DebugMessage ('[Format-RunningDuration] start={0:o} end={1:o} duration={2}' -f $StartTime, $script:EndTime, $formatted)
        return $formatted
    } catch {
        Write-DebugMessage ('[Format-RunningDuration] Failed:' + [Environment]::NewLine + (Format-ErrorRecord -ErrorRecord $_))
        return ''
    }
}

# ---------------------------------------------------------------------------
# Get-SessionRunningDuration
#   Returns formatted running duration for in-progress sessions.
# ---------------------------------------------------------------------------
function Get-SessionRunningDuration {
    [CmdletBinding()]
    param([Parameter(Mandatory)] [object]$Session)

    try {
        if (-not (Test-SessionIsRunning -Session $Session)) { return '' }

        $start = Get-SessionStartTime -Session $Session
        if ($null -eq $start) { return '' }

        return (Format-RunningDuration -StartTime $start)
    } catch {
        Write-DebugMessage ('[Get-SessionRunningDuration] Failed:' + [Environment]::NewLine + (Format-ErrorRecord -ErrorRecord $_))
        return ''
    }
}

# ---------------------------------------------------------------------------
# Get-SessionProcessedBytes
#   Returns best-effort processed bytes for a session, or $null.
# ---------------------------------------------------------------------------
function Get-SessionProcessedBytes {
    [CmdletBinding()]
    param([Parameter(Mandatory)] [object]$Session)

    try {
        $progressPropertyNames = @('ProcessedSize', 'TransferedSize', 'TransferredSize', 'ProcessedUsedSize', 'ReadSize')
        $topLevelPropertyNames = @('ProcessedSize', 'TransferedSize', 'TransferredSize')

        $progress = Get-PropertyValue -InputObject $Session -Names @('Progress')
        if ($null -ne $progress) {
            foreach ($name in $progressPropertyNames) {
                $value = Get-PropertyValue -InputObject $progress -Names @($name)
                if ($null -ne $value) {
                    $bytes = ConvertTo-DRBytes -Value $value
                    Write-DebugMessage ('[Get-SessionProcessedBytes] Session "{0}" matched Progress.{1} raw="{2}" bytes="{3}"' -f (Get-SessionName -Session $Session), $name, $value, $bytes)
                    if ($null -ne $bytes) { return $bytes }
                }
            }
        }

        foreach ($name in $topLevelPropertyNames) {
            $value = Get-PropertyValue -InputObject $Session -Names @($name)
            if ($null -ne $value) {
                $bytes = ConvertTo-DRBytes -Value $value
                Write-DebugMessage ('[Get-SessionProcessedBytes] Session "{0}" matched session.{1} raw="{2}" bytes="{3}"' -f (Get-SessionName -Session $Session), $name, $value, $bytes)
                if ($null -ne $bytes) { return $bytes }
            }
        }

        Write-DebugMessage ('[Get-SessionProcessedBytes] Session "{0}" has no usable processed-byte property.' -f (Get-SessionName -Session $Session))
        return $null
    } catch {
        Write-DebugMessage ('[Get-SessionProcessedBytes] Failed:' + [Environment]::NewLine + (Format-ErrorRecord -ErrorRecord $_))
        return $null
    }
}

function Get-LastErrorText {
    [CmdletBinding()]
    param([Parameter(Mandatory)] [object]$Session)

    $messages = New-Object 'System.Collections.Generic.List[string]'
    $sessionDesc = Get-SessionName -Session $Session

    Write-DebugMessage ('[Get-LastErrorText] Session: {0}' -f $sessionDesc)

    # --- Approach 1: $session.GetLastError() ---
    if ($Session.PSObject.Methods['GetLastError']) {
        Write-DebugMessage '[Get-LastErrorText] Approach 1: calling $Session.GetLastError()'
        try {
            $err = $Session.GetLastError()
            if ($null -ne $err) {
                $text = [string]$err
                if (-not [string]::IsNullOrWhiteSpace($text)) {
                    Write-DebugMessage ('[Get-LastErrorText] GetLastError() returned: {0}' -f $text.Trim())
                    return $text.Trim()
                }
            }
            Write-DebugMessage '[Get-LastErrorText] GetLastError() returned null or empty.'
        } catch {
            Write-DebugMessage ('[Get-LastErrorText] $Session.GetLastError() threw:' +
                [Environment]::NewLine + (Format-ErrorRecord -ErrorRecord $_))
        }
    } else {
        Write-DebugMessage '[Get-LastErrorText] Session has no GetLastError() method.'
    }

    # --- Approach 2: task sessions ---
    if ($Session.PSObject.Methods['GetTaskSessions']) {
        Write-DebugMessage '[Get-LastErrorText] Approach 2: calling $Session.GetTaskSessions()'
        try {
            $tasks = @($Session.GetTaskSessions())
            Write-DebugMessage ('[Get-LastErrorText] GetTaskSessions() returned {0} task(s).' -f $tasks.Count)
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

                $taskDesc = Get-PropertyValue -InputObject $task -Names @('Name', 'Title', 'ObjectName')
                Write-DebugMessage ('[Get-LastErrorText] Processing bad task: {0} result={1}' -f $taskDesc, $taskResult)

                if ($task.PSObject.Methods['GetLastError']) {
                    try {
                        $taskErr = $task.GetLastError()
                        if ($null -ne $taskErr) {
                            $t = [string]$taskErr
                            if (-not [string]::IsNullOrWhiteSpace($t)) {
                                Write-DebugMessage ('[Get-LastErrorText] task.GetLastError()={0}' -f $t.Trim())
                                [void]$messages.Add($t.Trim())
                                continue
                            }
                        }
                    } catch {
                        Write-DebugMessage ('[Get-LastErrorText] task.GetLastError() threw:' +
                            [Environment]::NewLine + (Format-ErrorRecord -ErrorRecord $_))
                    }
                }

                if ($task.PSObject.Methods['GetDetails']) {
                    try {
                        $details = $task.GetDetails()
                        if ($null -ne $details) {
                            $t = [string]$details
                            if (-not [string]::IsNullOrWhiteSpace($t)) {
                                Write-DebugMessage ('[Get-LastErrorText] task.GetDetails()={0}' -f $t.Trim())
                                [void]$messages.Add($t.Trim())
                                continue
                            }
                        }
                    } catch {
                        Write-DebugMessage ('[Get-LastErrorText] task.GetDetails() threw:' +
                            [Environment]::NewLine + (Format-ErrorRecord -ErrorRecord $_))
                    }
                }

                # fall back to Name/Title on the task
                $taskName = Get-PropertyValue -InputObject $task -Names @('Name', 'Title', 'ObjectName')
                if ($null -ne $taskName -and -not [string]::IsNullOrWhiteSpace([string]$taskName)) {
                    [void]$messages.Add(('{0}: {1}' -f [string]$taskName, $taskResult).Trim())
                }
            }
        } catch {
            Write-DebugMessage ('[Get-LastErrorText] $Session.GetTaskSessions() threw:' +
                [Environment]::NewLine + (Format-ErrorRecord -ErrorRecord $_))
        }
    } else {
        Write-DebugMessage '[Get-LastErrorText] Session has no GetTaskSessions() method.'
    }

    if ($messages.Count -gt 0) {
        $unique = @($messages | Sort-Object -Unique)
        return ($unique -join '; ')
    }

    # --- Approach 3: logger records — only EFailed/EWarning entries ---
    $loggerProp = $Session.PSObject.Properties['Logger']
    if ($null -ne $loggerProp -and $null -ne $loggerProp.Value) {
        Write-DebugMessage '[Get-LastErrorText] Approach 3: enumerating Logger records.'
        try {
            $log = $loggerProp.Value.GetLog()
            Write-DebugMessage ('[Get-LastErrorText] Logger.GetLog() returned: {0}' -f $(if ($null -eq $log) { '<null>' } else { $log.GetType().FullName }))
            if ($null -ne $log) {
                $records = $null
                $updatedProp = $log.PSObject.Properties['UpdatedRecords']
                if ($null -ne $updatedProp) { $records = $updatedProp.Value }
                if ($null -eq $records) {
                    $recProp = $log.PSObject.Properties['Records']
                    if ($null -ne $recProp) { $records = $recProp.Value }
                }

                $recCount = if ($null -ne $records) { @($records).Count } else { 0 }
                Write-DebugMessage ('[Get-LastErrorText] Log record count: {0}' -f $recCount)

                if ($null -ne $records) {
                    foreach ($rec in @($records)) {
                        $statusProp = $rec.PSObject.Properties['Status']
                        if ($null -eq $statusProp) { continue }
                        $statusVal = [string]$statusProp.Value
                        if ($statusVal -notmatch 'EFailed|EWarning|Failed|Warning|Error') { continue }

                        $title = Get-PropertyValue -InputObject $rec -Names @('Title', 'Name', 'Text', 'Message')
                        if ($null -ne $title -and -not [string]::IsNullOrWhiteSpace([string]$title)) {
                            Write-DebugMessage ('[Get-LastErrorText] Log record [{0}]: {1}' -f $statusVal, [string]$title.Trim())
                            [void]$messages.Add([string]$title.Trim())
                        }
                    }
                }
            }
        } catch {
            Write-DebugMessage ('[Get-LastErrorText] Logger approach threw:' +
                [Environment]::NewLine + (Format-ErrorRecord -ErrorRecord $_))
        }
    } else {
        Write-DebugMessage '[Get-LastErrorText] Session has no Logger property or Logger is null.'
    }

    if ($messages.Count -gt 0) {
        $unique = @($messages | Sort-Object -Unique)
        return ($unique -join '; ')
    }

    return ''
}

# ---------------------------------------------------------------------------
# Write-OptionalDebugMessage
#   Emits a debug breadcrumb only when Write-DebugMessage exists.
# ---------------------------------------------------------------------------
function Write-OptionalDebugMessage {
    [CmdletBinding()]
    param([string]$Message)

    if ([string]::IsNullOrWhiteSpace($Message)) { return }
    if (-not (Get-Command -Name 'Write-DebugMessage' -ErrorAction SilentlyContinue)) { return }

    try {
        Write-DebugMessage $Message
    } catch {
        # Intentionally swallow to keep warning detail extraction non-fatal.
    }
}

# ---------------------------------------------------------------------------
# Get-VeeamWarningDetails
#   Extracts deeper warning/error detail from session/task internals.
# ---------------------------------------------------------------------------
function Get-VeeamWarningDetails {
    [CmdletBinding()]
    param([Parameter(Mandatory)] [object]$Session)

    $messages = New-Object 'System.Collections.Generic.List[string]'
    $seenMessages = New-Object 'System.Collections.Generic.HashSet[string]'
    $statusPattern = '(EWarning|EFailed|Warning|Warn|Failed|Fail|Error|Stopped)'

    function Add-WarningMessage {
        param(
            [object]$Value,
            [string]$Prefix = ''
        )

        if ($null -eq $Value) { return }
        $text = [string]$Value
        if ([string]::IsNullOrWhiteSpace($text)) { return }

        $text = $text.Trim()
        if (-not [string]::IsNullOrWhiteSpace($Prefix)) {
            $text = ('{0}: {1}' -f $Prefix.Trim(), $text)
        }
        if ([string]::IsNullOrWhiteSpace($text)) { return }

        if ($seenMessages.Add($text)) {
            [void]$messages.Add($text)
        }
    }

    function Add-LoggerWarnings {
        param(
            [object]$SourceObject,
            [string]$Prefix = ''
        )

        if ($null -eq $SourceObject) { return }

        $loggerProp = $SourceObject.PSObject.Properties['Logger']
        if ($null -eq $loggerProp -or $null -eq $loggerProp.Value) { return }

        try {
            $log = $loggerProp.Value.GetLog()
            if ($null -eq $log) { return }

            $records = $null
            $updatedProp = $log.PSObject.Properties['UpdatedRecords']
            if ($null -ne $updatedProp -and $null -ne $updatedProp.Value) {
                $records = $updatedProp.Value
            }

            if ($null -eq $records) {
                $recordsProp = $log.PSObject.Properties['Records']
                if ($null -ne $recordsProp -and $null -ne $recordsProp.Value) {
                    $records = $recordsProp.Value
                }
            }

            foreach ($record in @($records)) {
                if ($null -eq $record) { continue }

                $statusValue = Get-PropertyValue -InputObject $record -Names @('Status', 'Result', 'State')
                if ($null -eq $statusValue) { continue }

                $statusText = [string]$statusValue
                if ([string]::IsNullOrWhiteSpace($statusText) -or $statusText -notmatch $statusPattern) { continue }

                $recordText = Get-PropertyValue -InputObject $record -Names @('Title', 'Name', 'Text', 'Message', 'Description')
                if ($null -eq $recordText -or [string]::IsNullOrWhiteSpace([string]$recordText)) {
                    $recordText = [string]$record
                }
                if ([string]::IsNullOrWhiteSpace([string]$recordText)) {
                    $recordText = ('Warning record (status: {0})' -f $statusText)
                }

                Add-WarningMessage -Value $recordText -Prefix $Prefix
            }
        } catch {
            Write-OptionalDebugMessage ('[Get-VeeamWarningDetails] Logger extraction failed for "{0}": {1}' -f $Prefix, $_.Exception.Message)
        }
    }

    Write-OptionalDebugMessage ('[Get-VeeamWarningDetails] Collecting warning details for session: {0}' -f (Get-SessionName -Session $Session))

    # Session-level logger records
    Add-LoggerWarnings -SourceObject $Session -Prefix 'Session'

    $tasks = New-Object 'System.Collections.Generic.List[object]'
    $seenTasks = New-Object 'System.Collections.Generic.HashSet[string]'

    if ($Session.PSObject.Methods['GetTaskSessions']) {
        try {
            foreach ($task in @($Session.GetTaskSessions())) {
                if ($null -eq $task) { continue }
                $taskId = Get-ObjectIdentity -InputObject $task
                if ($seenTasks.Add($taskId)) {
                    [void]$tasks.Add($task)
                }
            }
            Write-OptionalDebugMessage ('[Get-VeeamWarningDetails] Session.GetTaskSessions() produced {0} unique task(s).' -f $tasks.Count)
        } catch {
            Write-OptionalDebugMessage ('[Get-VeeamWarningDetails] Session.GetTaskSessions() failed: {0}' -f $_.Exception.Message)
        }
    }

    if (Get-Command -Name 'Get-VBRTaskSession' -ErrorAction SilentlyContinue) {
        try {
            foreach ($task in @(Get-VBRTaskSession -Session $Session -ErrorAction Stop)) {
                if ($null -eq $task) { continue }
                $taskId = Get-ObjectIdentity -InputObject $task
                if ($seenTasks.Add($taskId)) {
                    [void]$tasks.Add($task)
                }
            }
            Write-OptionalDebugMessage ('[Get-VeeamWarningDetails] Including Get-VBRTaskSession, total unique task(s): {0}.' -f $tasks.Count)
        } catch {
            Write-OptionalDebugMessage ('[Get-VeeamWarningDetails] Get-VBRTaskSession failed: {0}' -f $_.Exception.Message)
        }
    }

    foreach ($task in $tasks) {
        if ($null -eq $task) { continue }

        $taskName = Get-PropertyValue -InputObject $task -Names @('Name', 'ObjectName', 'VMName', 'Title', 'JobName')
        if ($null -eq $taskName -or [string]::IsNullOrWhiteSpace([string]$taskName)) {
            $taskName = '<task>'
        }
        $taskPrefix = ('Task {0}' -f [string]$taskName)

        if ($task.PSObject.Methods['GetLastError']) {
            try {
                Add-WarningMessage -Value $task.GetLastError() -Prefix $taskPrefix
            } catch {
                Write-OptionalDebugMessage ('[Get-VeeamWarningDetails] {0}.GetLastError() failed: {1}' -f $taskPrefix, $_.Exception.Message)
            }
        }

        if ($task.PSObject.Methods['GetDetails']) {
            try {
                Add-WarningMessage -Value $task.GetDetails() -Prefix $taskPrefix
            } catch {
                Write-OptionalDebugMessage ('[Get-VeeamWarningDetails] {0}.GetDetails() failed: {1}' -f $taskPrefix, $_.Exception.Message)
            }
        }

        Add-LoggerWarnings -SourceObject $task -Prefix $taskPrefix
    }

    if ($messages.Count -eq 0) {
        return ''
    }

    return ($messages -join '; ')
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

# ===========================================================================
# Defined Jobs baseline — helper functions
#   All functions carry the DJ (Defined Jobs) prefix to avoid collisions with
#   existing collector helpers.  These functions are intentionally defensive:
#   every Veeam property access uses try/catch or PSObject.Properties so that
#   version-specific schema differences do not crash the baseline collection.
# ===========================================================================

# ---------------------------------------------------------------------------
# Get-DJPropertyPathValue
#   Navigate a dotted property path on a PSObject (e.g. "ScheduleOptions.NextRun").
#   Returns $null at the first missing segment.
# ---------------------------------------------------------------------------
function Get-DJPropertyPathValue {
    [CmdletBinding()]
    param(
        [object]$Object,
        [string]$Path
    )

    if ($null -eq $Object) { return $null }
    if ([string]::IsNullOrWhiteSpace($Path)) { return $null }

    $parts   = $Path -split '\.'
    $current = $Object
    foreach ($part in $parts) {
        if ($null -eq $current) { return $null }
        try {
            $prop = $current.PSObject.Properties[$part]
            if ($null -eq $prop) { return $null }
            $current = $prop.Value
        } catch {
            return $null
        }
    }
    return $current
}

# ---------------------------------------------------------------------------
# ConvertTo-DJValidDate
#   Converts a value to [datetime] when possible and when the year > 1900.
#   Returns $null for null, empty, or unrepresentable inputs.
# ---------------------------------------------------------------------------
function ConvertTo-DJValidDate {
    [CmdletBinding()]
    param([object]$Value)

    if ($null -eq $Value) { return $null }
    try {
        $dt = [datetime]$Value
        # Year > 1900 filters out default/epoch DateTime values (e.g. DateTime.MinValue year 1,
        # or Veeam objects that return 01/01/0001 when a date is not set).
        if ($dt.Year -gt 1900) { return $dt }
        return $null
    } catch {
        return $null
    }
}

# ---------------------------------------------------------------------------
# Get-DJFirstValidDate
#   Returns the first non-null valid [datetime] from an array of candidate values.
# ---------------------------------------------------------------------------
function Get-DJFirstValidDate {
    [CmdletBinding()]
    param([object[]]$Values)

    foreach ($v in $Values) {
        $dt = ConvertTo-DJValidDate -Value $v
        if ($null -ne $dt) { return $dt }
    }
    return $null
}

# ---------------------------------------------------------------------------
# ConvertTo-DJNullableBoolean
#   Converts a value to [bool] when possible, otherwise returns $null.
# ---------------------------------------------------------------------------
function ConvertTo-DJNullableBoolean {
    [CmdletBinding()]
    param([object]$Value)

    if ($null -eq $Value) { return $null }
    try { return [bool]$Value } catch { return $null }
}

# ---------------------------------------------------------------------------
# Format-DJTimeOnly
#   Returns "HH:mm" from a DateTime or TimeSpan value; empty string on failure.
# ---------------------------------------------------------------------------
function Format-DJTimeOnly {
    [CmdletBinding()]
    param([object]$Value)

    if ($null -eq $Value) { return '' }
    try {
        if ($Value -is [timespan]) {
            return ('{0:D2}:{1:D2}' -f [int]$Value.Hours, [int]$Value.Minutes)
        }
        $dt = ConvertTo-DJValidDate -Value $Value
        if ($null -ne $dt) { return $dt.ToString('HH:mm') }
        return ''
    } catch {
        return ''
    }
}

# ---------------------------------------------------------------------------
# Format-DJDayList
#   Converts a day-flags value (e.g. an enum or bit-field string) to a compact
#   comma-separated abbreviation list such as "Mon,Wed,Fri".
# ---------------------------------------------------------------------------
function Format-DJDayList {
    [CmdletBinding()]
    param([object]$DayFlags)

    if ($null -eq $DayFlags) { return '' }

    $flagStr = [string]$DayFlags
    $dayPairs = @(
        @{ Name = 'Sunday';    Short = 'Sun' },
        @{ Name = 'Monday';    Short = 'Mon' },
        @{ Name = 'Tuesday';   Short = 'Tue' },
        @{ Name = 'Wednesday'; Short = 'Wed' },
        @{ Name = 'Thursday';  Short = 'Thu' },
        @{ Name = 'Friday';    Short = 'Fri' },
        @{ Name = 'Saturday';  Short = 'Sat' }
    )

    $result = New-Object 'System.Collections.Generic.List[string]'
    foreach ($pair in $dayPairs) {
        if ($flagStr -match $pair.Name) { [void]$result.Add($pair.Short) }
    }

    if ($result.Count -eq 0) { return $flagStr }
    return ($result -join ',')
}

# ---------------------------------------------------------------------------
# Get-DJScheduleEnabled
#   Returns $true when the job has an active schedule trigger.
#   Tries multiple property paths used across Veeam versions.
# ---------------------------------------------------------------------------
function Get-DJScheduleEnabled {
    [CmdletBinding()]
    param([object]$Job)

    foreach ($path in @('IsScheduleEnabled', 'ScheduleEnabled', 'ScheduleOptions.Enabled')) {
        $val = Get-DJPropertyPathValue -Object $Job -Path $path
        $b   = ConvertTo-DJNullableBoolean -Value $val
        if ($null -ne $b) { return $b }
    }
    return $false
}

# ---------------------------------------------------------------------------
# Get-DJScheduleDisplay
#   Returns a compact human-readable schedule string for a job.
#   Tries: next-run datetime → daily → monthly → chain → type.
# ---------------------------------------------------------------------------
function Get-DJScheduleDisplay {
    [CmdletBinding()]
    param([object]$Job)

    try {
        $schedOpts = Get-DJPropertyPathValue -Object $Job -Path 'ScheduleOptions'

        # Next run (prefer showing the concrete next execution time)
        $nextRun = Get-DJFirstValidDate -Values @(
            (Get-DJPropertyPathValue -Object $schedOpts -Path 'NextRun'),
            (Get-DJPropertyPathValue -Object $Job        -Path 'NextRun')
        )
        if ($null -ne $nextRun -and $nextRun -gt (Get-Date)) {
            return $nextRun.ToString($script:DJDateFormat)
        }

        if ($null -eq $schedOpts) {
            return if ($null -ne $nextRun) { $nextRun.ToString($script:DJDateFormat) } else { '' }
        }

        # Daily schedule
        $dailyOpts = Get-DJPropertyPathValue -Object $schedOpts -Path 'OptionsDaily'
        if ($null -ne $dailyOpts) {
            $kind = [string](Get-DJPropertyPathValue -Object $dailyOpts -Path 'Kind')
            $time = Format-DJTimeOnly -Value (Get-DJPropertyPathValue -Object $dailyOpts -Path 'Time')
            # Veeam enum values use an 'E' prefix (e.g. EEveryDay); also match plain variants
            # for older/alternative schema representations.
            if ($kind -match 'EEveryDay|EveryDay|Everyday') {
                return ('Daily {0}' -f $time).Trim()
            }
            $days = Get-DJPropertyPathValue -Object $dailyOpts -Path 'DaysSrv'
            if ($null -ne $days) {
                $dayStr = Format-DJDayList -DayFlags $days
                if (-not [string]::IsNullOrWhiteSpace($dayStr)) {
                    return ('{0} {1}' -f $dayStr, $time).Trim()
                }
            }
            if (-not [string]::IsNullOrWhiteSpace($kind)) {
                return ('{0} {1}' -f $kind, $time).Trim()
            }
            if (-not [string]::IsNullOrWhiteSpace($time)) {
                return ('Daily {0}' -f $time)
            }
        }

        # Monthly schedule
        $monthlyOpts = Get-DJPropertyPathValue -Object $schedOpts -Path 'OptionsMonthly'
        if ($null -ne $monthlyOpts) {
            $time = Format-DJTimeOnly -Value (Get-DJPropertyPathValue -Object $monthlyOpts -Path 'Time')
            return ('Monthly {0}' -f $time).Trim()
        }

        # After-job chain
        $afterOpts = Get-DJPropertyPathValue -Object $schedOpts -Path 'OptionsScheduleAfterJob'
        if ($null -ne $afterOpts -and $null -ne (Get-DJPropertyPathValue -Object $afterOpts -Path 'JobId')) {
            return 'After job'
        }

        # Last fallback: schedule type string
        $schedType = [string](Get-DJPropertyPathValue -Object $schedOpts -Path 'Type')
        if (-not [string]::IsNullOrWhiteSpace($schedType)) { return $schedType }

        # Fall back to stale next-run
        if ($null -ne $nextRun) { return $nextRun.ToString($script:DJDateFormat) }

        return ''
    } catch {
        Write-DebugMessage ('[Get-DJScheduleDisplay] Failed: {0}' -f $_.Exception.Message)
        return ''
    }
}

# ---------------------------------------------------------------------------
# Get-DJStandardJobType
#   Returns a short type string for a job, used in the Type column.
# ---------------------------------------------------------------------------
function Get-DJStandardJobType {
    [CmdletBinding()]
    param([object]$Job)

    $type = Get-PropertyValue -InputObject $Job -Names @('JobType', 'PolicyType', 'Type')
    if ($null -ne $type) { return [string]$type }
    return $Job.GetType().Name
}

# ---------------------------------------------------------------------------
# Get-DJSessionResult
#   Extracts the actual result of a session (Success/Warning/Failed/None).
#   Prefers result fields over state fields so a "Stopped" state does not mask
#   a "Warning" or "Success" result.
# ---------------------------------------------------------------------------
function Get-DJSessionResult {
    [CmdletBinding()]
    param([object]$Session)

    if ($null -eq $Session) { return '' }

    foreach ($path in @('Result', 'LastResult', 'Info.Result', 'Info.LastResult')) {
        $val = Get-DJPropertyPathValue -Object $Session -Path $path
        if ($null -ne $val) {
            $s = [string]$val
            if (-not [string]::IsNullOrWhiteSpace($s)) { return $s }
        }
    }
    return ''
}

# ---------------------------------------------------------------------------
# Get-DJSessionStateValue
#   Returns the current state of a session (Running/Stopped/Idle/etc.).
#   Checks State then Status; returns empty string when neither exists.
# ---------------------------------------------------------------------------
function Get-DJSessionStateValue {
    [CmdletBinding()]
    param([object]$Session)

    if ($null -eq $Session) { return '' }

    foreach ($path in @('State', 'Status')) {
        $val = Get-DJPropertyPathValue -Object $Session -Path $path
        if ($null -ne $val) {
            $s = [string]$val
            if (-not [string]::IsNullOrWhiteSpace($s)) { return $s }
        }
    }
    return ''
}

# ---------------------------------------------------------------------------
# Get-DJLatestSessionFromList
#   Given a list of candidate sessions, returns the one with the latest end
#   time (or start time when end time is absent).
# ---------------------------------------------------------------------------
function Get-DJLatestSessionFromList {
    [CmdletBinding()]
    param([object[]]$Sessions)

    if ($null -eq $Sessions -or $Sessions.Count -eq 0) { return $null }

    $best     = $null
    $bestTick = [long]::MinValue

    foreach ($s in $Sessions) {
        $endTime   = Get-SessionEndTime   -Session $s
        $startTime = Get-SessionStartTime -Session $s
        # Sessions with no timestamp at all are skipped — they should never
        # beat a session that has a real end or start time.
        if ($null -eq $endTime -and $null -eq $startTime) { continue }
        $t = if ($null -ne $endTime) { Get-SortableTicks -Value $endTime } `
             else                    { Get-SortableTicks -Value $startTime }
        if ($t -gt $bestTick) { $bestTick = $t; $best = $s }
    }

    return $best
}

# ---------------------------------------------------------------------------
# Get-DJSessionMatchScore
#   Returns > 0 when $Session can be matched to $Job; 0 means no match.
#   Tries (highest to lowest): GUID IDs → job-name equality → session-name prefix.
# ---------------------------------------------------------------------------
function Get-DJSessionMatchScore {
    [CmdletBinding()]
    param(
        [object]$Session,
        [object]$Job
    )

    if ($null -eq $Session -or $null -eq $Job) { return 0 }

    $jobId   = Get-DJPropertyPathValue -Object $Job -Path 'Id'
    $jobName = if ($null -ne $Job.PSObject.Properties['Name']) { [string]$Job.Name } else { $null }

    if ($null -ne $jobId) {
        $jobIdStr = [string]$jobId
        foreach ($path in @('JobId', 'Info.JobId', 'Info.Id')) {
            $val = Get-DJPropertyPathValue -Object $Session -Path $path
            if ($null -ne $val -and [string]$val -eq $jobIdStr) { return 3 }
        }
    }

    if (-not [string]::IsNullOrWhiteSpace($jobName)) {
        foreach ($path in @('JobName', 'Info.JobName')) {
            $val = Get-DJPropertyPathValue -Object $Session -Path $path
            if ($null -ne $val -and [string]$val -ieq $jobName) { return 2 }
        }
        $sesName = Get-DJPropertyPathValue -Object $Session -Path 'Name'
        if ($null -ne $sesName) {
            $sn = [string]$sesName
            if ($sn -ieq $jobName -or $sn -ilike "$jobName@*" -or $sn -ilike "$jobName *") { return 1 }
        }
    }

    return 0
}

# ---------------------------------------------------------------------------
# Get-DJLatestSessionForJob
#   Returns the latest session object for a job using type-specific cmdlets.
#   Uses pre-cached session arrays (populated by Get-DefinedJobsReport before
#   the per-job loops) to avoid repeated expensive cmdlet calls.
#   TypeOverride controls which retrieval path is attempted first.
# ---------------------------------------------------------------------------
function Get-DJLatestSessionForJob {
    [CmdletBinding()]
    param(
        [object]$Job,
        [string]$TypeOverride
    )

    $jobName = if ($null -ne $Job.PSObject.Properties['Name']) { [string]$Job.Name } else { $null }

    # -- Agent / computer backup jobs --
    if ($TypeOverride -eq 'Agent' -and $null -ne $script:DJCachedAgentSessions) {
        $matched = @($script:DJCachedAgentSessions | Where-Object { (Get-DJSessionMatchScore -Session $_ -Job $Job) -gt 0 })
        if ($matched.Count -gt 0) { return Get-DJLatestSessionFromList -Sessions $matched }
    }

    # -- Application backup jobs --
    if ($TypeOverride -eq 'Application' -and $null -ne $script:DJCachedAppSessions) {
        $matched = @($script:DJCachedAppSessions | Where-Object { (Get-DJSessionMatchScore -Session $_ -Job $Job) -gt 0 })
        if ($matched.Count -gt 0) { return Get-DJLatestSessionFromList -Sessions $matched }
    }

    # -- File / NAS / unstructured jobs --
    if ($TypeOverride -eq 'File/NAS') {
        if (Get-Command -Name 'Get-VBRUnstructuredBackupSession' -ErrorAction SilentlyContinue) {
            try {
                if (-not [string]::IsNullOrWhiteSpace($jobName)) {
                    $sessions = @(Get-VBRUnstructuredBackupSession -Name "$($jobName)*" -ErrorAction SilentlyContinue)
                    if ($sessions.Count -gt 0) { return Get-DJLatestSessionFromList -Sessions $sessions }
                }
            } catch {
                Write-DebugMessage ('[Get-DJLatestSessionForJob] Get-VBRUnstructuredBackupSession failed for "{0}": {1}' -f $jobName, $_.Exception.Message)
            }
        }
    }

    # -- Standard VBR jobs: FindLastSession() first (most accurate) --
    if ($Job.PSObject.Methods['FindLastSession']) {
        try {
            $s = $Job.FindLastSession()
            if ($null -ne $s) { return $s }
        } catch {}
    }

    # -- Generic fallback: match against pre-cached VBR backup sessions --
    if ($null -ne $script:DJCachedVBRSessions) {
        $matched = @($script:DJCachedVBRSessions | Where-Object { (Get-DJSessionMatchScore -Session $_ -Job $Job) -gt 0 })
        if ($matched.Count -gt 0) { return Get-DJLatestSessionFromList -Sessions $matched }
    }

    return $null
}

# ---------------------------------------------------------------------------
# Get-DJLastRunForJob
#   Returns a hash-table @{ LastRun; Status; LastResult } for a job.
#   Uses Get-DJLatestSessionForJob (type-specific cmdlets + FindLastSession) so
#   that Agent, Application, and File/NAS jobs return populated values.
#   Status  = current session state (Running/Stopped/Idle).
#   LastResult = actual job result (Success/Warning/Failed/None).
# ---------------------------------------------------------------------------
function Get-DJLastRunForJob {
    [CmdletBinding()]
    param(
        [object]$Job,
        [string]$TypeOverride = ''
    )

    $lastSession = Get-DJLatestSessionForJob -Job $Job -TypeOverride $TypeOverride

    if ($null -ne $lastSession) {
        $endTime   = Get-SessionEndTime   -Session $lastSession
        $startTime = Get-SessionStartTime -Session $lastSession
        $runTime   = if ($null -ne $endTime) { $endTime } elseif ($null -ne $startTime) { $startTime } else { $null }
        return @{
            LastRun    = if ($null -ne $runTime) { $runTime.ToString($script:DJDateFormat) } else { '' }
            Status     = Get-DJSessionStateValue -Session $lastSession
            LastResult = Get-DJSessionResult     -Session $lastSession
        }
    }

    # Fall back to direct job properties
    $lastRun = Get-DJFirstValidDate -Values @(
        (Get-DJPropertyPathValue -Object $Job -Path 'LastRun'),
        (Get-DJPropertyPathValue -Object $Job -Path 'LastRunTime')
    )
    $lastResult = Get-DJPropertyPathValue -Object $Job -Path 'LastResult'

    return @{
        LastRun    = if ($null -ne $lastRun) { $lastRun.ToString($script:DJDateFormat) } else { '' }
        Status     = ''
        LastResult = if ($null -ne $lastResult) { [string]$lastResult } else { '' }
    }
}

# ---------------------------------------------------------------------------
# Get-DJJobRepository
#   Returns the target repository name for a job.
#   Tries GetTargetRepository() method first, then multiple property paths.
#   Returns empty string when the repository cannot be determined.
# ---------------------------------------------------------------------------
function Get-DJJobRepository {
    [CmdletBinding()]
    param([object]$Job)

    if ($null -eq $Job) { return '' }

    # Try GetTargetRepository() method (most reliable for standard VBR jobs)
    if ($Job.PSObject.Methods['GetTargetRepository']) {
        try {
            $repo = $Job.GetTargetRepository()
            if ($null -ne $repo) {
                $rName = Get-DJPropertyPathValue -Object $repo -Path 'Name'
                if (-not [string]::IsNullOrWhiteSpace([string]$rName)) { return [string]$rName }
            }
        } catch {}
    }

    # Try common property paths across job types and versions
    foreach ($path in @(
        'TargetRepository.Name',
        'Repository.Name',
        'BackupStorageOptions.RepositoryFriendlyName',
        'BackupStorageOptions.Repository.Name',
        'Info.BackupTargetOptions.RepositoryName',
        'BackupTarget.Repository.Name',
        'BackupTarget.Name',
        'RepositoryName'
    )) {
        try {
            $val = Get-DJPropertyPathValue -Object $Job -Path $path
            if (-not [string]::IsNullOrWhiteSpace([string]$val)) { return [string]$val }
        } catch {}
    }

    return ''
}

# ---------------------------------------------------------------------------
# New-DJJobReportRow
#   Builds one [pscustomobject] row for the Defined Jobs table.
# ---------------------------------------------------------------------------
function New-DJJobReportRow {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [object]$Job,
        [Parameter(Mandatory)] [string]$TypeOverride
    )

    $jobName      = if ($null -ne $Job.PSObject.Properties['Name']) { [string]$Job.Name } else { '<unnamed>' }
    $schedEnabled = Get-DJScheduleEnabled  -Job $Job
    $schedDisplay = Get-DJScheduleDisplay  -Job $Job
    $lastRunInfo  = Get-DJLastRunForJob    -Job $Job -TypeOverride $TypeOverride

    return [pscustomobject][ordered]@{
        Job        = $jobName
        Type       = $TypeOverride
        On         = if ($schedEnabled) { 'Yes' } else { 'No' }
        Schedule   = $schedDisplay
        LastRun    = $lastRunInfo.LastRun
        Status     = $lastRunInfo.Status
        LastResult = $lastRunInfo.LastResult
    }
}

# ---------------------------------------------------------------------------
# Get-DefinedJobsReport
#   Enumerates all defined backup jobs and returns an array of report rows
#   sorted by job name.  Collects specialised job types first (agent,
#   application, unstructured) then remaining standard backup jobs, excluding
#   replication/copy/tape/SureBackup types.  Every optional cmdlet is guarded
#   with Get-Command and wrapped in try/catch so failures are non-fatal.
# ---------------------------------------------------------------------------
function Get-DefinedJobsReport {
    [CmdletBinding()]
    param()

    $rows      = New-Object 'System.Collections.Generic.List[object]'
    $seenNames = New-Object 'System.Collections.Generic.HashSet[string]'([System.StringComparer]::OrdinalIgnoreCase)

    # Pre-fetch sessions once per type to avoid repeated cmdlet calls in the
    # per-job row-builder. $null means the cmdlet is unavailable; @() means it
    # returned no sessions.  Get-DJLatestSessionForJob checks for $null before
    # using a cache so missing cmdlets are handled gracefully.

    $script:DJCachedAgentSessions = $null
    if (Get-Command -Name 'Get-VBRComputerBackupJobSession' -ErrorAction SilentlyContinue) {
        try {
            $script:DJCachedAgentSessions = @(Get-VBRComputerBackupJobSession -ErrorAction SilentlyContinue)
            Write-DebugMessage ('[Get-DefinedJobsReport] Cached {0} agent session(s).' -f $script:DJCachedAgentSessions.Count)
        } catch {
            Write-DebugMessage ('[Get-DefinedJobsReport] Get-VBRComputerBackupJobSession cache failed: {0}' -f $_.Exception.Message)
            $script:DJCachedAgentSessions = @()
        }
    }

    $script:DJCachedAppSessions = $null
    if (Get-Command -Name 'Get-VBRApplicationBackupJobSession' -ErrorAction SilentlyContinue) {
        try {
            $script:DJCachedAppSessions = @(Get-VBRApplicationBackupJobSession -ErrorAction SilentlyContinue)
            Write-DebugMessage ('[Get-DefinedJobsReport] Cached {0} application session(s).' -f $script:DJCachedAppSessions.Count)
        } catch {
            Write-DebugMessage ('[Get-DefinedJobsReport] Get-VBRApplicationBackupJobSession cache failed: {0}' -f $_.Exception.Message)
            $script:DJCachedAppSessions = @()
        }
    }

    $script:DJCachedVBRSessions = $null
    if (Get-Command -Name 'Get-VBRBackupSession' -ErrorAction SilentlyContinue) {
        try {
            $script:DJCachedVBRSessions = @(Get-VBRBackupSession -ErrorAction SilentlyContinue)
            Write-DebugMessage ('[Get-DefinedJobsReport] Cached {0} VBR backup session(s).' -f $script:DJCachedVBRSessions.Count)
        } catch {
            Write-DebugMessage ('[Get-DefinedJobsReport] Get-VBRBackupSession cache failed: {0}' -f $_.Exception.Message)
            $script:DJCachedVBRSessions = @()
        }
    }

    # ---- Agent / computer backup jobs ----
    if (Get-Command -Name 'Get-VBRComputerBackupJob' -ErrorAction SilentlyContinue) {
        Write-DebugMessage '[Get-DefinedJobsReport] Enumerating Get-VBRComputerBackupJob.'
        try {
            $agentJobs = @(Get-VBRComputerBackupJob -ErrorAction SilentlyContinue -WarningAction SilentlyContinue)
            Write-DebugMessage ('[Get-DefinedJobsReport] Agent jobs: {0}' -f $agentJobs.Count)
            foreach ($job in $agentJobs) {
                $jn = if ($null -ne $job.PSObject.Properties['Name']) { [string]$job.Name } else { '' }
                if ([string]::IsNullOrWhiteSpace($jn)) { continue }
                if (-not $seenNames.Add($jn)) { continue }
                try {
                    $row = New-DJJobReportRow -Job $job -TypeOverride 'Agent'
                    [void]$rows.Add($row)
                } catch {
                    Write-DebugMessage ('[Get-DefinedJobsReport] Row build failed for agent job "{0}": {1}' -f $jn, $_.Exception.Message)
                }
            }
        } catch {
            Write-DebugMessage ('[Get-DefinedJobsReport] Get-VBRComputerBackupJob failed: {0}' -f $_.Exception.Message)
        }
    }

    # ---- Application backup jobs ----
    if (Get-Command -Name 'Get-VBRApplicationBackupJob' -ErrorAction SilentlyContinue) {
        Write-DebugMessage '[Get-DefinedJobsReport] Enumerating Get-VBRApplicationBackupJob.'
        try {
            $appJobs = @(Get-VBRApplicationBackupJob -ErrorAction SilentlyContinue -WarningAction SilentlyContinue)
            Write-DebugMessage ('[Get-DefinedJobsReport] Application jobs: {0}' -f $appJobs.Count)
            foreach ($job in $appJobs) {
                $jn = if ($null -ne $job.PSObject.Properties['Name']) { [string]$job.Name } else { '' }
                if ([string]::IsNullOrWhiteSpace($jn)) { continue }
                if (-not $seenNames.Add($jn)) { continue }
                try {
                    $row = New-DJJobReportRow -Job $job -TypeOverride 'Application'
                    [void]$rows.Add($row)
                } catch {
                    Write-DebugMessage ('[Get-DefinedJobsReport] Row build failed for app job "{0}": {1}' -f $jn, $_.Exception.Message)
                }
            }
        } catch {
            Write-DebugMessage ('[Get-DefinedJobsReport] Get-VBRApplicationBackupJob failed: {0}' -f $_.Exception.Message)
        }
    }

    # ---- Unstructured backup jobs ----
    if (Get-Command -Name 'Get-VBRUnstructuredBackupJob' -ErrorAction SilentlyContinue) {
        Write-DebugMessage '[Get-DefinedJobsReport] Enumerating Get-VBRUnstructuredBackupJob.'
        try {
            $unstrJobs = @(Get-VBRUnstructuredBackupJob -Name '*' -ErrorAction SilentlyContinue -WarningAction SilentlyContinue)
            Write-DebugMessage ('[Get-DefinedJobsReport] Unstructured jobs: {0}' -f $unstrJobs.Count)
            foreach ($job in $unstrJobs) {
                $jn = if ($null -ne $job.PSObject.Properties['Name']) { [string]$job.Name } else { '' }
                if ([string]::IsNullOrWhiteSpace($jn)) { continue }
                if (-not $seenNames.Add($jn)) { continue }
                try {
                    $row = New-DJJobReportRow -Job $job -TypeOverride 'File/NAS'
                    [void]$rows.Add($row)
                } catch {
                    Write-DebugMessage ('[Get-DefinedJobsReport] Row build failed for unstructured job "{0}": {1}' -f $jn, $_.Exception.Message)
                }
            }
        } catch {
            Write-DebugMessage ('[Get-DefinedJobsReport] Get-VBRUnstructuredBackupJob failed: {0}' -f $_.Exception.Message)
        }
    }

    # ---- Remaining standard VBR jobs (excluding replication/copy/tape/SureBackup) ----
    # These non-backup job types are excluded because they are either captured by dedicated
    # phases in the collector (replication, backup copy, tape) or are infrastructure-only
    # jobs whose sessions are not meaningfully surfaced as backup job sessions.
    if (Get-Command -Name 'Get-VBRJob' -ErrorAction SilentlyContinue) {
        Write-DebugMessage '[Get-DefinedJobsReport] Enumerating Get-VBRJob (standard).'
        try {
            $vbrJobs       = @(Get-VBRJob -ErrorAction SilentlyContinue -WarningAction SilentlyContinue)
            $excludePattern = 'Replica|Replication|BackupSync|BackupCopy|FileCopy|VmCopy|Tape|SureBackup'
            Write-DebugMessage ('[Get-DefinedJobsReport] Standard VBR jobs: {0}' -f $vbrJobs.Count)
            foreach ($job in $vbrJobs) {
                $jn = if ($null -ne $job.PSObject.Properties['Name']) { [string]$job.Name } else { '' }
                if ([string]::IsNullOrWhiteSpace($jn)) { continue }
                if (-not $seenNames.Add($jn)) { continue }
                $jType = Get-DJStandardJobType -Job $job
                if ($jType -match $excludePattern) {
                    Write-DebugMessage ('[Get-DefinedJobsReport] Skipping excluded type "{0}" for job "{1}".' -f $jType, $jn)
                    continue
                }
                try {
                    $row = New-DJJobReportRow -Job $job -TypeOverride $jType
                    [void]$rows.Add($row)
                } catch {
                    Write-DebugMessage ('[Get-DefinedJobsReport] Row build failed for VBR job "{0}": {1}' -f $jn, $_.Exception.Message)
                }
            }
        } catch {
            Write-DebugMessage ('[Get-DefinedJobsReport] Get-VBRJob failed: {0}' -f $_.Exception.Message)
        }
    }

    Write-DebugMessage ('[Get-DefinedJobsReport] Total rows collected: {0}' -f $rows.Count)
    $sorted = @($rows | Sort-Object -Property Job)
    return $sorted
}

# ---------------------------------------------------------------------------
# New-DefinedJobsSectionText
#   Builds and returns the Defined Jobs baseline block as a single string.
#   The block is delimited with the required markers and contains a fixed-width
#   table with columns: Job(38), Type(11), On(3), Schedule(18), LastRun(16),
#   Status(8).  Returns empty string in JSON mode or when collection fails.
# ---------------------------------------------------------------------------
function New-DefinedJobsSectionText {
    [CmdletBinding()]
    param()

    if ($Json) { return '' }

    try {
        Write-ProgressMessage 'Defined Jobs — collecting job inventory...'
        Write-DebugMessage '[New-DefinedJobsSectionText] Starting Defined Jobs collection.'

        $rows = @(Get-DefinedJobsReport)

        Write-ProgressMessage ('Defined Jobs — {0} job(s) found.' -f $rows.Count)
        Write-DebugMessage ('[New-DefinedJobsSectionText] Collection complete: {0} row(s).' -f $rows.Count)

        # Column widths
        $wJob    = 38
        $wType   = 11
        $wOn     = 3
        $wSch    = 18
        $wLast   = 16
        $wStat   = 11
        $wResult = 11

        $lines = New-Object 'System.Collections.Generic.List[string]'
        [void]$lines.Add('############### Defined Jobs BEGIN ###################')

        # Header row
        [void]$lines.Add(('{0} {1} {2} {3} {4} {5} {6}' -f
            'Job'.PadRight($wJob),
            'Type'.PadRight($wType),
            'On'.PadRight($wOn),
            'Next / schedule'.PadRight($wSch),
            'Last run'.PadRight($wLast),
            'Status'.PadRight($wStat),
            'Last Result'.PadRight($wResult)))

        # Separator
        [void]$lines.Add('-' * ($wJob + $wType + $wOn + $wSch + $wLast + $wStat + $wResult + 6))

        if ($rows.Count -eq 0) {
            [void]$lines.Add('(no defined jobs found)')
        } else {
            foreach ($r in $rows) {
                $jobCol    = [string]$r.Job
                $typCol    = [string]$r.Type
                $onCol     = [string]$r.On
                $schCol    = [string]$r.Schedule
                $lastCol   = [string]$r.LastRun
                $statCol   = [string]$r.Status
                $resultCol = [string]$r.LastResult

                $jobField    = if ($jobCol.Length    -gt $wJob)    { $jobCol.Substring(0, $wJob)       } else { $jobCol.PadRight($wJob)       }
                $typeField   = if ($typCol.Length    -gt $wType)   { $typCol.Substring(0, $wType)      } else { $typCol.PadRight($wType)      }
                $onField     = if ($onCol.Length     -gt $wOn)     { $onCol.Substring(0, $wOn)         } else { $onCol.PadRight($wOn)         }
                $schField    = if ($schCol.Length    -gt $wSch)    { $schCol.Substring(0, $wSch)       } else { $schCol.PadRight($wSch)       }
                $lastField   = if ($lastCol.Length   -gt $wLast)   { $lastCol.Substring(0, $wLast)     } else { $lastCol.PadRight($wLast)     }
                $statField   = if ($statCol.Length   -gt $wStat)   { $statCol.Substring(0, $wStat)     } else { $statCol.PadRight($wStat)     }
                $resultField = if ($resultCol.Length -gt $wResult) { $resultCol.Substring(0, $wResult) } else { $resultCol.PadRight($wResult) }

                [void]$lines.Add(('{0} {1} {2} {3} {4} {5} {6}' -f $jobField, $typeField, $onField, $schField, $lastField, $statField, $resultField))
            }
        }

        [void]$lines.Add('############### Defined Jobs END ###################')

        return ($lines -join [Environment]::NewLine)
    } catch {
        Write-Warning ('Defined Jobs baseline failed to build: {0}' -f $_.Exception.Message)
        Write-DebugMessage ('[New-DefinedJobsSectionText] Failed:' + [Environment]::NewLine + (Format-ErrorRecord -ErrorRecord $_))
        return ''
    }
}

# ===========================================================================
# Defined Repository baseline — helper functions
#   These functions are repository-specific and DR-prefixed to avoid collisions
#   with existing helpers from other collector phases.
# ===========================================================================

function Get-DRPropertyPathValue {
    [CmdletBinding()]
    param(
        [object]$Object,
        [string]$Path
    )

    if ($null -eq $Object) { return $null }
    if ([string]::IsNullOrWhiteSpace($Path)) { return $null }

    $parts = $Path -split '\.'
    $current = $Object
    foreach ($part in $parts) {
        if ($null -eq $current) { return $null }
        try {
            $prop = $current.PSObject.Properties[$part]
            if ($null -eq $prop) { return $null }
            $current = $prop.Value
        } catch {
            return $null
        }
    }

    return $current
}

function ConvertTo-DRNullableBoolean {
    [CmdletBinding()]
    param([object]$Value)

    if ($null -eq $Value) { return $null }
    try { return [bool]$Value } catch { return $null }
}

function ConvertTo-DRSizeBytes {
    [CmdletBinding()]
    param([object]$Value)

    if ($null -eq $Value) { return $null }

    if ($Value -is [byte] -or
        $Value -is [int16] -or
        $Value -is [int32] -or
        $Value -is [int64] -or
        $Value -is [uint16] -or
        $Value -is [uint32] -or
        $Value -is [uint64] -or
        $Value -is [single] -or
        $Value -is [double] -or
        $Value -is [decimal]) {
        return [double]$Value
    }

    foreach ($path in @('InBytes', 'Bytes', 'Value', 'Size', 'TotalBytes', 'UsedBytes', 'FreeBytes')) {
        $nested = Get-DRPropertyPathValue -Object $Value -Path $path
        if ($null -ne $nested -and -not [object]::ReferenceEquals($nested, $Value)) {
            $converted = ConvertTo-DRSizeBytes -Value $nested
            if ($null -ne $converted) { return $converted }
        }
    }

    if ($Value.PSObject.Methods['ToBytes']) {
        try {
            $toBytes = $Value.ToBytes()
            if ($null -ne $toBytes) { return [double]$toBytes }
        } catch {}
    }

    $text = [string]$Value
    if ([string]::IsNullOrWhiteSpace($text)) { return $null }
    $text = $text.Trim()
    $textNoComma = $text -replace ',', ''

    $match = [regex]::Match($textNoComma, '^\s*(?<num>[-+]?\d+(\.\d+)?)\s*(?<unit>[KMGTP]?i?B)?\s*$', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    if (-not $match.Success) { return $null }

    try {
        $num  = [double]$match.Groups['num'].Value
        $unit = [string]$match.Groups['unit'].Value
        if ([string]::IsNullOrWhiteSpace($unit)) { return $num }

        switch -Regex ($unit.ToUpperInvariant()) {
            '^B$'   { return $num }
            '^KIB$' { return ($num * 1024) }
            '^KB$'  { return ($num * 1024) }
            '^MIB$' { return ($num * 1024 * 1024) }
            '^MB$'  { return ($num * 1024 * 1024) }
            '^GIB$' { return ($num * 1024 * 1024 * 1024) }
            '^GB$'  { return ($num * 1024 * 1024 * 1024) }
            '^TIB$' { return ($num * 1024 * 1024 * 1024 * 1024) }
            '^TB$'  { return ($num * 1024 * 1024 * 1024 * 1024) }
            '^PIB$' { return ($num * 1024 * 1024 * 1024 * 1024 * 1024) }
            '^PB$'  { return ($num * 1024 * 1024 * 1024 * 1024 * 1024) }
            default { return $null }
        }
    } catch {
        return $null
    }
}

function Get-DRFirstSizeBytes {
    [CmdletBinding()]
    param([object[]]$Values)

    foreach ($v in $Values) {
        $bytes = ConvertTo-DRSizeBytes -Value $v
        if ($null -ne $bytes) { return [double]$bytes }
    }
    return $null
}

function ConvertTo-DRBytes {
    [CmdletBinding()]
    param([object]$Value)

    return (ConvertTo-DRSizeBytes -Value $Value)
}

function Get-DRFirstSizeValue {
    [CmdletBinding()]
    param([object[]]$Values)

    return (Get-DRFirstSizeBytes -Values $Values)
}

function Format-DRByteSize {
    [CmdletBinding()]
    param([object]$Bytes)

    $value = ConvertTo-DRBytes -Value $Bytes
    if ($null -eq $value) { return '' }
    if ($value -lt 0) { $value = 0 }

    $units = @('B', 'KiB', 'MiB', 'GiB', 'TiB', 'PiB')
    $idx = 0
    while ($value -ge 1024 -and $idx -lt ($units.Count - 1)) {
        $value = $value / 1024
        $idx++
    }

    if ($idx -eq 0) {
        return ('{0:N0} {1}' -f $value, $units[$idx])
    }
    return ('{0:N2} {1}' -f $value, $units[$idx])
}

function Format-DRUsedPercent {
    [CmdletBinding()]
    param(
        [object]$TotalBytes,
        [object]$UsedBytes
    )

    $total = ConvertTo-DRBytes -Value $TotalBytes
    $used  = ConvertTo-DRBytes -Value $UsedBytes

    if ($null -eq $total -or $null -eq $used -or $total -le 0) { return '' }
    $pct = ($used / $total) * 100
    return ('{0:N2}%' -f $pct)
}

function Get-DRNonNegativeDifference {
    [CmdletBinding()]
    param(
        [Alias('TotalBytes')]
        [AllowNull()]
        [object]$Left,

        [Alias('UsedBytes')]
        [AllowNull()]
        [object]$Right
    )

    $leftValue  = ConvertTo-DRBytes -Value $Left
    $rightValue = ConvertTo-DRBytes -Value $Right

    if ($null -eq $leftValue -or $null -eq $rightValue) { return $null }
    [double]$result = [double]$leftValue - [double]$rightValue
    if ($result -lt 0.0) { return [double]0.0 }
    return $result
}

function Get-DRRepositoryName {
    [CmdletBinding()]
    param([object]$Repository)

    if ($null -eq $Repository) { return '' }
    foreach ($path in @('Name', 'FriendlyName')) {
        $v = Get-DRPropertyPathValue -Object $Repository -Path $path
        if (-not [string]::IsNullOrWhiteSpace([string]$v)) { return [string]$v }
    }
    return [string]$Repository
}

function Get-DRRepositoryKey {
    [CmdletBinding()]
    param([object]$Repository)

    if ($null -eq $Repository) { return '' }
    foreach ($path in @('Id', 'Uid', 'Info.Id', 'Name', 'FriendlyName')) {
        $v = Get-DRPropertyPathValue -Object $Repository -Path $path
        if (-not [string]::IsNullOrWhiteSpace([string]$v)) { return [string]$v }
    }
    $location = Get-DRRepositoryLocation -Repository $Repository
    if (-not [string]::IsNullOrWhiteSpace($location)) { return $location }
    return [string]$Repository.GetHashCode()
}

function Get-DRRepositoryLocation {
    [CmdletBinding()]
    param([object]$Repository)

    if ($null -eq $Repository) { return '' }

    foreach ($path in @(
        'Path',
        'Repository.Path',
        'Info.Path',
        'Container.Path',
        'BucketName',
        'Folder',
        'SharePath',
        'Host.Name'
    )) {
        $value = Get-DRPropertyPathValue -Object $Repository -Path $path
        if (-not [string]::IsNullOrWhiteSpace([string]$value)) { return [string]$value }
    }

    return ''
}

function Get-DRRepositoryStatus {
    [CmdletBinding()]
    param([object]$Repository)

    if ($null -eq $Repository) { return '' }

    $isUnavailable = ConvertTo-DRNullableBoolean -Value (Get-DRPropertyPathValue -Object $Repository -Path 'IsUnavailable')
    if ($isUnavailable) { return 'Unavailable' }

    $isOutOfOrder = ConvertTo-DRNullableBoolean -Value (Get-DRPropertyPathValue -Object $Repository -Path 'IsOutOfOrder')
    if ($isOutOfOrder) { return 'OutOfOrder' }

    $isMaintenance = ConvertTo-DRNullableBoolean -Value (Get-DRPropertyPathValue -Object $Repository -Path 'IsMaintenanceMode')
    if ($isMaintenance) { return 'Maintenance' }

    foreach ($path in @('Status', 'State', 'Info.Status')) {
        $s = Get-DRPropertyPathValue -Object $Repository -Path $path
        if (-not [string]::IsNullOrWhiteSpace([string]$s)) { return [string]$s }
    }

    $enabled = ConvertTo-DRNullableBoolean -Value (Get-DRPropertyPathValue -Object $Repository -Path 'Enabled')
    if ($null -eq $enabled) {
        $enabled = ConvertTo-DRNullableBoolean -Value (Get-DRPropertyPathValue -Object $Repository -Path 'IsEnabled')
    }
    if ($enabled -eq $false) { return 'Disabled' }
    if ($enabled -eq $true)  { return 'Ready' }

    return ''
}

function Get-DRRepositorySpaceInfo {
    [CmdletBinding()]
    param(
        [object]$Repository,
        [switch]$ObjectStorage
    )

    if ($null -eq $Repository) {
        return @{ TotalBytes = $null; UsedBytes = $null; FreeBytes = $null }
    }

    $container = $null
    if ($Repository.PSObject.Methods['GetContainer']) {
        try {
            $container = $Repository.GetContainer()
        } catch {
            Write-DebugMessage ('[Get-DRRepositorySpaceInfo] GetContainer() failed for "{0}": {1}' -f (Get-DRRepositoryName -Repository $Repository), $_.Exception.Message)
        }
    }

    $total = Get-DRFirstSizeValue -Values @(
        (Get-DRPropertyPathValue -Object $Repository -Path 'CachedTotalSpace'),
        (Get-DRPropertyPathValue -Object $Repository -Path 'TotalSpace'),
        (Get-DRPropertyPathValue -Object $Repository -Path 'Container.TotalSpace'),
        (Get-DRPropertyPathValue -Object $container  -Path 'CachedTotalSpace'),
        (Get-DRPropertyPathValue -Object $container  -Path 'TotalSpace'),
        (Get-DRPropertyPathValue -Object $Repository -Path 'Capacity'),
        (Get-DRPropertyPathValue -Object $Repository -Path 'Quota'),
        (Get-DRPropertyPathValue -Object $Repository -Path 'Info.TotalSpace')
    )
    $used = Get-DRFirstSizeValue -Values @(
        (Get-DRPropertyPathValue -Object $Repository -Path 'CachedUsedSpace'),
        (Get-DRPropertyPathValue -Object $Repository -Path 'UsedSpace'),
        (Get-DRPropertyPathValue -Object $Repository -Path 'Container.UsedSpace'),
        (Get-DRPropertyPathValue -Object $container  -Path 'CachedUsedSpace'),
        (Get-DRPropertyPathValue -Object $container  -Path 'UsedSpace'),
        (Get-DRPropertyPathValue -Object $Repository -Path 'UsedSize'),
        (Get-DRPropertyPathValue -Object $Repository -Path 'Info.UsedSpace')
    )
    $free = Get-DRFirstSizeValue -Values @(
        (Get-DRPropertyPathValue -Object $Repository -Path 'CachedFreeSpace'),
        (Get-DRPropertyPathValue -Object $Repository -Path 'FreeSpace'),
        (Get-DRPropertyPathValue -Object $Repository -Path 'Container.FreeSpace'),
        (Get-DRPropertyPathValue -Object $container  -Path 'CachedFreeSpace'),
        (Get-DRPropertyPathValue -Object $container  -Path 'FreeSpace'),
        (Get-DRPropertyPathValue -Object $Repository -Path 'AvailableSpace'),
        (Get-DRPropertyPathValue -Object $Repository -Path 'Info.FreeSpace')
    )

    if ($ObjectStorage -and $null -eq $total) {
        $sizeLimitEnabled = $null
        foreach ($path in @(
            'ObjectStorageSettings.SizeLimitEnabled',
            'CapacityTierSettings.SizeLimitEnabled',
            'SizeLimitEnabled',
            'UseSizeLimit'
        )) {
            $sizeLimitEnabled = ConvertTo-DRNullableBoolean -Value (Get-DRPropertyPathValue -Object $Repository -Path $path)
            if ($null -ne $sizeLimitEnabled) { break }
        }

        if ($sizeLimitEnabled) {
            $sizeLimit = Get-DRFirstSizeValue -Values @(
                (Get-DRPropertyPathValue -Object $Repository -Path 'ObjectStorageSettings.SizeLimit'),
                (Get-DRPropertyPathValue -Object $Repository -Path 'CapacityTierSettings.SizeLimit'),
                (Get-DRPropertyPathValue -Object $Repository -Path 'SizeLimit'),
                (Get-DRPropertyPathValue -Object $Repository -Path 'StorageLimit')
            )
            if ($null -ne $sizeLimit) { $total = $sizeLimit }
        }
    }

    if ($null -eq $used -and $null -ne $total -and $null -ne $free) {
        $used = Get-DRNonNegativeDifference -Left ([double]$total) -Right ([double]$free)
    }
    if ($null -eq $free -and $null -ne $total -and $null -ne $used) {
        $free = Get-DRNonNegativeDifference -Left ([double]$total) -Right ([double]$used)
    }

    return @{
        TotalBytes = $total
        UsedBytes  = $used
        FreeBytes  = $free
    }
}

function Get-DRPhysicalRepositorySpace {
    [CmdletBinding()]
    param([object]$Repository)

    return (Get-DRRepositorySpaceInfo -Repository $Repository)
}

function Get-DRObjectRepositorySpace {
    [CmdletBinding()]
    param([object]$Repository)

    return (Get-DRRepositorySpaceInfo -Repository $Repository -ObjectStorage)
}

function New-DRRepositoryRow {
    [CmdletBinding()]
    param(
        [string]$Repository,
        [string]$Tier,
        [string]$Parent,
        [string]$Status,
        [object]$TotalBytes,
        [object]$UsedBytes,
        [object]$FreeBytes,
        [int]$SortGroup,
        [int]$SortOrder
    )

    return [pscustomobject][ordered]@{
        Repository = $Repository
        Tier       = $Tier
        Parent     = $Parent
        Status     = $Status
        Total      = Format-DRByteSize    -Bytes $TotalBytes
        Used       = Format-DRByteSize    -Bytes $UsedBytes
        Free       = Format-DRByteSize    -Bytes $FreeBytes
        'Used %'   = Format-DRUsedPercent -TotalBytes $TotalBytes -UsedBytes $UsedBytes
        SortGroup  = $SortGroup
        SortOrder  = $SortOrder
    }
}

function Format-DRFixedWidth {
    [CmdletBinding()]
    param(
        [AllowNull()] [string]$Value,
        [Parameter(Mandatory)] [int]$Width
    )

    $text = if ($null -ne $Value) { [string]$Value } else { '' }
    if ($text.Length -gt $Width) { return $text.Substring(0, $Width) }
    return $text.PadRight($Width)
}

function New-DefinedRepositoryPlaceholderSection {
    [CmdletBinding()]
    param(
        [string]$Message = '(repository utilisation unavailable)'
    )

    $lines = New-Object 'System.Collections.Generic.List[string]'
    [void]$lines.Add('############### Defined Repository BEGIN ###################')
    [void]$lines.Add($Message)
    [void]$lines.Add('############### Defined Repository END ###################')
    return ($lines -join [Environment]::NewLine)
}

function Get-DefinedRepositoryReport {
    [CmdletBinding()]
    param()

    $rows = New-Object 'System.Collections.Generic.List[object]'

    # Separate hash sets for physical extents attached to SOBRs and object extents
    # attached to SOBRs (either as capacity extents or object-backed performance extents).
    $attachedPhysRepoKeys = New-Object 'System.Collections.Generic.HashSet[string]'([System.StringComparer]::OrdinalIgnoreCase)
    $attachedObjRepoKeys  = New-Object 'System.Collections.Generic.HashSet[string]'([System.StringComparer]::OrdinalIgnoreCase)

    # Build the object repository key lookup upfront so the standard-repository pass
    # can skip COS/S3 repos and prevent them from appearing as physical/standalone.
    $objectRepoKeys = New-Object 'System.Collections.Generic.HashSet[string]'([System.StringComparer]::OrdinalIgnoreCase)
    $objectRepoList = @()
    if (Get-Command -Name 'Get-VBRObjectStorageRepository' -ErrorAction SilentlyContinue) {
        try {
            $objectRepoList = @(Get-VBRObjectStorageRepository -ErrorAction SilentlyContinue -WarningAction SilentlyContinue)
            foreach ($objR in $objectRepoList) {
                $k = Get-DRRepositoryKey -Repository $objR
                if (-not [string]::IsNullOrWhiteSpace($k)) { [void]$objectRepoKeys.Add($k) }
            }
        } catch {
            Write-DebugMessage ('[Get-DefinedRepositoryReport] Get-VBRObjectStorageRepository (key build) failed: {0}' -f $_.Exception.Message)
            $objectRepoList = @()
        }
    }

    $sobrList = @()
    if (Get-Command -Name 'Get-VBRBackupRepository' -ErrorAction SilentlyContinue) {
        try {
            $sobrList = @(Get-VBRBackupRepository -ScaleOut -ErrorAction SilentlyContinue -WarningAction SilentlyContinue)
        } catch {
            Write-DebugMessage ('[Get-DefinedRepositoryReport] Get-VBRBackupRepository -ScaleOut failed: {0}' -f $_.Exception.Message)
            $sobrList = @()
        }
    }

    $sobrSort = 0
    foreach ($sobr in $sobrList) {
        $sobrSort++
        $sobrName   = Get-DRRepositoryName   -Repository $sobr
        $sobrStatus = Get-DRRepositoryStatus -Repository $sobr

        $perfTotal = $null
        $perfUsed  = $null

        $perfExtents = @()
        if (Get-Command -Name 'Get-VBRRepositoryExtent' -ErrorAction SilentlyContinue) {
            try {
                $perfExtents = @(Get-VBRRepositoryExtent -Repository $sobr -ErrorAction SilentlyContinue -WarningAction SilentlyContinue)
            } catch {
                Write-DebugMessage ('[Get-DefinedRepositoryReport] Get-VBRRepositoryExtent failed for SOBR "{0}": {1}' -f $sobrName, $_.Exception.Message)
                $perfExtents = @()
            }
        }

        $perfOrder = 0
        foreach ($extent in $perfExtents) {
            $perfOrder++
            $extentRepo   = Get-DRPropertyPathValue -Object $extent -Path 'Repository'
            if ($null -eq $extentRepo) { $extentRepo = $extent }
            $extentName   = Get-DRRepositoryName   -Repository $extentRepo
            $extentStatus = Get-DRRepositoryStatus -Repository $extentRepo
            $extentKey    = Get-DRRepositoryKey    -Repository $extentRepo

            # Determine whether this performance extent is backed by object storage.
            $extentIsObj = (-not [string]::IsNullOrWhiteSpace($extentKey)) -and $objectRepoKeys.Contains($extentKey)
            if (-not [string]::IsNullOrWhiteSpace($extentKey)) {
                if ($extentIsObj) {
                    [void]$attachedObjRepoKeys.Add($extentKey)
                } else {
                    [void]$attachedPhysRepoKeys.Add($extentKey)
                }
            }

            $extentTier = if ($extentIsObj) { 'Perf extent (Obj)' } else { 'Perf extent' }
            $space = if ($extentIsObj) {
                Get-DRObjectRepositorySpace   -Repository $extentRepo
            } else {
                Get-DRPhysicalRepositorySpace -Repository $extentRepo
            }

            [void]$rows.Add((New-DRRepositoryRow -Repository $extentName -Tier $extentTier -Parent $sobrName -Status $extentStatus `
                -TotalBytes $space.TotalBytes -UsedBytes $space.UsedBytes -FreeBytes $space.FreeBytes -SortGroup 2 -SortOrder ($sobrSort * 1000 + $perfOrder)))

            # Accumulate all performance extents (including object-backed) into the SOBR aggregate.
            # Capacity extents are intentionally excluded from this aggregate.
            if ($null -ne $space.TotalBytes) { $perfTotal = if ($null -eq $perfTotal) { 0 } else { $perfTotal }; $perfTotal += $space.TotalBytes }
            if ($null -ne $space.UsedBytes)  { $perfUsed  = if ($null -eq $perfUsed)  { 0 } else { $perfUsed  }; $perfUsed  += $space.UsedBytes  }
        }

        $capExtents = @()
        if (Get-Command -Name 'Get-VBRCapacityExtent' -ErrorAction SilentlyContinue) {
            try {
                $capExtents = @(Get-VBRCapacityExtent -Repository $sobr -ErrorAction SilentlyContinue -WarningAction SilentlyContinue)
            } catch {
                Write-DebugMessage ('[Get-DefinedRepositoryReport] Get-VBRCapacityExtent failed for SOBR "{0}": {1}' -f $sobrName, $_.Exception.Message)
                $capExtents = @()
            }
        }

        $capOrder = 0
        foreach ($cap in $capExtents) {
            $capOrder++
            $capRepo   = Get-DRPropertyPathValue -Object $cap -Path 'Repository'
            if ($null -eq $capRepo) { $capRepo = $cap }
            $capName   = Get-DRRepositoryName   -Repository $capRepo
            $capStatus = Get-DRRepositoryStatus -Repository $capRepo
            $capKey    = Get-DRRepositoryKey    -Repository $capRepo
            if (-not [string]::IsNullOrWhiteSpace($capKey)) { [void]$attachedObjRepoKeys.Add($capKey) }

            $capSpace = Get-DRObjectRepositorySpace -Repository $capRepo
            [void]$rows.Add((New-DRRepositoryRow -Repository $capName -Tier 'Capacity' -Parent $sobrName -Status $capStatus `
                -TotalBytes $capSpace.TotalBytes -UsedBytes $capSpace.UsedBytes -FreeBytes $capSpace.FreeBytes -SortGroup 3 -SortOrder ($sobrSort * 1000 + $capOrder)))
        }

        if ($null -eq $perfTotal -or $null -eq $perfUsed) {
            $sobrSpace = Get-DRPhysicalRepositorySpace -Repository $sobr
            if ($null -eq $perfTotal) { $perfTotal = $sobrSpace.TotalBytes }
            if ($null -eq $perfUsed)  { $perfUsed  = $sobrSpace.UsedBytes  }
            $perfFree = $sobrSpace.FreeBytes
        } else {
            $perfFree = Get-DRNonNegativeDifference -Left $perfTotal -Right $perfUsed
        }

        [void]$rows.Add((New-DRRepositoryRow -Repository $sobrName -Tier 'SOBR' -Parent '' -Status $sobrStatus `
            -TotalBytes $perfTotal -UsedBytes $perfUsed -FreeBytes $perfFree -SortGroup 1 -SortOrder $sobrSort))
    }

    if (Get-Command -Name 'Get-VBRBackupRepository' -ErrorAction SilentlyContinue) {
        try {
            $standardRepos = @(Get-VBRBackupRepository -ErrorAction SilentlyContinue -WarningAction SilentlyContinue)
        } catch {
            Write-DebugMessage ('[Get-DefinedRepositoryReport] Get-VBRBackupRepository failed: {0}' -f $_.Exception.Message)
            $standardRepos = @()
        }

        $standardOrder = 0
        foreach ($repo in $standardRepos) {
            $isScaleOut = ConvertTo-DRNullableBoolean -Value (Get-DRPropertyPathValue -Object $repo -Path 'IsScaleOut')
            if ($isScaleOut) { continue }

            $key = Get-DRRepositoryKey -Repository $repo
            # Skip repos attached to a SOBR as physical performance extents.
            if (-not [string]::IsNullOrWhiteSpace($key) -and $attachedPhysRepoKeys.Contains($key)) { continue }
            # Skip object-storage (COS/S3) repositories — they must not appear as physical/standalone.
            # This is the critical fix: Get-VBRBackupRepository returns COS repos too, but they should
            # only appear in the Object storage section, not as a second standalone Performance entry.
            if (-not [string]::IsNullOrWhiteSpace($key) -and $objectRepoKeys.Contains($key)) { continue }

            $standardOrder++
            $name   = Get-DRRepositoryName   -Repository $repo
            $status = Get-DRRepositoryStatus -Repository $repo
            $space  = Get-DRPhysicalRepositorySpace -Repository $repo
            [void]$rows.Add((New-DRRepositoryRow -Repository $name -Tier 'Performance' -Parent 'Standalone' -Status $status `
                -TotalBytes $space.TotalBytes -UsedBytes $space.UsedBytes -FreeBytes $space.FreeBytes -SortGroup 4 -SortOrder $standardOrder))
        }
    }

    # Process object-storage repositories using the list already collected above.
    # Skip only those that are already represented as SOBR extents (performance or capacity).
    $objectOrder = 0
    foreach ($objRepo in $objectRepoList) {
        $key = Get-DRRepositoryKey -Repository $objRepo
        if (-not [string]::IsNullOrWhiteSpace($key) -and $attachedObjRepoKeys.Contains($key)) { continue }

        $objectOrder++
        $name   = Get-DRRepositoryName   -Repository $objRepo
        $status = Get-DRRepositoryStatus -Repository $objRepo
        $space  = Get-DRObjectRepositorySpace -Repository $objRepo
        [void]$rows.Add((New-DRRepositoryRow -Repository $name -Tier 'Object storage' -Parent 'Standalone' -Status $status `
            -TotalBytes $space.TotalBytes -UsedBytes $space.UsedBytes -FreeBytes $space.FreeBytes -SortGroup 5 -SortOrder $objectOrder))
    }

    return @($rows | Sort-Object -Property SortGroup, SortOrder, Repository)
}

function New-DefinedRepositorySectionText {
    [CmdletBinding()]
    param()

    if ($Json) { return '' }

    try {
        Write-ProgressMessage 'Defined Repository — collecting repository utilisation...'
        Write-DebugMessage '[New-DefinedRepositorySectionText] Starting Defined Repository collection.'

        $rows = @(Get-DefinedRepositoryReport)
        Write-ProgressMessage ('Defined Repository — {0} repository row(s) found.' -f $rows.Count)
        Write-DebugMessage ('[New-DefinedRepositorySectionText] Collection complete: {0} row(s).' -f $rows.Count)

        $lines = New-Object 'System.Collections.Generic.List[string]'
        [void]$lines.Add('############### Defined Repository BEGIN ###################')

        if ($rows.Count -eq 0) {
            [void]$lines.Add('(no repositories found)')
        } else {
            $wRepository = 30
            $wTier       = 16
            $wParent     = 20
            $wStatus     = 12
            $wTotal      = 11
            $wUsed       = 11
            $wFree       = 11
            $wUsedPct    = 7
            $separatorWidth = $wRepository + $wTier + $wParent + $wStatus + $wTotal + $wUsed + $wFree + $wUsedPct + 7

            [void]$lines.Add(('{0} {1} {2} {3} {4} {5} {6} {7}' -f `
                (Format-DRFixedWidth -Value 'Repository' -Width $wRepository),
                (Format-DRFixedWidth -Value 'Tier'       -Width $wTier),
                (Format-DRFixedWidth -Value 'Parent'     -Width $wParent),
                (Format-DRFixedWidth -Value 'Status'     -Width $wStatus),
                (Format-DRFixedWidth -Value 'Total'      -Width $wTotal),
                (Format-DRFixedWidth -Value 'Used'       -Width $wUsed),
                (Format-DRFixedWidth -Value 'Free'       -Width $wFree),
                (Format-DRFixedWidth -Value 'Used %'     -Width $wUsedPct)))
            [void]$lines.Add(('-' * $separatorWidth))

            foreach ($row in $rows) {
                [void]$lines.Add(('{0} {1} {2} {3} {4} {5} {6} {7}' -f `
                    (Format-DRFixedWidth -Value ([string]$row.Repository) -Width $wRepository),
                    (Format-DRFixedWidth -Value ([string]$row.Tier)       -Width $wTier),
                    (Format-DRFixedWidth -Value ([string]$row.Parent)     -Width $wParent),
                    (Format-DRFixedWidth -Value ([string]$row.Status)     -Width $wStatus),
                    (Format-DRFixedWidth -Value ([string]$row.Total)      -Width $wTotal),
                    (Format-DRFixedWidth -Value ([string]$row.Used)       -Width $wUsed),
                    (Format-DRFixedWidth -Value ([string]$row.Free)       -Width $wFree),
                    (Format-DRFixedWidth -Value ([string]$row.'Used %')   -Width $wUsedPct)))
            }
        }

        [void]$lines.Add('############### Defined Repository END ###################')
        return ($lines -join [Environment]::NewLine)
    } catch {
        Write-Warning ('Defined Repository baseline failed to build: {0}' -f $_.Exception.Message)
        Write-DebugMessage ('[New-DefinedRepositorySectionText] Failed:' + [Environment]::NewLine + (Format-ErrorRecord -ErrorRecord $_))
        return (New-DefinedRepositoryPlaceholderSection -Message '(repository utilisation unavailable)')
    }
}

# ===========================================================================
# VBR Licensing baseline — helper functions
#   These functions are licensing-specific and DL-prefixed to avoid collisions
#   with existing helpers from other collector phases.
# ===========================================================================

function Format-DLVBRLicenseDate {
    [CmdletBinding()]
    param([object]$Date)

    if ($null -eq $Date) { return 'N/A' }
    try { return ([datetime]$Date).ToString('yyyy-MM-dd') } catch { return [string]$Date }
}

function Get-DLDaysRemaining {
    [CmdletBinding()]
    param([object]$Date)

    if ($null -eq $Date) { return 'N/A' }
    try {
        $d = [datetime]$Date
        return [int]($d - (Get-Date)).TotalDays
    } catch { return 'N/A' }
}

function ConvertTo-DLDisplayText {
    [CmdletBinding()]
    param([object]$Value)

    if ($null -eq $Value) { return '' }
    return [string]$Value
}

# ---------------------------------------------------------------------------
# New-VBRLicensingSectionText
#   Builds and returns the VBR Licensing baseline block as a single string.
#   The block is delimited with the required markers and contains subsections
#   for installed license, instance usage, socket usage and workloads.
#   Returns an error line inside the delimiters on failure; never throws.
#   Always returns empty string in JSON mode.
# ---------------------------------------------------------------------------
function New-VBRLicensingSectionText {
    [CmdletBinding()]
    param()

    if ($Json) { return '' }

    $lines = New-Object 'System.Collections.Generic.List[string]'
    [void]$lines.Add('############### VBR Licensing BEGIN ###################')

    try {
        Write-ProgressMessage 'Phase 8 — Collecting VBR licensing information...'
        Write-DebugMessage '[New-VBRLicensingSectionText] Starting VBR licensing collection.'

        if (-not (Get-Command -Name 'Get-VBRInstalledLicense' -ErrorAction SilentlyContinue)) {
            [void]$lines.Add('(Get-VBRInstalledLicense cmdlet not available)')
            [void]$lines.Add('############### VBR Licensing END ###################')
            return ($lines -join [Environment]::NewLine)
        }

        $license = Get-VBRInstalledLicense -ErrorAction Stop

        if ($null -eq $license) {
            [void]$lines.Add('(no VBR license found)')
            [void]$lines.Add('############### VBR Licensing END ###################')
            return ($lines -join [Environment]::NewLine)
        }

        # ─── INSTALLED VBR LICENCE ───────────────────────────────────────
        [void]$lines.Add('')
        [void]$lines.Add('INSTALLED VBR LICENCE')
        [void]$lines.Add('')

        $licenseSummary = [PSCustomObject]@{
            Status         = ConvertTo-DLDisplayText (Get-DRPropertyPathValue -Object $license -Path 'Status')
            Edition        = ConvertTo-DLDisplayText (Get-DRPropertyPathValue -Object $license -Path 'Edition')
            Type           = ConvertTo-DLDisplayText (Get-DRPropertyPathValue -Object $license -Path 'Type')
            LicensedTo     = ConvertTo-DLDisplayText (Get-DRPropertyPathValue -Object $license -Path 'LicensedTo')
            Expires        = Format-DLVBRLicenseDate (Get-DRPropertyPathValue -Object $license -Path 'ExpirationDate')
            DaysRemaining  = Get-DLDaysRemaining     (Get-DRPropertyPathValue -Object $license -Path 'ExpirationDate')
            SupportExpires = Format-DLVBRLicenseDate (Get-DRPropertyPathValue -Object $license -Path 'SupportExpirationDate')
            AutoUpdate     = ConvertTo-DLDisplayText (Get-DRPropertyPathValue -Object $license -Path 'AutoUpdate')
        }

        $tableText = $licenseSummary | Format-Table -AutoSize | Out-String
        foreach ($tl in ($tableText -split '\r?\n')) {
            [void]$lines.Add($tl)
        }

        $licenseExtra = [PSCustomObject]@{
            SupportID    = ConvertTo-DLDisplayText (Get-DRPropertyPathValue -Object $license -Path 'SupportId')
            FreeAgentUse = ConvertTo-DLDisplayText (Get-DRPropertyPathValue -Object $license -Path 'FreeAgentUse')
            CloudConnect = ConvertTo-DLDisplayText (Get-DRPropertyPathValue -Object $license -Path 'CloudConnect')
        }

        $extraText = $licenseExtra | Format-Table -AutoSize | Out-String
        foreach ($el in ($extraText -split '\r?\n')) {
            [void]$lines.Add($el)
        }

        # ─── INSTANCE LICENCE USAGE ─────────────────────────────────────
        $instanceSummary = $null
        try {
            if (Get-Command -Name 'Get-VBRInstanceLicenseSummary' -ErrorAction SilentlyContinue) {
                $instanceSummary = Get-VBRInstanceLicenseSummary -License $license -ErrorAction Stop
            }
            else {
                $instanceSummary = Get-DRPropertyPathValue -Object $license -Path 'InstanceLicenseSummary'
            }
        }
        catch {
            Write-DebugMessage ('[New-VBRLicensingSectionText] Get-VBRInstanceLicenseSummary failed: {0}' -f $_.Exception.Message)
            $instanceSummary = Get-DRPropertyPathValue -Object $license -Path 'InstanceLicenseSummary'
        }

        if ($null -ne $instanceSummary) {
            [void]$lines.Add('')
            [void]$lines.Add('INSTANCE LICENCE USAGE')
            [void]$lines.Add('')

            $instanceReport = foreach ($Summary in @($instanceSummary)) {
                $Licensed = Get-DRPropertyPathValue -Object $Summary -Path 'LicensedInstancesNumber'
                $Used     = Get-DRPropertyPathValue -Object $Summary -Path 'UsedInstancesNumber'

                $Remaining = if ($null -ne $Licensed -and $null -ne $Used) {
                    try { [math]::Max([int64]0, [int64]$Licensed - [int64]$Used) }
                    catch { $null }
                }
                else {
                    $null
                }

                $WorkloadProp = Get-DRPropertyPathValue -Object $Summary -Path 'Workload'

                [PSCustomObject]@{
                    Licensed  = $Licensed
                    Used      = $Used
                    Remaining = $Remaining
                    New       = Get-DRPropertyPathValue -Object $Summary -Path 'NewInstancesNumber'
                    Rental    = Get-DRPropertyPathValue -Object $Summary -Path 'RentalInstancesNumber'
                    Workloads = if ($WorkloadProp) { @($WorkloadProp).Count } else { 0 }
                }
            }

            $instanceText = $instanceReport |
                Format-Table Licensed, Used, Remaining, New, Rental, Workloads -AutoSize |
                Out-String
            foreach ($il in ($instanceText -split '\r?\n')) {
                [void]$lines.Add($il)
            }
        }

        # ─── SOCKET LICENCE USAGE ───────────────────────────────────────
        $socketSummary = $null
        if (Get-Command -Name 'Get-VBRSocketLicenseSummary' -ErrorAction SilentlyContinue) {
            try {
                $socketSummary = Get-VBRSocketLicenseSummary -License $license -ErrorAction Stop
            } catch {
                Write-DebugMessage ('[New-VBRLicensingSectionText] Get-VBRSocketLicenseSummary failed: {0}' -f $_.Exception.Message)
            }
        }
        if ($null -eq $socketSummary) {
            $socketSummary = Get-DRPropertyPathValue -Object $license -Path 'SocketLicenseSummary'
        }

        if ($null -ne $socketSummary) {
            [void]$lines.Add('')
            [void]$lines.Add('SOCKET LICENCE USAGE')
            [void]$lines.Add('')

            $socketItems = @($socketSummary)
            $socketRows = foreach ($s in $socketItems) {
                [PSCustomObject]@{
                    Platform  = ConvertTo-DLDisplayText (Get-DRPropertyPathValue -Object $s -Path 'Platform')
                    Licensed  = ConvertTo-DLDisplayText (Get-DRPropertyPathValue -Object $s -Path 'LicensedSocketsNumber')
                    Used      = ConvertTo-DLDisplayText (Get-DRPropertyPathValue -Object $s -Path 'UsedSocketsNumber')
                    Remaining = ConvertTo-DLDisplayText (Get-DRPropertyPathValue -Object $s -Path 'RemainingSocketsNumber')
                }
            }

            $socketText = $socketRows | Format-Table -AutoSize | Out-String
            foreach ($sl in ($socketText -split '\r?\n')) {
                [void]$lines.Add($sl)
            }
        }

        # ─── LICENSED INSTANCE WORKLOADS ────────────────────────────────
        if (Get-Command -Name 'Get-VBRLicensedInstanceWorkload' -ErrorAction SilentlyContinue) {
            try {
                $workloads = @(Get-VBRLicensedInstanceWorkload -License $license -ErrorAction Stop)
                if ($workloads.Count -gt 0) {
                    [void]$lines.Add('')
                    [void]$lines.Add('LICENSED INSTANCE WORKLOADS')
                    [void]$lines.Add('')

                    $workloadRows = foreach ($w in $workloads) {
                        $wObj = New-Object PSObject
                        foreach ($prop in @('Name','Workload','Type','Platform','HostName','ObjectName','LicenseType','InstancesNumber','InstanceCount','LastProcessingTime')) {
                            $val = Get-DRPropertyPathValue -Object $w -Path $prop
                            Add-Member -InputObject $wObj -MemberType NoteProperty -Name $prop -Value (ConvertTo-DLDisplayText $val) -Force
                        }
                        $wObj
                    }

                    $workloadText = $workloadRows | Format-Table -AutoSize | Out-String
                    foreach ($wl in ($workloadText -split '\r?\n')) {
                        [void]$lines.Add($wl)
                    }
                }
            } catch {
                Write-DebugMessage ('[New-VBRLicensingSectionText] Get-VBRLicensedInstanceWorkload failed: {0}' -f $_.Exception.Message)
            }
        }

        Write-DebugMessage '[New-VBRLicensingSectionText] VBR licensing collection complete.'

    } catch {
        Write-Warning ('VBR Licensing collection failed: {0}' -f $_.Exception.Message)
        Write-DebugMessage ('[New-VBRLicensingSectionText] Failed:' + [Environment]::NewLine + (Format-ErrorRecord -ErrorRecord $_))
        [void]$lines.Add(('(VBR licensing information unavailable: {0})' -f $_.Exception.Message))
    }

    [void]$lines.Add('############### VBR Licensing END ###################')
    return ($lines -join [Environment]::NewLine)
}

# ===========================================================================
# Backup Versions baseline — helper functions
#   BV-prefixed to avoid collisions with other collector phase helpers.
# ===========================================================================

function ConvertTo-BVIdKey {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [object]$Id
    )

    if ($null -eq $Id) {
        return $null
    }

    $Text = ([string]$Id).Trim().Trim('{}')

    if ([string]::IsNullOrWhiteSpace($Text)) {
        return $null
    }

    return $Text.ToLowerInvariant()
}

function Get-BVObjectProperty {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [object]$InputObject,

        [Parameter(Mandatory)]
        [string[]]$Names
    )

    if ($null -eq $InputObject) {
        return $null
    }

    foreach ($Name in $Names) {
        $Property = $InputObject.PSObject.Properties[$Name]

        if (
            $null -ne $Property -and
            $null -ne $Property.Value
        ) {
            return $Property.Value
        }
    }

    return $null
}

function Get-BVRepositoryId {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [object]$Repository
    )

    return ConvertTo-BVIdKey (
        Get-BVObjectProperty `
            -InputObject $Repository `
            -Names @(
                'Id'
                'RepositoryId'
            )
    )
}

function Get-BVGroupKey {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Machine,

        [Parameter(Mandatory)]
        [string]$Parent,

        [Parameter(Mandatory)]
        [string]$Repository
    )

    return (
        $Machine.ToLowerInvariant() +
        [char]0 +
        $Parent.ToLowerInvariant() +
        [char]0 +
        $Repository.ToLowerInvariant()
    )
}

function Get-BVOrCreateCountGroup {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Table,

        [Parameter(Mandatory)]
        [string]$Machine,

        [Parameter(Mandatory)]
        [string]$Parent,

        [Parameter(Mandatory)]
        [string]$Repository,

        [Parameter(Mandatory)]
        [string]$Tier
    )

    $Key = Get-BVGroupKey `
        -Machine $Machine `
        -Parent $Parent `
        -Repository $Repository

    if (-not $Table.ContainsKey($Key)) {
        $Table[$Key] = [PSCustomObject]@{
            Machine    = $Machine
            Parent     = $Parent
            Repository = $Repository
            Tier       = $Tier
            Versions   = [int64]0
        }
    }

    return $Table[$Key]
}

function Get-BVOrCreateRestorePointSet {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Table,

        [Parameter(Mandatory)]
        [string]$Machine,

        [Parameter(Mandatory)]
        [string]$Parent,

        [Parameter(Mandatory)]
        [string]$Repository,

        [Parameter(Mandatory)]
        [string]$Tier
    )

    $Key = Get-BVGroupKey `
        -Machine $Machine `
        -Parent $Parent `
        -Repository $Repository

    if (-not $Table.ContainsKey($Key)) {
        $Table[$Key] = [PSCustomObject]@{
            Machine     = $Machine
            Parent      = $Parent
            Repository  = $Repository
            Tier        = $Tier
            RestoreIds  = [System.Collections.Generic.HashSet[string]]::new(
                [System.StringComparer]::OrdinalIgnoreCase
            )
        }
    }

    return $Table[$Key]
}

# ---------------------------------------------------------------------------
# New-BackupVersionsSectionText
#   Builds and returns the Backup Versions baseline block as a single string.
#   Reports the number of backup versions per machine in each repository.
#   Returns an error line inside the delimiters on failure; never throws.
#   Always returns empty string in JSON mode.
# ---------------------------------------------------------------------------
function New-BackupVersionsSectionText {
    [CmdletBinding()]
    param()

    if ($Json) { return '' }

    $lines = New-Object 'System.Collections.Generic.List[string]'
    [void]$lines.Add('############### Backup Versions BEGIN ###################')

    try {
        Write-ProgressMessage 'Phase 9 — Collecting backup version counts...'
        Write-DebugMessage '[New-BackupVersionsSectionText] Starting backup version collection.'

        $RepositoryNameById = @{}

        $SobrByRepositoryId = @{}

        $SobrByName = @{}


        $StandardRepositories = @(
            Get-VBRBackupRepository `
                -ErrorAction SilentlyContinue
        )

        foreach ($Repository in $StandardRepositories) {
            $RepositoryId = Get-BVRepositoryId -Repository $Repository

            if ($RepositoryId) {
                $RepositoryNameById[$RepositoryId] = [string]$Repository.Name
            }
        }


        $ObjectRepositories = @()

        if (
            Get-Command Get-VBRObjectStorageRepository `
                -ErrorAction SilentlyContinue
        ) {
            $ObjectRepositories = @(
                Get-VBRObjectStorageRepository `
                    -ErrorAction SilentlyContinue
            )

            foreach ($Repository in $ObjectRepositories) {
                $RepositoryId = Get-BVRepositoryId -Repository $Repository

                if ($RepositoryId) {
                    $RepositoryNameById[$RepositoryId] = [string]$Repository.Name
                }
            }
        }


        $ScaleOutRepositories = @(
            Get-VBRBackupRepository `
                -ScaleOut `
                -ErrorAction SilentlyContinue
        )

        foreach ($SOBR in $ScaleOutRepositories) {
            $SobrId = Get-BVRepositoryId -Repository $SOBR

            $PerformanceExtents = @()

            try {
                $PerformanceExtents = @(
                    Get-VBRRepositoryExtent `
                        -Repository $SOBR `
                        -ErrorAction Stop
                )
            }
            catch {
                $PerformanceExtents = @()
            }

            $CapacityExtents = @()

            try {
                $CapacityExtents = @(
                    Get-VBRCapacityExtent `
                        -Repository $SOBR `
                        -ErrorAction Stop
                )
            }
            catch {
                $CapacityExtents = @()
            }

            $PerformanceNames = @(
                $PerformanceExtents |
                    ForEach-Object {
                        [string]$_.Repository.Name
                    } |
                    Where-Object {
                        -not [string]::IsNullOrWhiteSpace($_)
                    }
            )

            $CapacityNames = @(
                $CapacityExtents |
                    ForEach-Object {
                        [string]$_.Repository.Name
                    } |
                    Where-Object {
                        -not [string]::IsNullOrWhiteSpace($_)
                    }
            )

            $SobrInformation = [PSCustomObject]@{
                Name               = [string]$SOBR.Name
                Id                 = $SobrId
                PerformanceExtents = $PerformanceExtents
                CapacityExtents    = $CapacityExtents
                PerformanceNames   = $PerformanceNames
                CapacityNames      = $CapacityNames
            }

            $SobrByName[$SOBR.Name.ToLowerInvariant()] = $SobrInformation

            if ($SobrId) {
                $SobrByRepositoryId[$SobrId] = $SobrInformation
                $RepositoryNameById[$SobrId] = [string]$SOBR.Name
            }

            foreach ($Extent in $PerformanceExtents) {
                $ExtentRepository = $Extent.Repository
                $ExtentId = Get-BVRepositoryId -Repository $ExtentRepository

                if ($ExtentId) {
                    $SobrByRepositoryId[$ExtentId] = $SobrInformation
                    $RepositoryNameById[$ExtentId] =
                        [string]$ExtentRepository.Name
                }
            }

            foreach ($Extent in $CapacityExtents) {
                $ExtentRepository = $Extent.Repository
                $ExtentId = Get-BVRepositoryId -Repository $ExtentRepository

                if ($ExtentId) {
                    $SobrByRepositoryId[$ExtentId] = $SobrInformation
                    $RepositoryNameById[$ExtentId] =
                        [string]$ExtentRepository.Name
                }
            }
        }


        $PerformanceGroups = @{}

        $MachineByObjectAndSobr = @{}
        $MachineByObjectId      = @{}


        $Backups = @(
            Get-VBRBackup `
                -ErrorAction SilentlyContinue
        )

        foreach ($Backup in $Backups) {
            $BackupObjects = @()

            try {
                $BackupObjects = @(
                    Get-VBRBackupObject `
                        -Backup $Backup `
                        -ErrorAction Stop
                )
            }
            catch {
                Write-Warning (
                    "Unable to read objects from backup '$($Backup.Name)': " +
                    $_.Exception.Message
                )

                continue
            }

            foreach ($BackupObject in $BackupObjects) {
                $Machine = [string]$BackupObject.Name

                if ([string]::IsNullOrWhiteSpace($Machine)) {
                    continue
                }

                [int64]$VersionCount = 0

                if ($null -ne $BackupObject.RestorePointsCount) {
                    $VersionCount = [int64]$BackupObject.RestorePointsCount
                }

                $RepositoryId = ConvertTo-BVIdKey $BackupObject.RepositoryId

                if (-not $RepositoryId) {
                    $RepositoryId = ConvertTo-BVIdKey (
                        Get-BVObjectProperty `
                            -InputObject $Backup `
                            -Names @(
                                'RepositoryId'
                                'TargetRepositoryId'
                            )
                    )
                }

                $SobrInformation = $null

                if (
                    $RepositoryId -and
                    $SobrByRepositoryId.ContainsKey($RepositoryId)
                ) {
                    $SobrInformation = $SobrByRepositoryId[$RepositoryId]
                }

                if ($SobrInformation) {
                    $SobrName = $SobrInformation.Name

                    if (
                        $RepositoryId -and
                        $RepositoryNameById.ContainsKey($RepositoryId) -and
                        $RepositoryId -ne $SobrInformation.Id
                    ) {
                        $PerformanceRepository =
                            $RepositoryNameById[$RepositoryId]
                    }
                    elseif ($SobrInformation.PerformanceNames.Count -eq 1) {
                        $PerformanceRepository =
                            $SobrInformation.PerformanceNames[0]
                    }
                    elseif ($SobrInformation.PerformanceNames.Count -gt 1) {
                        $PerformanceRepository =
                            "$SobrName performance tier"
                    }
                    else {
                        $PerformanceRepository =
                            "$SobrName performance tier"
                    }

                    $Group = Get-BVOrCreateCountGroup `
                        -Table $PerformanceGroups `
                        -Machine $Machine `
                        -Parent $SobrName `
                        -Repository $PerformanceRepository `
                        -Tier 'Performance'

                    $Group.Versions += $VersionCount

                    foreach ($ObjectIdentifier in @(
                        $BackupObject.ObjectId
                        $BackupObject.Id
                    )) {
                        $ObjectKey = ConvertTo-BVIdKey $ObjectIdentifier

                        if (-not $ObjectKey) {
                            continue
                        }

                        $CombinedObjectKey = (
                            $ObjectKey +
                            [char]0 +
                            $SobrName.ToLowerInvariant()
                        )

                        $MachineByObjectAndSobr[$CombinedObjectKey] = $Machine
                        $MachineByObjectId[$ObjectKey] = $Machine
                    }
                }
                else {
                    $RepositoryName = $null

                    if (
                        $RepositoryId -and
                        $RepositoryNameById.ContainsKey($RepositoryId)
                    ) {
                        $RepositoryName =
                            $RepositoryNameById[$RepositoryId]
                    }

                    if ([string]::IsNullOrWhiteSpace($RepositoryName)) {
                        try {
                            if (
                                $Backup.PSObject.Methods.Name -contains
                                'GetRepository'
                            ) {
                                $BackupRepository = $Backup.GetRepository()
                                $RepositoryName = [string]$BackupRepository.Name
                            }
                        }
                        catch {
                        }
                    }

                    if ([string]::IsNullOrWhiteSpace($RepositoryName)) {
                        $RepositoryName = 'Unknown repository'
                    }

                    $Group = Get-BVOrCreateCountGroup `
                        -Table $PerformanceGroups `
                        -Machine $Machine `
                        -Parent 'Standalone' `
                        -Repository $RepositoryName `
                        -Tier 'Performance'

                    $Group.Versions += $VersionCount
                }
            }
        }


        $CapacityGroups = @{}

        if (
            (Get-Command Get-VBRSOBRObjectStorageBackup `
                -ErrorAction SilentlyContinue) -and
            (Get-Command Get-VBRSOBRObjectStorageRestorePoint `
                -ErrorAction SilentlyContinue)
        ) {
            $CapacityBackups = @()

            try {
                $CapacityBackups = @(
                    Get-VBRSOBRObjectStorageBackup `
                        -CapacityTier `
                        -ErrorAction Stop
                )
            }
            catch {
                Write-Warning (
                    'Unable to retrieve capacity-tier backups: ' +
                    $_.Exception.Message
                )

                $CapacityBackups = @()
            }

            foreach ($CapacityBackup in $CapacityBackups) {
                $RepositoryId = ConvertTo-BVIdKey `
                    $CapacityBackup.RepositoryId

                $SobrInformation = $null

                if (
                    $RepositoryId -and
                    $SobrByRepositoryId.ContainsKey($RepositoryId)
                ) {
                    $SobrInformation =
                        $SobrByRepositoryId[$RepositoryId]
                }

                if (-not $SobrInformation) {
                    Write-Warning (
                        "Unable to identify the parent SOBR for capacity backup " +
                        "'$($CapacityBackup.Name)' with repository ID " +
                        "'$RepositoryId'."
                    )

                    continue
                }

                $SobrName = $SobrInformation.Name

                if (
                    $RepositoryId -and
                    $RepositoryNameById.ContainsKey($RepositoryId) -and
                    $RepositoryId -ne $SobrInformation.Id
                ) {
                    $CapacityRepository =
                        $RepositoryNameById[$RepositoryId]
                }
                elseif ($SobrInformation.CapacityNames.Count -eq 1) {
                    $CapacityRepository =
                        $SobrInformation.CapacityNames[0]
                }
                elseif ($SobrInformation.CapacityNames.Count -gt 1) {
                    $CapacityRepository =
                        "$SobrName capacity tier"
                }
                else {
                    $CapacityRepository =
                        "$SobrName capacity tier"
                }

                $CapacityRestorePoints = @()

                try {
                    $CapacityRestorePoints = @(
                        Get-VBRSOBRObjectStorageRestorePoint `
                            -Backup $CapacityBackup `
                            -ErrorAction Stop |
                            Where-Object {
                                $_.IsCapacity -ne $false
                            }
                    )
                }
                catch {
                    Write-Warning (
                        "Unable to read capacity-tier restore points from " +
                        "'$($CapacityBackup.Name)': " +
                        $_.Exception.Message
                    )

                    continue
                }

                $FallbackCounter = 0

                foreach ($RestorePoint in $CapacityRestorePoints) {
                    $ObjectId = ConvertTo-BVIdKey $RestorePoint.ObjectId

                    $Machine = $null

                    if ($ObjectId) {
                        $CombinedObjectKey = (
                            $ObjectId +
                            [char]0 +
                            $SobrName.ToLowerInvariant()
                        )

                        if (
                            $MachineByObjectAndSobr.ContainsKey(
                                $CombinedObjectKey
                            )
                        ) {
                            $Machine =
                                $MachineByObjectAndSobr[$CombinedObjectKey]
                        }
                        elseif ($MachineByObjectId.ContainsKey($ObjectId)) {
                            $Machine = $MachineByObjectId[$ObjectId]
                        }
                    }

                    if ([string]::IsNullOrWhiteSpace($Machine)) {
                        $Machine = if ($ObjectId) {
                            "Unknown machine [$ObjectId]"
                        }
                        else {
                            'Unknown machine'
                        }
                    }

                    $Group = Get-BVOrCreateRestorePointSet `
                        -Table $CapacityGroups `
                        -Machine $Machine `
                        -Parent $SobrName `
                        -Repository $CapacityRepository `
                        -Tier 'Capacity'

                    $RestorePointId = ConvertTo-BVIdKey `
                        $RestorePoint.BackupId

                    if (-not $RestorePointId) {
                        $FallbackCounter++
                        $RestorePointId = (
                            "$ObjectId-$FallbackCounter-" +
                            "$($CapacityBackup.CreationTime.Ticks)"
                        )
                    }

                    $null = $Group.RestoreIds.Add($RestorePointId)
                }
            }
        }


        $Report = [System.Collections.Generic.List[object]]::new()

        foreach ($Group in $PerformanceGroups.Values) {
            $SortOrder = if ($Group.Parent -eq 'Standalone') {
                1
            }
            else {
                1
            }

            $Report.Add(
                [PSCustomObject]@{
                    Machine    = $Group.Machine
                    Repository = $Group.Repository
                    Tier       = $Group.Tier
                    ParentSOBR = if ($Group.Parent -eq 'Standalone') {
                        '-'
                    }
                    else {
                        $Group.Parent
                    }
                    Versions   = [int64]$Group.Versions
                    SortOrder  = $SortOrder
                }
            )
        }


        foreach ($Group in $CapacityGroups.Values) {
            $Report.Add(
                [PSCustomObject]@{
                    Machine    = $Group.Machine
                    Repository = $Group.Repository
                    Tier       = 'Capacity'
                    ParentSOBR = $Group.Parent
                    Versions   = [int64]$Group.RestoreIds.Count
                    SortOrder  = 2
                }
            )
        }


        $SobrMachineGroups = @{}

        foreach ($Row in @($Report)) {
            if ($Row.ParentSOBR -eq '-') {
                continue
            }

            $Key = (
                $Row.Machine.ToLowerInvariant() +
                [char]0 +
                $Row.ParentSOBR.ToLowerInvariant()
            )

            if (-not $SobrMachineGroups.ContainsKey($Key)) {
                $SobrMachineGroups[$Key] = [PSCustomObject]@{
                    Machine  = $Row.Machine
                    SOBR     = $Row.ParentSOBR
                    Versions = [int64]0
                }
            }

            $SobrMachineGroups[$Key].Versions +=
                [int64]$Row.Versions
        }


        foreach ($Combined in $SobrMachineGroups.Values) {
            $Report.Add(
                [PSCustomObject]@{
                    Machine    = $Combined.Machine
                    Repository = $Combined.SOBR
                    Tier       = 'SOBR combined'
                    ParentSOBR = $Combined.SOBR
                    Versions   = [int64]$Combined.Versions
                    SortOrder  = 0
                }
            )
        }


        $SortedReport = @(
            $Report |
                Sort-Object Machine,
                            ParentSOBR,
                            SortOrder,
                            Repository
        )

        [void]$lines.Add('')

        $tableText = $SortedReport |
            Format-Table `
                @{Label='Machine';    Expression={$_.Machine};    Width=48},
                @{Label='Repository'; Expression={$_.Repository}; Width=28},
                @{Label='Tier';       Expression={$_.Tier};       Width=14},
                @{Label='Versions';   Expression={$_.Versions};   Width=8} `
                -AutoSize |
            Out-String

        foreach ($tl in ($tableText -split '\r?\n')) {
            [void]$lines.Add($tl)
        }

        Write-DebugMessage '[New-BackupVersionsSectionText] Backup version collection complete.'

    } catch {
        Write-Warning ('Backup Versions collection failed: {0}' -f $_.Exception.Message)
        Write-DebugMessage ('[New-BackupVersionsSectionText] Failed:' + [Environment]::NewLine + (Format-ErrorRecord -ErrorRecord $_))
        [void]$lines.Add(('(backup version information unavailable: {0})' -f $_.Exception.Message))
    }

    [void]$lines.Add('############### Backup Versions END ###################')
    return ($lines -join [Environment]::NewLine)
}

# ===========================================================================
# SOBR Offload section — helper functions
#   P10-prefixed to avoid collisions with existing collector functions.
#   These helpers are used by New-SobrOffloadSectionText.
# ===========================================================================

function Get-P10PropertyPathValue {
    param (
        [AllowNull()]
        [object]$InputObject,

        [Parameter(Mandatory)]
        [string[]]$Paths
    )

    if ($null -eq $InputObject) {
        return $null
    }

    foreach ($Path in $Paths) {
        $Value = $InputObject
        $Found = $true

        foreach ($Part in ($Path -split '\.')) {
            if ($null -eq $Value) {
                $Found = $false
                break
            }

            $Property = $Value.PSObject.Properties[$Part]

            if ($null -eq $Property) {
                $Found = $false
                break
            }

            $Value = $Property.Value
        }

        if ($Found -and $null -ne $Value) {
            return $Value
        }
    }

    return $null
}


function Get-P10FirstNumericValue {
    param (
        [AllowNull()]
        [object]$InputObject,

        [Parameter(Mandatory)]
        [string[]]$Paths
    )

    foreach ($Path in $Paths) {
        $Value = Get-P10PropertyPathValue `
            -InputObject $InputObject `
            -Paths @($Path)

        if ($null -eq $Value) {
            continue
        }

        # Handle size objects that expose a Bytes property.
        foreach ($ByteProperty in @(
            'Bytes'
            'InBytes'
            'Value'
        )) {
            $Property = $Value.PSObject.Properties[$ByteProperty]

            if (
                $null -ne $Property -and
                $null -ne $Property.Value
            ) {
                [double]$Number = 0.0

                if (
                    [double]::TryParse(
                        [string]$Property.Value,
                        [ref]$Number
                    )
                ) {
                    return [PSCustomObject]@{
                        Found = $true
                        Value = $Number
                        Path  = "$Path.$ByteProperty"
                    }
                }
            }
        }

        [double]$Number = 0.0

        if (
            [double]::TryParse(
                [string]$Value,
                [ref]$Number
            )
        ) {
            return [PSCustomObject]@{
                Found = $true
                Value = $Number
                Path  = $Path
            }
        }
    }

    return [PSCustomObject]@{
        Found = $false
        Value = $null
        Path  = $null
    }
}


function Format-P10ByteSize {
    param (
        [AllowNull()]
        [object]$Bytes
    )

    if ($null -eq $Bytes) {
        return 'N/A'
    }

    [double]$Value = [double]$Bytes

    if ($Value -ge [double]1PB) {
        return '{0:N2} PiB' -f ($Value / [double]1PB)
    }

    if ($Value -ge [double]1TB) {
        return '{0:N2} TiB' -f ($Value / [double]1TB)
    }

    if ($Value -ge [double]1GB) {
        return '{0:N2} GiB' -f ($Value / [double]1GB)
    }

    if ($Value -ge [double]1MB) {
        return '{0:N2} MiB' -f ($Value / [double]1MB)
    }

    if ($Value -ge [double]1KB) {
        return '{0:N2} KiB' -f ($Value / [double]1KB)
    }

    return '{0:N0} B' -f $Value
}


function Format-P10RunTime {
    param (
        [Parameter(Mandatory)]
        [timespan]$Duration
    )

    if ($Duration.Days -gt 0) {
        return '{0}d {1:00}:{2:00}:{3:00}' -f `
            $Duration.Days,
            $Duration.Hours,
            $Duration.Minutes,
            $Duration.Seconds
    }

    return '{0:00}:{1:00}:{2:00}' -f `
        [math]::Floor($Duration.TotalHours),
        $Duration.Minutes,
        $Duration.Seconds
}


function Get-P10SessionProgressPercent {
    param (
        [Parameter(Mandatory)]
        [object]$Session
    )

    $Result = Get-P10FirstNumericValue `
        -InputObject $Session `
        -Paths @(
            'Progress'
            'Progress.Percent'
            'Progress.Percents'
            'Info.Progress.Percent'
            'Info.Progress.Percents'
        )

    if (-not $Result.Found) {
        return $null
    }

    [int]$Percent = [math]::Round($Result.Value)

    if ($Percent -lt 0) {
        $Percent = 0
    }

    if ($Percent -gt 100) {
        $Percent = 100
    }

    return $Percent
}


function Get-P10TaskTransferInformation {
    param (
        [Parameter(Mandatory)]
        [object[]]$Tasks
    )

    [double]$TransferredTotal = 0.0
    [double]$ProcessedTotal   = 0.0

    $TransferredFound = $false
    $ProcessedFound   = $false

    foreach ($Task in $Tasks) {
        # Veeam has historically exposed the property as TransferedSize
        # with one "r". Check both spellings and both common object paths.
        $Transferred = Get-P10FirstNumericValue `
            -InputObject $Task `
            -Paths @(
                'Progress.TransferedSize'
                'Progress.TransferredSize'
                'Info.Progress.TransferedSize'
                'Info.Progress.TransferredSize'
                'TransferedSize'
                'TransferredSize'
            )

        if ($Transferred.Found) {
            $TransferredTotal += [double]$Transferred.Value
            $TransferredFound = $true
        }

        $Processed = Get-P10FirstNumericValue `
            -InputObject $Task `
            -Paths @(
                'Progress.ProcessedSize'
                'Info.Progress.ProcessedSize'
                'ProcessedSize'
            )

        if ($Processed.Found) {
            $ProcessedTotal += [double]$Processed.Value
            $ProcessedFound = $true
        }
    }

    if ($TransferredFound) {
        return [PSCustomObject]@{
            Bytes   = $TransferredTotal
            Measure = 'Transferred'
        }
    }

    if ($ProcessedFound) {
        return [PSCustomObject]@{
            Bytes   = $ProcessedTotal
            Measure = 'Processed fallback'
        }
    }

    return [PSCustomObject]@{
        Bytes   = $null
        Measure = 'Unavailable'
    }
}


# ---------------------------------------------------------------------------
# New-SobrOffloadSectionText
#   Builds and returns the SOBR Offloads baseline block as a single string.
#   Reports all SOBR archive/capacity-tier offload sessions within the time
#   window (running sessions are always included).
#   Returns an error line inside the delimiters on failure; never throws.
#   Always returns empty string in JSON mode.
# ---------------------------------------------------------------------------
function New-SobrOffloadSectionText {
    [CmdletBinding()]
    param()

    if ($Json) { return '' }

    $lines = New-Object 'System.Collections.Generic.List[string]'
    [void]$lines.Add('############### SOBR Offloads BEGIN ###################')

    try {
        Write-ProgressMessage 'SOBR Offloads — collecting archive/offload sessions...'
        Write-DebugMessage '[New-SobrOffloadSectionText] Starting SOBR offload collection.'

        if (-not (Get-Command -Name 'Get-VBRSession' -ErrorAction SilentlyContinue)) {
            [void]$lines.Add('(Get-VBRSession cmdlet not available; skipping SOBR offload section)')
            [void]$lines.Add('############### SOBR Offloads END ###################')
            return ($lines -join [Environment]::NewLine)
        }

        if (-not (Test-CmdletHasParameter -CmdletName 'Get-VBRSession' -ParameterName 'Type')) {
            [void]$lines.Add('(Get-VBRSession does not support -Type parameter; skipping SOBR offload section)')
            [void]$lines.Add('############### SOBR Offloads END ###################')
            return ($lines -join [Environment]::NewLine)
        }

        $sobrSessions = @(
            Get-VBRSession -Type ArchiveBackup -ErrorAction Stop |
                Where-Object { $null -ne $_ -and (Test-SessionInWindow -Session $_) } |
                Sort-Object -Property @(
                    @{ Expression = { [string]$_.Name }; Descending = $false }
                )
        )

        Write-ProgressMessage ('SOBR Offloads — {0} session(s) found in window.' -f $sobrSessions.Count)
        Write-DebugMessage ('[New-SobrOffloadSectionText] {0} session(s) in window.' -f $sobrSessions.Count)

        if ($sobrSessions.Count -eq 0) {
            [void]$lines.Add('No SOBR offloads found.')
            [void]$lines.Add('############### SOBR Offloads END ###################')
            return ($lines -join [Environment]::NewLine)
        }

        $reportRows = [System.Collections.Generic.List[object]]::new()

        foreach ($OriginalSession in $sobrSessions) {
            # Attempt to refresh session state; fall back to original on error.
            $Session = $OriginalSession
            try {
                $Refreshed = @(Get-VBRSession -Session $OriginalSession -ErrorAction Stop) |
                    Where-Object { $null -ne $_ } |
                    Select-Object -First 1
                if ($null -ne $Refreshed) {
                    $Session = $Refreshed
                }
            } catch {
                Write-DebugMessage ('[New-SobrOffloadSectionText] Session refresh failed for "{0}": {1}' -f [string]$OriginalSession.Name, $_.Exception.Message)
            }

            $StartTime = Get-P10PropertyPathValue `
                -InputObject $Session `
                -Paths @(
                    'CreationTime'
                    'StartTime'
                    'CreationTimeLocal'
                    'Info.CreationTime'
                )

            try {
                $StartTime = [datetime]$StartTime
            } catch {
                $StartTime = $null
            }

            $RunTime = if ($null -ne $StartTime) {
                (Get-Date) - $StartTime
            } else {
                [timespan]::Zero
            }

            $Tasks = @(
                Get-VBRTaskSession -Session $Session -ErrorAction SilentlyContinue |
                    Where-Object { $null -ne $_ }
            )

            $Transfer = Get-P10TaskTransferInformation -Tasks $Tasks

            # Some builds expose transfer information at session level instead.
            if ($null -eq $Transfer.Bytes) {
                $SessionTransfer = Get-P10FirstNumericValue `
                    -InputObject $Session `
                    -Paths @(
                        'Info.Progress.TransferedSize'
                        'Info.Progress.TransferredSize'
                        'Progress.TransferedSize'
                        'Progress.TransferredSize'
                    )

                if ($SessionTransfer.Found) {
                    $Transfer = [PSCustomObject]@{
                        Bytes   = [double]$SessionTransfer.Value
                        Measure = 'Session transferred'
                    }
                }
            }

            $Progress = Get-P10SessionProgressPercent -Session $Session
            $Errors = ''

            try {
                $Errors = [string](Get-LastErrorText -Session $Session)
                if ($null -eq $Errors) {
                    $Errors = ''
                } else {
                    $Errors = $Errors.Trim()
                }
            } catch {
                $Errors = ''
                Write-DebugMessage ('[New-SobrOffloadSectionText] Error text retrieval failed for "{0}": {1}' -f [string]$Session.Name, $_.Exception.Message)
                Write-DebugMessage ('[New-SobrOffloadSectionText] Error text retrieval details:' + [Environment]::NewLine + (Format-ErrorRecord -ErrorRecord $_))
            }

            Write-DebugMessage ('[New-SobrOffloadSectionText] Session "{0}" state="{1}" tasks={2} transfer.found={3} transfer.bytes="{4}" transfer.measure="{5}" progress="{6}"' -f `
                [string]$Session.Name,
                [string]$Session.State,
                $Tasks.Count,
                ($null -ne $Transfer.Bytes),
                $Transfer.Bytes,
                [string]$Transfer.Measure,
                $(if ($null -ne $Progress) { $Progress } else { '<null>' }))

            $reportRows.Add(
                [PSCustomObject]@{
                    SOBROffload = [string]$Session.Name
                    State       = [string]$Session.State
                    Started     = if ($null -ne $StartTime) {
                        $StartTime.ToString('dd/MM/yyyy HH:mm')
                    } else {
                        'Unknown'
                    }
                    Running     = Format-P10RunTime -Duration $RunTime
                    Progress    = if ($null -ne $Progress) { "$Progress%" } else { 'N/A' }
                    DataMoved   = Format-P10ByteSize -Bytes $Transfer.Bytes
                    Measure     = $Transfer.Measure
                    Tasks       = $Tasks.Count
                    Errors      = $Errors
                }
            )
        }

        if ($reportRows.Count -eq 0) {
            [void]$lines.Add('No SOBR offloads found.')
        } else {
            $tableText = $reportRows |
                Format-Table `
                    @{Label='SOBR offload'; Expression={$_.SOBROffload}; Width=36},
                    @{Label='State';        Expression={$_.State};       Width=17},
                    @{Label='Started';      Expression={$_.Started};     Width=16},
                    @{Label='Running';      Expression={$_.Running};     Width=13},
                    @{Label='Progress';     Expression={$_.Progress};    Width=8},
                    @{Label='Data moved';   Expression={$_.DataMoved};   Width=12},
                    @{Label='Measure';      Expression={$_.Measure};     Width=18},
                    @{Label='Tasks';        Expression={$_.Tasks};       Width=5} |
                Out-String

            foreach ($tl in ($tableText -split '\r?\n')) {
                [void]$lines.Add($tl)
            }

            $errorRows = @(
                $reportRows |
                    Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_.Errors) }
            )

            if ($errorRows.Count -gt 0) {
                [void]$lines.Add('Errors:')
                foreach ($errorRow in $errorRows) {
                    $errorText = ([string]$errorRow.Errors -replace '\r?\n', ' | ').Trim()
                    [void]$lines.Add(('  {0}: {1}' -f [string]$errorRow.SOBROffload, $errorText))
                }
            }
        }

        Write-DebugMessage '[New-SobrOffloadSectionText] SOBR offload section complete.'

    } catch {
        Write-Warning ('SOBR Offloads collection failed: {0}' -f $_.Exception.Message)
        Write-DebugMessage ('[New-SobrOffloadSectionText] Failed:' + [Environment]::NewLine + (Format-ErrorRecord -ErrorRecord $_))
        [void]$lines.Add(('Unable to collect SOBR offload sessions: {0}' -f $_.Exception.Message))
    }

    [void]$lines.Add('############### SOBR Offloads END ###################')
    return ($lines -join [Environment]::NewLine)
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
    $isRunning = Test-SessionIsRunning -Session $Session
    $runningFor = if ($isRunning) { Get-SessionRunningDuration -Session $Session } else { '' }
    $processedBytes = if ($isRunning) { Get-SessionProcessedBytes -Session $Session } else { $null }
    $dataProcessed = if ($isRunning -and $null -ne $processedBytes) { Format-DRByteSize -Bytes $processedBytes } else { '' }
    $lastErr = Get-LastErrorText    -Session $Session
    $warningDetails = Get-VeeamWarningDetails -Session $Session

    return [pscustomobject][ordered]@{
        job_name        = $name
        job_type        = $type
        result          = $result
        start_time      = if ($null -ne $start) { $start.ToString('o') } else { $null }
        end_time        = if ($null -ne $end)   { $end.ToString('o')   } else { $null }
        running_for     = $runningFor
        data_processed  = $dataProcessed
        last_error      = $lastErr
        warning_details = $warningDetails
        source          = $Source
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

    Write-DebugMessage ('[Add-JobReportFromJob] Job="{0}" Type="{1}" Source={2}' -f $jobName, $jobType, $Source)
    if ($script:CollectorDebugEnabled) {
        Write-DebugMessage ('[Add-JobReportFromJob] Job object summary:' + [Environment]::NewLine + (Format-VeeamObjectSummary -InputObject $Job))
    }

    # Collect candidate sessions from the job object.
    $candidates = New-Object 'System.Collections.Generic.List[object]'

    if ($Job.PSObject.Methods['FindLastSession']) {
        Write-DebugMessage ('[Add-JobReportFromJob] Calling $Job.FindLastSession() for "{0}"' -f $jobName)
        try {
            $s = $Job.FindLastSession()
            if ($null -ne $s) {
                Write-DebugMessage ('[Add-JobReportFromJob] FindLastSession() returned: {0}' -f (Get-SessionName -Session $s))
                [void]$candidates.Add($s)
            } else {
                Write-DebugMessage '[Add-JobReportFromJob] FindLastSession() returned null.'
            }
        } catch {
            Write-DebugMessage ('[Add-JobReportFromJob] $Job.FindLastSession() threw:' +
                [Environment]::NewLine + (Format-ErrorRecord -ErrorRecord $_))
        }
    }

    if ($Job.PSObject.Methods['FindLastSessions']) {
        Write-DebugMessage ('[Add-JobReportFromJob] Calling $Job.FindLastSessions() for "{0}"' -f $jobName)
        try {
            $found = @($Job.FindLastSessions())
            Write-DebugMessage ('[Add-JobReportFromJob] FindLastSessions() returned {0} session(s).' -f $found.Count)
            foreach ($s in $found) {
                if ($null -ne $s) { [void]$candidates.Add($s) }
            }
        } catch {
            Write-DebugMessage ('[Add-JobReportFromJob] $Job.FindLastSessions() threw:' +
                [Environment]::NewLine + (Format-ErrorRecord -ErrorRecord $_))
        }
    }

    if ($Job.PSObject.Methods['GetSessions']) {
        Write-DebugMessage ('[Add-JobReportFromJob] Calling $Job.GetSessions() for "{0}"' -f $jobName)
        try {
            $found = @($Job.GetSessions())
            Write-DebugMessage ('[Add-JobReportFromJob] GetSessions() returned {0} session(s).' -f $found.Count)
            foreach ($s in $found) {
                if ($null -ne $s) { [void]$candidates.Add($s) }
            }
        } catch {
            Write-DebugMessage ('[Add-JobReportFromJob] $Job.GetSessions() threw:' +
                [Environment]::NewLine + (Format-ErrorRecord -ErrorRecord $_))
        }
    }

    # Also try Get-VBRBackupSession filtered by job when available.
    if (Get-Command -Name 'Get-VBRBackupSession' -ErrorAction SilentlyContinue) {
        Write-DebugMessage ('[Add-JobReportFromJob] Calling Get-VBRBackupSession -Job "{0}"' -f $jobName)
        try {
            $bsSessions = @(Get-VBRBackupSession -Job $Job -ErrorAction Stop)
            Write-DebugMessage ('[Add-JobReportFromJob] Get-VBRBackupSession returned {0} session(s).' -f $bsSessions.Count)
            foreach ($s in $bsSessions) {
                if ($null -ne $s) { [void]$candidates.Add($s) }
            }
        } catch {
            Write-DebugMessage ('[Add-JobReportFromJob] Get-VBRBackupSession -Job "{0}" threw:' -f $jobName +
                [Environment]::NewLine + (Format-ErrorRecord -ErrorRecord $_))
        }
    }

    # Filter to window and pick the most recent.
    $inWindow = @($candidates | Where-Object { $null -ne $_ -and (Test-SessionInWindow -Session $_) })
    Write-DebugMessage ('[Add-JobReportFromJob] "{0}": {1} candidate(s), {2} in window.' -f $jobName, $candidates.Count, $inWindow.Count)
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

    Write-DebugMessage ('[Add-JobReportFromJob] Selected session for "{0}": {1}' -f $jobName, (Get-SessionName -Session $session))
    if ($script:CollectorDebugEnabled) {
        Write-DebugMessage ('[Add-JobReportFromJob] Session object summary:' + [Environment]::NewLine + (Format-VeeamObjectSummary -InputObject $session))
    }

    $sessionId = Get-ObjectIdentity -InputObject $session
    if (-not $script:SeenSessions.Add($sessionId)) {
        Write-DebugMessage ('[Add-JobReportFromJob] Session "{0}" already seen; skipping duplicate.' -f $sessionId)
        return
    }

    $report = Build-JobReport -Session $session -JobName $jobName -JobType $jobType -Source $Source
    Write-DebugMessage ('[Add-JobReportFromJob] Built report for "{0}": result={1}  lastError={2}' `
        -f $jobName, $report.result, $(if ([string]::IsNullOrWhiteSpace($report.last_error)) { '<none>' } else { $report.last_error }))
    [void]$Results.Value.Add($report)
}

# ---------------------------------------------------------------------------
# Get-CollectorHostName
# ---------------------------------------------------------------------------
function Get-CollectorHostName {
    [CmdletBinding()]
    param()

    if (-not [string]::IsNullOrWhiteSpace($env:COMPUTERNAME)) {
        return $env:COMPUTERNAME
    }

    return [System.Environment]::MachineName
}

# ---------------------------------------------------------------------------
# New-CollectorReportBody
#   Returns the single canonical human-readable report string used for console,
#   disk, and email output. Never includes progress/debug lines.
# ---------------------------------------------------------------------------
function New-CollectorReportBody {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [AllowEmptyCollection()] [object[]]$Reports,
        [Parameter(Mandatory)] [int]$TotalJobs,
        [Parameter(Mandatory)] [int]$FailedCount,
        [Parameter(Mandatory)] [int]$WarnCount,
        [Parameter(Mandatory)] [int]$SuccessCount,
        [Parameter(Mandatory)] [int]$WithError,
        # Optional Defined Jobs baseline block (text mode only; empty in JSON mode).
        [string]$DefinedJobsSection = '',
        # Optional Defined Repository baseline block (text mode only; empty in JSON mode).
        [string]$DefinedRepositorySection = '',
        # Optional VBR Licensing baseline block (text mode only; empty in JSON mode).
        [string]$LicensingSection = '',
        # Optional Backup Versions baseline block (text mode only; empty in JSON mode).
        [string]$BackupVersionsSection = '',
        # Optional SOBR Offloads section (text mode only; empty in JSON mode).
        [string]$SobrOffloadSection = ''
    )

    $lines = New-Object 'System.Collections.Generic.List[string]'
    $ed = if ($PSVersionTable.PSEdition) { $PSVersionTable.PSEdition } else { 'Desktop' }

    [void]$lines.Add('============================================================')
    [void]$lines.Add('Veeam Last-Error Report')
    [void]$lines.Add(('Window     : last {0} hour(s)  ({1:o} to {2:o})' -f $Hours, $script:StartTime, $script:EndTime))
    [void]$lines.Add(('Host       : {0}' -f (Get-CollectorHostName)))
    [void]$lines.Add(('PowerShell : {0} {1}' -f $ed, $PSVersionTable.PSVersion))
    [void]$lines.Add('============================================================')
    [void]$lines.Add('')

    # Defined baseline sections (text mode only — never populated in JSON mode)
    $hasBaselineSection = $false
    if (-not [string]::IsNullOrWhiteSpace($DefinedJobsSection)) {
        [void]$lines.Add($DefinedJobsSection)
        $hasBaselineSection = $true
    }
    if (-not [string]::IsNullOrWhiteSpace($DefinedRepositorySection)) {
        [void]$lines.Add($DefinedRepositorySection)
        $hasBaselineSection = $true
    }
    if (-not [string]::IsNullOrWhiteSpace($LicensingSection)) {
        [void]$lines.Add($LicensingSection)
        $hasBaselineSection = $true
    }
    if (-not [string]::IsNullOrWhiteSpace($BackupVersionsSection)) {
        [void]$lines.Add($BackupVersionsSection)
        $hasBaselineSection = $true
    }
    if (-not [string]::IsNullOrWhiteSpace($SobrOffloadSection)) {
        [void]$lines.Add($SobrOffloadSection)
        $hasBaselineSection = $true
    }
    if ($hasBaselineSection) {
        [void]$lines.Add('')
    }

    if ($Reports.Count -eq 0) {
        if ($OnlyFailures) {
            [void]$lines.Add('No Failed or Warning sessions found in the specified window.')
        } else {
            [void]$lines.Add('No sessions found in the specified window.')
        }
    } else {
        foreach ($r in $Reports) {
            [void]$lines.Add(('Job      : {0}' -f $r.job_name))
            [void]$lines.Add(('Type     : {0}' -f $r.job_type))
            [void]$lines.Add(('Result   : {0}' -f $r.result))
            [void]$lines.Add(('End Time : {0}' -f $(if ($null -ne $r.end_time) { $r.end_time } else { '(running/unknown)' })))

            $runningFor = Get-PropertyValue -InputObject $r -Names @('running_for')
            if (-not [string]::IsNullOrWhiteSpace([string]$runningFor)) {
                [void]$lines.Add(('Running  : {0}' -f $runningFor))
            }

            $dataProcessed = Get-PropertyValue -InputObject $r -Names @('data_processed')
            if (-not [string]::IsNullOrWhiteSpace([string]$dataProcessed)) {
                [void]$lines.Add(('Processed: {0}' -f $dataProcessed))
            }

            if (-not [string]::IsNullOrWhiteSpace([string]$r.last_error)) {
                [void]$lines.Add(('Error    : {0}' -f $r.last_error))
            }

            $warningDetails = Get-PropertyValue -InputObject $r -Names @('warning_details')
            if (-not [string]::IsNullOrWhiteSpace([string]$warningDetails)) {
                [void]$lines.Add(('Warning  : {0}' -f $warningDetails))
            }

            [void]$lines.Add('')
        }
    }

    [void]$lines.Add('------------------------------------------------------------')
    [void]$lines.Add(('Window   : last {0} hour(s)  ({1:o}  to  {2:o})' -f $Hours, $script:StartTime, $script:EndTime))
    [void]$lines.Add(('Jobs     : {0}  (Failed: {1}  Warning: {2}  Success: {3}  WithError: {4})' `
        -f $TotalJobs, $FailedCount, $WarnCount, $SuccessCount, $WithError))
    [void]$lines.Add('------------------------------------------------------------')

    return ($lines -join [Environment]::NewLine)
}

# ---------------------------------------------------------------------------
# Get-CollectorReportFilePath
# ---------------------------------------------------------------------------
function Get-CollectorReportFilePath {
    [CmdletBinding()]
    param()

    $timestamp = (Get-Date).ToString('yyyyMMdd_HHmmss')
    return [IO.Path]::Combine($ReportOutputDirectory, ('Veeam_Collector_Report_{0}.txt' -f $timestamp))
}

# ---------------------------------------------------------------------------
# Write-CollectorReportBodyToDisk
# ---------------------------------------------------------------------------
function Write-CollectorReportBodyToDisk {
    [CmdletBinding()]
    param([Parameter(Mandatory)] [string]$Body)

    try {
        if (-not (Test-Path -LiteralPath $ReportOutputDirectory)) {
            $null = New-Item -ItemType Directory -Path $ReportOutputDirectory -Force
        }

        $path = Get-CollectorReportFilePath
        [System.IO.File]::WriteAllText($path, $Body, [System.Text.Encoding]::UTF8)
        Write-ProgressMessage ('Report body written to: {0}' -f $path)
        return $path
    } catch {
        Write-Warning ('Unable to write report body to "{0}": {1}' -f $ReportOutputDirectory, $_.Exception.Message)
        Write-DebugMessage ('[Write-CollectorReportBodyToDisk] Failure:' + [Environment]::NewLine + (Format-ErrorRecord -ErrorRecord $_))
        return ''
    }
}

# ---------------------------------------------------------------------------
# Get-CollectorMailSubject
# ---------------------------------------------------------------------------
function Get-CollectorMailSubject {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [int]$FailedCount,
        [Parameter(Mandatory)] [int]$WarnCount
    )

    return ('Veeam Last-Error Report - {0} - Failed: {1} Warning: {2}' -f (Get-CollectorHostName), $FailedCount, $WarnCount)
}

# ---------------------------------------------------------------------------
# Send-CollectorReportEmail
# ---------------------------------------------------------------------------
function Send-CollectorReportEmail {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string]$Body,
        [Parameter(Mandatory)] [string]$Subject
    )

    if ($DisableEmail) {
        Write-ProgressMessage 'Email delivery disabled by -DisableEmail.'
        return $false
    }

    $recipients = @(
        foreach ($recipient in $MailTo) {
            if (-not [string]::IsNullOrWhiteSpace([string]$recipient)) {
                [string]$recipient
            }
        }
    )

    if ($recipients.Count -eq 0) {
        Write-Warning 'Email delivery skipped because no recipients were configured.'
        return $false
    }

    try {
        $mailMessage = New-Object 'System.Net.Mail.MailMessage'
        try {
            $mailMessage.From = $MailFrom
            foreach ($recipient in $recipients) {
                [void]$mailMessage.To.Add($recipient)
            }
            $mailMessage.Subject = $Subject
            $mailMessage.Body = $Body
            $mailMessage.IsBodyHtml = $false

            $smtpClient = New-Object 'System.Net.Mail.SmtpClient'($SmtpServer)
            try {
                $smtpClient.Send($mailMessage)
            } finally {
                $smtpClient.Dispose()
            }
        } finally {
            $mailMessage.Dispose()
        }

        Write-ProgressMessage ('Report email sent to: {0}' -f ($recipients -join ', '))
        return $true
    } catch {
        Write-Warning ('Unable to send report email via "{0}": {1}' -f $SmtpServer, $_.Exception.Message)
        Write-DebugMessage ('[Send-CollectorReportEmail] Failure:' + [Environment]::NewLine + (Format-ErrorRecord -ErrorRecord $_))
        return $false
    }
}

# ---------------------------------------------------------------------------
# Remove-OldCollectorFiles
# ---------------------------------------------------------------------------
function Remove-OldCollectorFiles {
    [CmdletBinding()]
    param()

    if (-not (Test-Path -LiteralPath $ReportOutputDirectory)) {
        Write-DebugMessage ('[Remove-OldCollectorFiles] Directory not found, skipping cleanup: {0}' -f $ReportOutputDirectory)
        return
    }

    $cutoff = (Get-Date).AddDays(-1 * $RetentionDays)
    $patterns = @(
        'Veeam_Collector_Report_*.txt',
        'Veeam_Collector_*.log',
        'veeam-collector-debug-*.log'
    )

    try {
        $staleFiles = @(Get-ChildItem -LiteralPath $ReportOutputDirectory -File -ErrorAction Stop | Where-Object {
            $file = $_
            if ($file.LastWriteTime -ge $cutoff) { return $false }

            foreach ($pattern in $patterns) {
                if ($file.Name -like $pattern) {
                    return $true
                }
            }

            return $false
        })

        foreach ($file in $staleFiles) {
            try {
                Remove-Item -LiteralPath $file.FullName -Force -ErrorAction Stop
                Write-ProgressMessage ('Removed old collector file: {0}' -f $file.FullName)
            } catch {
                Write-Warning ('Unable to remove old collector file "{0}": {1}' -f $file.FullName, $_.Exception.Message)
                Write-DebugMessage ('[Remove-OldCollectorFiles] Remove failure:' + [Environment]::NewLine + (Format-ErrorRecord -ErrorRecord $_))
            }
        }

        if ($staleFiles.Count -eq 0) {
            Write-ProgressMessage ('No collector report/log files older than {0} day(s) found in {1}.' -f $RetentionDays, $ReportOutputDirectory)
        }
    } catch {
        Write-Warning ('Unable to clean up old collector files in "{0}": {1}' -f $ReportOutputDirectory, $_.Exception.Message)
        Write-DebugMessage ('[Remove-OldCollectorFiles] Enumeration failure:' + [Environment]::NewLine + (Format-ErrorRecord -ErrorRecord $_))
    }
}

# ===========================================================================
# Main
# ===========================================================================

# Top-level fatal error trap — catches any terminating error that escapes the
# structured try/catch blocks below, emits a FATAL diagnostic record, and
# exits with a non-zero code so callers detect the failure.
trap {
    $fatalMsg = '[FATAL] Veeam_Collector.ps1 terminated with an unhandled error.'
    Write-Warning $fatalMsg
    # Use Write-DebugMessage for the detail record; it handles both Warning stream
    # and file append in one place.  The plain Write-Warning above always fires so
    # callers see the FATAL line even when -CollectorDebug is not set.
    Write-DebugMessage ('FATAL error detail:' + [Environment]::NewLine + (Format-ErrorRecord -ErrorRecord $_))
    $host.SetShouldExit(1)
    break
}

Write-ProgressMessage ('Veeam Last-Error Report starting. Window: last {0} hour(s) ({1:o} to {2:o}).' `
    -f $Hours, $script:StartTime, $script:EndTime)

Import-VeeamPowerShell
Write-EnvironmentDiagnostics

# ---------------------------------------------------------------------------
# Defined Jobs baseline — collected once immediately after Veeam PowerShell
# loads so it reflects the current job inventory.  Skipped in JSON mode to
# keep stdout a pure JSON array.
# ---------------------------------------------------------------------------
$definedJobsSection = ''
$definedRepositorySection = ''
$licensingSection = ''
$backupVersionsSection = ''
$sobrOffloadSection = ''
if (-not $Json) {
    Write-DebugMessage '[Main] Building Defined Jobs baseline section.'
    try {
        $definedJobsSection = New-DefinedJobsSectionText
        Write-DebugMessage ('[Main] Defined Jobs section ready, {0} char(s).' -f $definedJobsSection.Length)
    } catch {
        Write-Warning ('Defined Jobs baseline failed: {0}' -f $_.Exception.Message)
        Write-DebugMessage ('[Main] Defined Jobs baseline failed:' + [Environment]::NewLine + (Format-ErrorRecord -ErrorRecord $_))
        $definedJobsSection = ''
    }

}

$allReports = New-Object 'System.Collections.Generic.List[object]'

# Job objects collected from Phase 1/2 (retained for potential future use).
$sessionFallbackJobs = New-Object 'System.Collections.Generic.List[object]'

# ---------------------------------------------------------------------------
# Phase 1 — Regular VBR jobs via Get-VBRJob
# ---------------------------------------------------------------------------
Write-ProgressMessage 'Phase 1 — Enumerating regular jobs (Get-VBRJob).'
Write-DebugMessage '[Main] Phase 1 — Get-VBRJob'
if (Get-Command -Name 'Get-VBRJob' -ErrorAction SilentlyContinue) {
    try {
        Write-DebugMessage '[Main] Calling Get-VBRJob ...'
        $vbrJobs = @(Get-VBRJob -ErrorAction Stop -WarningAction SilentlyContinue)
        Write-ProgressMessage ('  Found {0} job(s) via Get-VBRJob.' -f $vbrJobs.Count)
        Write-DebugMessage ('[Main] Get-VBRJob returned {0} job(s).' -f $vbrJobs.Count)
        $idx = 0
        foreach ($job in $vbrJobs) {
            $idx++
            $jn = if ($null -ne $job.PSObject.Properties['Name']) { $job.Name } else { '<unnamed>' }
            Write-ProgressMessage ('  Job {0}/{1}: {2}' -f $idx, $vbrJobs.Count, $jn)
            try {
                Add-JobReportFromJob -Job $job -Source 'Get-VBRJob' -Results ([ref]$allReports)
            } catch {
                Write-Warning ('  Unable to process job "{0}": {1}' -f $jn, $_.Exception.Message)
                Write-DebugMessage ('[Main] Add-JobReportFromJob failed for "{0}":' -f $jn +
                    [Environment]::NewLine + (Format-ErrorRecord -ErrorRecord $_))
            }
            [void]$sessionFallbackJobs.Add($job)
            Write-DebugMessage ('[Main] Added Phase 1 job "{0}" to sessionFallbackJobs (count={1}).' -f $jn, $sessionFallbackJobs.Count)
        }
    } catch {
        Write-Warning ('Unable to enumerate jobs via Get-VBRJob: {0}' -f $_.Exception.Message)
        Write-DebugMessage ('[Main] Get-VBRJob threw:' + [Environment]::NewLine + (Format-ErrorRecord -ErrorRecord $_))
    }
} else {
    Write-ProgressMessage '  Get-VBRJob not available. Skipping.'
    Write-DebugMessage '[Main] Get-VBRJob cmdlet not found.'
}

# ---------------------------------------------------------------------------
# Phase 2 — Computer/agent backup jobs via Get-VBRComputerBackupJob
# ---------------------------------------------------------------------------
Write-ProgressMessage 'Phase 2 — Computer/agent backup jobs (Get-VBRComputerBackupJob).'
Write-DebugMessage '[Main] Phase 2 — Get-VBRComputerBackupJob'
if (Get-Command -Name 'Get-VBRComputerBackupJob' -ErrorAction SilentlyContinue) {
    try {
        Write-DebugMessage '[Main] Calling Get-VBRComputerBackupJob ...'
        $computerJobs = @(Get-VBRComputerBackupJob -ErrorAction Stop -WarningAction SilentlyContinue)
        Write-ProgressMessage ('  Found {0} computer backup job(s).' -f $computerJobs.Count)
        Write-DebugMessage ('[Main] Get-VBRComputerBackupJob returned {0} job(s).' -f $computerJobs.Count)
        $idx = 0
        foreach ($job in $computerJobs) {
            $idx++
            $jn = if ($null -ne $job.PSObject.Properties['Name']) { $job.Name } else { '<unnamed>' }
            Write-ProgressMessage ('  Computer job {0}/{1}: {2}' -f $idx, $computerJobs.Count, $jn)
            try {
                Add-JobReportFromJob -Job $job -Source 'Get-VBRComputerBackupJob' -Results ([ref]$allReports)
            } catch {
                Write-Warning ('  Unable to process computer backup job "{0}": {1}' -f $jn, $_.Exception.Message)
                Write-DebugMessage ('[Main] Add-JobReportFromJob failed for computer job "{0}":' -f $jn +
                    [Environment]::NewLine + (Format-ErrorRecord -ErrorRecord $_))
            }
            [void]$sessionFallbackJobs.Add($job)
            Write-DebugMessage ('[Main] Added Phase 2 job "{0}" to sessionFallbackJobs (count={1}).' -f $jn, $sessionFallbackJobs.Count)
        }
    } catch {
        Write-Warning ('Unable to enumerate computer backup jobs: {0}' -f $_.Exception.Message)
        Write-DebugMessage ('[Main] Get-VBRComputerBackupJob threw:' + [Environment]::NewLine + (Format-ErrorRecord -ErrorRecord $_))
    }
} else {
    Write-ProgressMessage '  Get-VBRComputerBackupJob not available. Skipping.'
    Write-DebugMessage '[Main] Get-VBRComputerBackupJob cmdlet not found.'
}

# ---------------------------------------------------------------------------
# Phase 3 — SOBR capacity-tier offload sessions
# ---------------------------------------------------------------------------
Write-ProgressMessage 'Phase 3 — SOBR capacity-tier offload (Get-VBRCapacityTierSyncSession).'
Write-DebugMessage '[Main] Phase 3 — Get-VBRCapacityTierSyncSession'
if (Get-Command -Name 'Get-VBRCapacityTierSyncSession' -ErrorAction SilentlyContinue) {
    try {
        Write-DebugMessage '[Main] Calling Get-VBRCapacityTierSyncSession ...'
        $sobrSessions = @(Get-VBRCapacityTierSyncSession -ErrorAction Stop)
        Write-ProgressMessage ('  Found {0} capacity-tier session(s).' -f $sobrSessions.Count)
        Write-DebugMessage ('[Main] Get-VBRCapacityTierSyncSession returned {0} session(s).' -f $sobrSessions.Count)
        $inWindow = @($sobrSessions | Where-Object { Test-SessionInWindow -Session $_ })
        Write-ProgressMessage ('  {0} session(s) within window.' -f $inWindow.Count)
        Write-DebugMessage ('[Main] SOBR sessions in window: {0}' -f $inWindow.Count)

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
            Write-DebugMessage ('[Main] SOBR report: name={0} result={1}' -f $report.job_name, $report.result)
            [void]$allReports.Add($report)
        }
    } catch {
        Write-Warning ('Unable to enumerate capacity-tier sessions: {0}' -f $_.Exception.Message)
        Write-DebugMessage ('[Main] Get-VBRCapacityTierSyncSession threw:' + [Environment]::NewLine + (Format-ErrorRecord -ErrorRecord $_))
    }
} else {
    Write-ProgressMessage '  Get-VBRCapacityTierSyncSession not available. Skipping.'
    Write-DebugMessage '[Main] Get-VBRCapacityTierSyncSession cmdlet not found.'
}

# ---------------------------------------------------------------------------
# Phase 4 — Configuration backup sessions (housekeeping)
# ---------------------------------------------------------------------------
Write-ProgressMessage 'Phase 4 — Configuration backup (Get-VBRConfigurationBackupJobSession / Get-VBRConfigurationBackupJob).'
Write-DebugMessage '[Main] Phase 4 — Configuration backup sessions'
$configBackupHandled = $false

if (Get-Command -Name 'Get-VBRConfigurationBackupJobSession' -ErrorAction SilentlyContinue) {
    try {
        Write-DebugMessage '[Main] Calling Get-VBRConfigurationBackupJobSession ...'
        $configSessions = @(Get-VBRConfigurationBackupJobSession -ErrorAction Stop)
        Write-ProgressMessage ('  Found {0} configuration backup session(s).' -f $configSessions.Count)
        Write-DebugMessage ('[Main] Get-VBRConfigurationBackupJobSession returned {0} session(s).' -f $configSessions.Count)
        $inWindow = @($configSessions | Where-Object { Test-SessionInWindow -Session $_ })
        Write-ProgressMessage ('  {0} session(s) within window.' -f $inWindow.Count)
        Write-DebugMessage ('[Main] Configuration backup sessions in window: {0}' -f $inWindow.Count)

        # There is normally one configuration backup job; group defensively to pick
        # the most-recent session per logical job name.
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
            $report = Build-JobReport -Session $s -JobType 'ConfigurationBackup' -Source 'Get-VBRConfigurationBackupJobSession'
            Write-DebugMessage ('[Main] Config backup report: name={0} result={1}' -f $report.job_name, $report.result)
            [void]$allReports.Add($report)
        }
        $configBackupHandled = $true
    } catch {
        Write-Warning ('Unable to enumerate configuration backup sessions: {0}' -f $_.Exception.Message)
        Write-DebugMessage ('[Main] Get-VBRConfigurationBackupJobSession threw:' + [Environment]::NewLine + (Format-ErrorRecord -ErrorRecord $_))
    }
} else {
    Write-ProgressMessage '  Get-VBRConfigurationBackupJobSession not available.'
    Write-DebugMessage '[Main] Get-VBRConfigurationBackupJobSession cmdlet not found.'
}

# Fallback: if the dedicated session cmdlet was unavailable or threw, try via the job object.
if (-not $configBackupHandled) {
    if (Get-Command -Name 'Get-VBRConfigurationBackupJob' -ErrorAction SilentlyContinue) {
        try {
            Write-DebugMessage '[Main] Calling Get-VBRConfigurationBackupJob (fallback) ...'
            $configJob = Get-VBRConfigurationBackupJob -ErrorAction Stop
            if ($null -ne $configJob) {
                $jn = if ($null -ne $configJob.PSObject.Properties['Name']) { [string]$configJob.Name } else { 'ConfigurationBackup' }
                Write-ProgressMessage ('  Configuration backup job found: {0}' -f $jn)
                try {
                    Add-JobReportFromJob -Job $configJob -Source 'Get-VBRConfigurationBackupJob' -Results ([ref]$allReports)
                } catch {
                    Write-Warning ('  Unable to process configuration backup job "{0}": {1}' -f $jn, $_.Exception.Message)
                    Write-DebugMessage ('[Main] Add-JobReportFromJob failed for config backup job "{0}":' -f $jn +
                        [Environment]::NewLine + (Format-ErrorRecord -ErrorRecord $_))
                }
            } else {
                Write-ProgressMessage '  Get-VBRConfigurationBackupJob returned no job.'
                Write-DebugMessage '[Main] Get-VBRConfigurationBackupJob returned null.'
            }
        } catch {
            Write-Warning ('Unable to get configuration backup job: {0}' -f $_.Exception.Message)
            Write-DebugMessage ('[Main] Get-VBRConfigurationBackupJob threw:' + [Environment]::NewLine + (Format-ErrorRecord -ErrorRecord $_))
        }
    } else {
        Write-ProgressMessage '  Get-VBRConfigurationBackupJob not available. Skipping.'
        Write-DebugMessage '[Main] Get-VBRConfigurationBackupJob cmdlet not found.'
    }
}

# ---------------------------------------------------------------------------
# Phase 5 — Repository offload / extent-sync sessions (housekeeping)
# ---------------------------------------------------------------------------
Write-ProgressMessage 'Phase 5 — Repository offload sessions (Get-VBRRepositoryExtentSyncSession).'
Write-DebugMessage '[Main] Phase 5 — Repository offload / extent-sync sessions'
if (Get-Command -Name 'Get-VBRRepositoryExtentSyncSession' -ErrorAction SilentlyContinue) {
    try {
        Write-DebugMessage '[Main] Calling Get-VBRRepositoryExtentSyncSession ...'
        $repoOffloadSessions = @(Get-VBRRepositoryExtentSyncSession -ErrorAction Stop)
        Write-ProgressMessage ('  Found {0} repository offload session(s).' -f $repoOffloadSessions.Count)
        Write-DebugMessage ('[Main] Get-VBRRepositoryExtentSyncSession returned {0} session(s).' -f $repoOffloadSessions.Count)
        $inWindow = @($repoOffloadSessions | Where-Object { Test-SessionInWindow -Session $_ })
        Write-ProgressMessage ('  {0} session(s) within window.' -f $inWindow.Count)
        Write-DebugMessage ('[Main] Repository offload sessions in window: {0}' -f $inWindow.Count)

        # Group by name to pick the most-recent session per repository/offload job.
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
            $report = Build-JobReport -Session $s -JobType 'RepositoryOffload' -Source 'Get-VBRRepositoryExtentSyncSession'
            Write-DebugMessage ('[Main] Repo offload report: name={0} result={1}' -f $report.job_name, $report.result)
            [void]$allReports.Add($report)
        }
    } catch {
        Write-Warning ('Unable to enumerate repository offload sessions: {0}' -f $_.Exception.Message)
        Write-DebugMessage ('[Main] Get-VBRRepositoryExtentSyncSession threw:' + [Environment]::NewLine + (Format-ErrorRecord -ErrorRecord $_))
    }
} else {
    Write-ProgressMessage '  Get-VBRRepositoryExtentSyncSession not available. Skipping.'
    Write-DebugMessage '[Main] Get-VBRRepositoryExtentSyncSession cmdlet not found.'
}

# ---------------------------------------------------------------------------
# Phase 7 — Defined Repository baseline (text mode only)
#   Collects repository utilisation and stores it in $definedRepositorySection
#   so it can be included in the human-readable report body.
#   In -Json mode this phase is skipped so stdout remains a pure JSON array.
# ---------------------------------------------------------------------------
Write-ProgressMessage 'Phase 7 — Defined Repository baseline (repository utilisation).'
Write-DebugMessage '[Main] Phase 7 — Defined Repository baseline.'
if (-not $Json) {
    try {
        $definedRepositorySection = New-DefinedRepositorySectionText
        Write-DebugMessage ('[Main] Defined Repository section ready, {0} char(s).' -f $definedRepositorySection.Length)
    } catch {
        Write-Warning ('Defined Repository baseline failed: {0}' -f $_.Exception.Message)
        Write-DebugMessage ('[Main] Defined Repository baseline failed:' + [Environment]::NewLine + (Format-ErrorRecord -ErrorRecord $_))
        $definedRepositorySection = New-DefinedRepositoryPlaceholderSection -Message '(repository utilisation unavailable)'
    }
}

# ---------------------------------------------------------------------------
# Phase 8 — VBR Licensing (text mode only)
#   Collects license information and stores it in $licensingSection so it can
#   be included in the human-readable report body after the repository section.
#   In -Json mode this phase is skipped so stdout remains a pure JSON array.
# ---------------------------------------------------------------------------
Write-ProgressMessage 'Phase 8 — VBR Licensing (license information).'
Write-DebugMessage '[Main] Phase 8 — VBR Licensing.'
if (-not $Json) {
    try {
        $licensingSection = New-VBRLicensingSectionText
        Write-DebugMessage ('[Main] Licensing section ready, {0} char(s).' -f $licensingSection.Length)
    } catch {
        Write-Warning ('VBR Licensing phase failed: {0}' -f $_.Exception.Message)
        Write-DebugMessage ('[Main] VBR Licensing phase failed:' + [Environment]::NewLine + (Format-ErrorRecord -ErrorRecord $_))
        $licensingSection = ''
    }
}

# ---------------------------------------------------------------------------
# Phase 9 — Backup Versions (text mode only)
#   Counts the number of backup versions per machine in each repository and
#   stores the result in $backupVersionsSection so it can be included in the
#   human-readable report body after the licensing section.
#   In -Json mode this phase is skipped so stdout remains a pure JSON array.
# ---------------------------------------------------------------------------
Write-ProgressMessage 'Phase 9 — Backup Versions (versions per machine per repository).'
Write-DebugMessage '[Main] Phase 9 — Backup Versions.'
if (-not $Json) {
    try {
        $backupVersionsSection = New-BackupVersionsSectionText
        Write-DebugMessage ('[Main] Backup Versions section ready, {0} char(s).' -f $backupVersionsSection.Length)
    } catch {
        Write-Warning ('Backup Versions phase failed: {0}' -f $_.Exception.Message)
        Write-DebugMessage ('[Main] Backup Versions phase failed:' + [Environment]::NewLine + (Format-ErrorRecord -ErrorRecord $_))
        $backupVersionsSection = ''
    }
}

# ---------------------------------------------------------------------------
# Phase 10 — SOBR Offloads (text mode only)
#   Collects all SOBR archive/capacity-tier offload sessions within the window
#   and stores the result in $sobrOffloadSection for inclusion in the report.
#   In -Json mode this phase is skipped so stdout remains a pure JSON array.
# ---------------------------------------------------------------------------
Write-ProgressMessage 'Phase 10 — SOBR Offloads (archive/offload sessions).'
Write-DebugMessage '[Main] Phase 10 — SOBR Offloads.'
if (-not $Json) {
    try {
        $sobrOffloadSection = New-SobrOffloadSectionText
        Write-DebugMessage ('[Main] SOBR Offloads section ready, {0} char(s).' -f $sobrOffloadSection.Length)
    } catch {
        Write-Warning ('SOBR Offloads phase failed: {0}' -f $_.Exception.Message)
        Write-DebugMessage ('[Main] SOBR Offloads phase failed:' + [Environment]::NewLine + (Format-ErrorRecord -ErrorRecord $_))
        $sobrOffloadSection = ''
    }
}

Write-ProgressMessage ('Enumeration complete. Total report entries before filtering: {0}.' -f $allReports.Count)
Write-DebugMessage ('[Main] Enumeration complete. Total entries: {0}' -f $allReports.Count)

# ---------------------------------------------------------------------------
# Apply -OnlyFailures filter
# ---------------------------------------------------------------------------
Write-DebugMessage ('[Main] Applying OnlyFailures filter. OnlyFailures={0}; input entries={1}' -f [bool]$OnlyFailures, $allReports.Count)

if ($OnlyFailures) {
    $filtered = foreach ($report in $allReports) {
        $resultText = if ($null -ne $report.result) { [string]$report.result } else { '' }
        $runningFor = Get-PropertyValue -InputObject $report -Names @('running_for')
        if ($resultText -imatch 'Failed|Warning|Warn|Error|Stopped') {
            $report
            continue
        }
        if (-not [string]::IsNullOrWhiteSpace([string]$runningFor)) {
            Write-DebugMessage ('[Main] Retaining running session under -OnlyFailures: {0}' -f $report.job_name)
            $report
        }
    }
} else {
    $filtered = foreach ($report in $allReports) {
        $report
    }
}

$filtered = @($filtered)
Write-DebugMessage ('[Main] Filtered entries: {0}.' -f $filtered.Count)

# ---------------------------------------------------------------------------
# Sort: Failed first (0), Warning (1), other (2); then end_time descending.
# ---------------------------------------------------------------------------
Write-DebugMessage '[Main] Sorting results by severity then end time.'
$sorted = foreach ($report in ($filtered | Sort-Object -Property @(
    @{ Expression = { [int](Get-ResultSeverityOrder -Result $_.result) }; Descending = $false },
    @{ Expression = { Get-SortableTicks -Value $_.end_time }; Descending = $true }
))) {
    $report
}

$sorted = @($sorted)

# ---------------------------------------------------------------------------
# Compute summary counts
# ---------------------------------------------------------------------------
$totalJobs   = $sorted.Count
$failedCount = 0
$warnCount   = 0
$successCount= 0
$withError   = 0

foreach ($report in $sorted) {
    $resultText = if ($null -ne $report.result) { [string]$report.result } else { '' }

    if ($resultText -imatch 'Failed|Fail') {
        $failedCount++
    }
    if ($resultText -imatch 'Warning|Warn') {
        $warnCount++
    }
    if ($resultText -imatch 'Success') {
        $successCount++
    }
    if (-not [string]::IsNullOrWhiteSpace($report.last_error)) {
        $withError++
    }
}

Write-DebugMessage ('[Main] Summary: total={0} failed={1} warning={2} success={3} withError={4}' `
    -f $totalJobs, $failedCount, $warnCount, $successCount, $withError)

$reportBody = New-CollectorReportBody -Reports $sorted `
    -TotalJobs $totalJobs -FailedCount $failedCount -WarnCount $warnCount `
    -SuccessCount $successCount -WithError $withError `
    -DefinedJobsSection $definedJobsSection -DefinedRepositorySection $definedRepositorySection `
    -LicensingSection $licensingSection -BackupVersionsSection $backupVersionsSection `
    -SobrOffloadSection $sobrOffloadSection
Write-DebugMessage ('[Main] Canonical report body length: {0} characters.' -f $reportBody.Length)

# ---------------------------------------------------------------------------
# Output
# ---------------------------------------------------------------------------
Write-DebugMessage '[Main] Producing output.'
if ($Json) {
    ConvertTo-Json -InputObject @($sorted) -Depth 6
    Write-Warning ('Summary: jobs={0} failed={1} warning={2} success={3} with_error={4}' `
        -f $totalJobs, $failedCount, $warnCount, $successCount, $withError)
} else {
    Write-Output $reportBody
}

$null = Write-CollectorReportBodyToDisk -Body $reportBody
$mailSubject = Get-CollectorMailSubject -FailedCount $failedCount -WarnCount $warnCount
$null = Send-CollectorReportEmail -Body $reportBody -Subject $mailSubject
Remove-OldCollectorFiles

Write-DebugMessage '[Main] Script completed successfully.'
if ($script:CollectorDebugEnabled -and $null -ne $script:DebugLogFile) {
    Write-Warning ('[CollectorDebug] Debug log written to: {0}' -f $script:DebugLogFile)
}
