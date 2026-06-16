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
    When set, only include jobs whose most recent session result is Failed,
    Warning, Error, or Stopped.  Successful/skipped jobs are omitted.

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
      - After the report is built, the same human-readable body is written to
        E:\VEEAM_LOGS\COLLECTOR by default, emailed by default, and old
        collector-created files in that directory are removed after 7 days.

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
      - When dedicated housekeeping session cmdlets are unavailable, the script
        also uses a Get-VBRSession fallback to find repository/offload/
        configuration housekeeping sessions.

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
    $warningDetails = Get-VeeamWarningDetails -Session $Session

    return [pscustomobject][ordered]@{
        job_name        = $name
        job_type        = $type
        result          = $result
        start_time      = if ($null -ne $start) { $start.ToString('o') } else { $null }
        end_time        = if ($null -ne $end)   { $end.ToString('o')   } else { $null }
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
        [Parameter(Mandatory)] [object[]]$Reports,
        [Parameter(Mandatory)] [int]$TotalJobs,
        [Parameter(Mandatory)] [int]$FailedCount,
        [Parameter(Mandatory)] [int]$WarnCount,
        [Parameter(Mandatory)] [int]$SuccessCount,
        [Parameter(Mandatory)] [int]$WithError
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

$allReports = New-Object 'System.Collections.Generic.List[object]'

# Job objects collected from Phase 1/2 for use as Get-VBRSession -Job arguments
# when bare Get-VBRSession invocation is unsafe (mandatory -Job parameter set).
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
# Phase 6 — Generic housekeeping fallback via Get-VBRSession
# ---------------------------------------------------------------------------
Write-ProgressMessage 'Phase 6 — Housekeeping fallback (Get-VBRSession).'
Write-DebugMessage '[Main] Phase 6 — Get-VBRSession fallback'
if (Get-Command -Name 'Get-VBRSession' -ErrorAction SilentlyContinue) {
    try {
        $housekeepingTerms = @('Offload', 'Capacity', 'Archive', 'Repository', 'Object', 'SOBR', 'Sync', 'Config', 'Configuration')
        $housekeepingPattern = (($housekeepingTerms | ForEach-Object { [regex]::Escape($_) }) -join '|')
        $fallbackCandidates = New-Object 'System.Collections.Generic.List[object]'

        $canInvokeGetVBRSessionBare = Test-CmdletCanInvokeWithoutArguments -CmdletName 'Get-VBRSession'
        $hasGetVBRSessionJob = Test-CmdletHasParameter -CmdletName 'Get-VBRSession' -ParameterName 'Job'
        Write-DebugMessage ('[Main] Get-VBRSession: canInvokeBare={0} hasJobParam={1} sessionFallbackJobs.Count={2}' -f $canInvokeGetVBRSessionBare, $hasGetVBRSessionJob, $sessionFallbackJobs.Count)

        if ($canInvokeGetVBRSessionBare) {
            Write-DebugMessage '[Main] Get-VBRSession can be invoked without arguments; using bare/type-enum path.'

            # Try -Type enum discovery first when available.
            try {
                $sessionCommand = Get-Command -Name 'Get-VBRSession' -ErrorAction Stop
                $typeParameter = $sessionCommand.Parameters['Type']
                $typeNames = @()
                if ($null -ne $typeParameter) {
                    $typeParameterType = $typeParameter.ParameterType
                    if ($typeParameterType.IsArray) {
                        $typeParameterType = $typeParameterType.GetElementType()
                    }
                    if ($null -ne $typeParameterType -and $typeParameterType.IsEnum) {
                        $typeNames = [enum]::GetNames($typeParameterType)
                    }
                }

                $matchingTypes = @($typeNames | Where-Object {
                    $typeName = [string]$_
                    $housekeepingTerms | Where-Object { $typeName -imatch [regex]::Escape($_) }
                } | Select-Object -Unique)

                foreach ($typeName in $matchingTypes) {
                    try {
                        Write-DebugMessage ('[Main] Get-VBRSession fallback querying -Type {0}' -f $typeName)
                        $typedSessions = @(Get-VBRSession -Type $typeName -ErrorAction Stop)
                        foreach ($s in $typedSessions) {
                            if ($null -ne $s) { [void]$fallbackCandidates.Add($s) }
                        }
                    } catch {
                        Write-DebugMessage ('[Main] Get-VBRSession -Type {0} threw:' -f $typeName +
                            [Environment]::NewLine + (Format-ErrorRecord -ErrorRecord $_))
                    }
                }
            } catch {
                Write-DebugMessage ('[Main] Get-VBRSession -Type discovery failed:' + [Environment]::NewLine + (Format-ErrorRecord -ErrorRecord $_))
            }

            # Conservative text-match fallback over generic session properties.
            try {
                Write-DebugMessage '[Main] Calling Get-VBRSession (no arguments) for text-match fallback.'
                $allSessions = @(Get-VBRSession -ErrorAction Stop)
                foreach ($s in $allSessions) {
                    if ($null -eq $s) { continue }
                    $fields = @(
                        Get-PropertyValue -InputObject $s -Names @('Name'),
                        Get-PropertyValue -InputObject $s -Names @('JobName'),
                        Get-PropertyValue -InputObject $s -Names @('SessionName'),
                        Get-PropertyValue -InputObject $s -Names @('SessionType'),
                        Get-PropertyValue -InputObject $s -Names @('JobType'),
                        Get-PropertyValue -InputObject $s -Names @('Type'),
                        Get-PropertyValue -InputObject $s -Names @('Operation'),
                        Get-PropertyValue -InputObject $s -Names @('Description')
                    )
                    $searchText = (($fields | Where-Object { $null -ne $_ } | ForEach-Object { [string]$_ }) -join ' ')
                    if (-not [string]::IsNullOrWhiteSpace($searchText) -and $searchText -imatch $housekeepingPattern) {
                        [void]$fallbackCandidates.Add($s)
                    }
                }
            } catch {
                Write-Warning ('Unable to enumerate Get-VBRSession fallback sessions: {0}' -f $_.Exception.Message)
                Write-DebugMessage ('[Main] Get-VBRSession fallback enumeration threw:' + [Environment]::NewLine + (Format-ErrorRecord -ErrorRecord $_))
            }
        } elseif ($hasGetVBRSessionJob -and $sessionFallbackJobs.Count -gt 0) {
            # Bare invocation is unsafe (mandatory -Job parameter set); use job-scoped querying.
            Write-ProgressMessage ('  Get-VBRSession requires -Job; querying {0} known job(s) individually.' -f $sessionFallbackJobs.Count)
            Write-DebugMessage ('[Main] Get-VBRSession bare call unsafe; using job-scoped fallback with {0} job(s).' -f $sessionFallbackJobs.Count)
            foreach ($job in $sessionFallbackJobs) {
                $jn = if ($null -ne $job -and $null -ne $job.PSObject.Properties['Name']) { [string]$job.Name } else { '<unnamed>' }
                Write-DebugMessage ('[Main] Calling Get-VBRSession -Job "{0}".' -f $jn)
                try {
                    $jobSessions = @(Get-VBRSession -Job $job -ErrorAction Stop)
                    Write-DebugMessage ('[Main] Get-VBRSession -Job "{0}" returned {1} session(s).' -f $jn, $jobSessions.Count)
                    foreach ($s in $jobSessions) {
                        if ($null -ne $s) { [void]$fallbackCandidates.Add($s) }
                    }
                } catch {
                    Write-DebugMessage ('[Main] Get-VBRSession -Job "{0}" threw:' -f $jn +
                        [Environment]::NewLine + (Format-ErrorRecord -ErrorRecord $_))
                }
            }
        } else {
            Write-ProgressMessage '  Get-VBRSession cannot be invoked safely without mandatory parameters and no job-scoped fallback is available. Skipping.'
            Write-DebugMessage ('[Main] Get-VBRSession skipped: canInvokeBare={0} hasJobParam={1} knownJobs={2}.' -f $canInvokeGetVBRSessionBare, $hasGetVBRSessionJob, $sessionFallbackJobs.Count)
        }

        Write-DebugMessage ('[Main] Get-VBRSession fallback candidate count before window filter: {0}' -f $fallbackCandidates.Count)
        $inWindow = @($fallbackCandidates | Where-Object { $null -ne $_ -and (Test-SessionInWindow -Session $_) })
        Write-ProgressMessage ('  {0} fallback housekeeping session candidate(s) in window.' -f $inWindow.Count)
        Write-DebugMessage ('[Main] Get-VBRSession fallback candidates in window: {0}' -f $inWindow.Count)

        # Group by logical name/type and keep the newest session per key.
        $grouped = @{}
        foreach ($s in $inWindow) {
            $name = Get-SessionName -Session $s
            $type = Get-SessionType -Session $s
            $groupKey = ('{0}|{1}' -f $name, $type)
            $sEnd  = Get-SessionEndTime -Session $s
            $sTime = if ($null -ne $sEnd) { Get-SortableTicks -Value $sEnd } else {
                $st = Get-SessionStartTime -Session $s
                if ($null -ne $st) { Get-SortableTicks -Value $st } else { [long]0 }
            }
            if (-not $grouped.ContainsKey($groupKey)) {
                $grouped[$groupKey] = @{ Session = $s; Time = $sTime }
            } elseif ($sTime -gt $grouped[$groupKey].Time) {
                $grouped[$groupKey] = @{ Session = $s; Time = $sTime }
            }
        }

        foreach ($entry in $grouped.Values) {
            $s = $entry.Session
            $sessionId = Get-ObjectIdentity -InputObject $s
            if (-not $script:SeenSessions.Add($sessionId)) { continue }
            $sessionType = Get-SessionType -Session $s
            $fallbackJobType = if ([string]::IsNullOrWhiteSpace($sessionType)) { 'HousekeepingSessionFallback' } else { ('HousekeepingSessionFallback/{0}' -f $sessionType) }
            $report = Build-JobReport -Session $s -JobType $fallbackJobType -Source 'Get-VBRSessionFallback'
            Write-DebugMessage ('[Main] Fallback report: name={0} type={1} result={2}' -f $report.job_name, $report.job_type, $report.result)
            [void]$allReports.Add($report)
        }
    } catch {
        Write-Warning ('Unable to execute Get-VBRSession fallback phase: {0}' -f $_.Exception.Message)
        Write-DebugMessage ('[Main] Get-VBRSession fallback phase threw:' + [Environment]::NewLine + (Format-ErrorRecord -ErrorRecord $_))
    }
} else {
    Write-ProgressMessage '  Get-VBRSession not available. Skipping fallback.'
    Write-DebugMessage '[Main] Get-VBRSession cmdlet not found for fallback phase.'
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
        if ($resultText -imatch 'Failed|Warning|Warn|Error|Stopped') {
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
    -SuccessCount $successCount -WithError $withError
Write-DebugMessage ('[Main] Canonical report body length: {0} characters.' -f $reportBody.Length)

# ---------------------------------------------------------------------------
# Output
# ---------------------------------------------------------------------------
Write-DebugMessage '[Main] Producing output.'
if ($Json) {
    $sorted | ConvertTo-Json -Depth 6
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
