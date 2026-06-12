<#
.SYNOPSIS
    Exports Veeam Backup & Replication logs for the last N hours using Export-VBRLogs.

.DESCRIPTION
    Uses the Veeam Backup PowerShell module/snap-in to call Export-VBRLogs, which
    produces an archive/bundle of Veeam log files for the selected time window.

    The exported log bundle is written to -OutputPath (or a timestamped temporary
    directory when -OutputPath is not specified).  The script prints the path to
    the exported logs so you know exactly where to find them.

    Progress and status messages are printed throughout so you can see what the
    script is doing while it runs.  In default (human-readable) mode these go to
    standard output.  In -Json mode they are sent to the Warning stream so that
    standard output remains a parseable JSON object describing the export result.

    The script inspects Export-VBRLogs parameter metadata at runtime and binds
    the appropriate time-window parameters (e.g. -From/-To, -StartTime/-EndTime,
    or a duration-style parameter such as -Last/-Hours) depending on what the
    installed Veeam PowerShell version exposes.

.PARAMETER Hours
    Time window (in hours) to export logs for. Default is 24 (last 24 hours).

.PARAMETER OutputPath
    Directory where the exported log bundle will be written.  If omitted, a
    timestamped sub-directory is created under the system temporary folder and
    its path is printed to stdout on completion.

.PARAMETER Json
    Emit a single JSON object on stdout describing the export result (path,
    start/end time, status).  Progress/status messages are sent to the Warning
    stream in this mode so that standard output remains parseable JSON.

.EXAMPLE
    .\Veeam_Collector.ps1

    Exports Veeam logs for the last 24 hours to a temp directory and prints the path.

.EXAMPLE
    .\Veeam_Collector.ps1 -Hours 48 -OutputPath C:\Temp\VeeamLogs

    Exports logs for the last 48 hours to C:\Temp\VeeamLogs.

.EXAMPLE
    .\Veeam_Collector.ps1 -Json

    Exports logs for the last 24 hours and emits a JSON result object on stdout.

.NOTES
    Usage notes:
      - Run this script in PowerShell on a Veeam Backup & Replication server or a host
        with the Veeam console/PowerShell components installed.
      - The script tries Veeam.Backup.PowerShell first, then VeeamPSSnapIn fallback.
      - Export-VBRLogs must be available after loading Veeam PowerShell components.
        If it is not present, the script throws a clear error explaining which Veeam
        PowerShell version is required.
      - Human-readable mode (default, -Json not specified) prints timestamped progress
        and the final export path to standard output.
      - -Json mode routes all progress/status messages to the Warning stream; standard
        output contains only a single JSON object suitable for machine consumption.

    PowerShell version requirements:
      - PowerShell 7.0 or later: the modern Veeam.Backup.PowerShell module is loaded.
      - Windows PowerShell 5.1 (Desktop edition): the modern module manifest declares a
        minimum PS version of 7.0 and cannot be loaded by PS 5.1. The script catches that
        failure and automatically falls back to the legacy VeeamPSSnapIn snap-in.
        Ensure VeeamPSSnapIn is registered (it is included with the Veeam Backup & Replication
        console components) when running under Windows PowerShell 5.1.

    Run requirements:
    Run in an elevated PowerShell session on the Veeam Backup & Replication server or a host
    with the Veeam console/PowerShell components installed.
#>

[CmdletBinding()]
param(
    [ValidateRange(1, 8760)]
    [int]$Hours = 24,

    # Directory where the exported log bundle will be written.
    # Defaults to a timestamped sub-directory under the system temp folder.
    [string]$OutputPath = '',

    # Emit a JSON result object on stdout instead of human-readable text.
    # Progress goes to the Warning stream in this mode.
    [switch]$Json
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$script:EndTime   = Get-Date
$script:StartTime = $script:EndTime.AddHours(-[Math]::Abs($Hours))

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
# Test-CmdletHasParameter
#   Returns $true when the named cmdlet exposes the named parameter in any
#   parameter set.
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

# ---------------------------------------------------------------------------
# Resolve-ExportOutputPath
#   Returns the resolved output directory.  Creates the directory if it does
#   not already exist.  When -OutputPath was not supplied, a timestamped
#   sub-directory is created under the system temp folder.
# ---------------------------------------------------------------------------
function Resolve-ExportOutputPath {
    [CmdletBinding()]
    param([string]$RequestedPath)

    if ([string]::IsNullOrWhiteSpace($RequestedPath)) {
        $stamp = (Get-Date).ToString('yyyyMMdd_HHmmss')
        $RequestedPath = [System.IO.Path]::Combine(
            [System.IO.Path]::GetTempPath(),
            ('VeeamLogs_{0}' -f $stamp)
        )
    }

    if (-not (Test-Path -LiteralPath $RequestedPath -PathType Container)) {
        Write-ProgressMessage ('Creating output directory: {0}' -f $RequestedPath)
        [void](New-Item -ItemType Directory -Path $RequestedPath -Force -ErrorAction Stop)
    }

    return (Resolve-Path -LiteralPath $RequestedPath).ProviderPath
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

    # Validate Export-VBRLogs is available.
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
    # Different Veeam versions expose the destination folder under different parameter
    # names. The list below covers known variants; the first match wins.
    $pathParamCandidates = @('Path', 'Folder', 'OutputPath', 'TargetPath', 'DestinationPath',
                              'FilePath', 'ExportPath', 'Target', 'Destination', 'Directory')
    $pathParam = $pathParamCandidates | Where-Object { $availableParams -contains $_ } |
                 Select-Object -First 1
    if ($null -ne $pathParam) {
        Write-ProgressMessage ('  Binding output path via -{0}' -f $pathParam)
        $exportParams[$pathParam] = $ResolvedOutputPath
    } else {
        # Fall back to positional argument if no recognised named parameter exists.
        Write-ProgressMessage '  No recognised path parameter found; passing output path as positional argument.'
        $exportParams['PositionalPath'] = $ResolvedOutputPath
    }

    # --- Resolve time-window parameters ---
    # Prefer explicit From/To (or equivalent) date range parameters.
    # The duration in whole hours is derived from StartTime/EndTime for use with
    # duration-style parameters (-Last, -Hours, etc.).
    $durationHours = [int][Math]::Ceiling(($EndTime - $StartTime).TotalHours)

    # Parameter name candidates listed in rough order of likelihood per Veeam version.
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
        # No From/To style params; look for a duration-style parameter.
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

    # Invoke Export-VBRLogs.
    Write-ProgressMessage ('Calling Export-VBRLogs ...')

    try {
        if ($exportParams.ContainsKey('PositionalPath')) {
            # Positional path fallback: remove the sentinel key and pass the value as
            # the first positional argument.
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
    }
    catch {
        throw ('Export-VBRLogs failed: {0}' -f $_.Exception.Message)
    }

    Write-ProgressMessage 'Export-VBRLogs completed successfully.'
    return $result
}

# ---------------------------------------------------------------------------
# Write-CollectorHeader
#   Prints a human-readable banner to stdout (skipped in -Json mode).
# ---------------------------------------------------------------------------
function Write-CollectorHeader {
    [CmdletBinding()]
    param()

    if ($Json) { return }

    Write-Output '============================================================'
    Write-Output 'Veeam Log Collector'
    Write-Output ('Window     : last {0} hour(s)  ({1:o} to {2:o})' -f $Hours, $script:StartTime, $script:EndTime)
    Write-Output ('Host       : {0}' -f $env:COMPUTERNAME)
    Write-Output ('PowerShell : {0} {1}' -f $PSVersionTable.PSEdition, $PSVersionTable.PSVersion)
    Write-Output '============================================================'
    Write-Output ''
}

# ---------------------------------------------------------------------------
# Collect-ExportedPaths
#   Inspects the return value of Export-VBRLogs (which varies by Veeam version)
#   and returns a list of file/folder path strings.  Falls back to the resolved
#   output directory itself when the cmdlet returns nothing.
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
            # Object with a path-like property.
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
        # Export-VBRLogs wrote files to disk but returned nothing.  Report the
        # output directory so the caller always has a useful path.
        $paths.Add($ResolvedOutputPath)
    }

    return $paths.ToArray()
}

# ===========================================================================
# Main
# ===========================================================================

Write-ProgressMessage ('Veeam Log Collector starting. Window: last {0} hour(s) ({1:o} to {2:o}).' `
    -f $Hours, $script:StartTime, $script:EndTime)

Import-VeeamPowerShell
Write-CollectorHeader

$resolvedOutputPath = Resolve-ExportOutputPath -RequestedPath $OutputPath
Write-ProgressMessage ('Output path : {0}' -f $resolvedOutputPath)

$exportResult = Invoke-VBRLogsExport `
    -StartTime          $script:StartTime `
    -EndTime            $script:EndTime `
    -ResolvedOutputPath $resolvedOutputPath

$exportedPaths = Collect-ExportedPaths -ExportResult $exportResult -ResolvedOutputPath $resolvedOutputPath

if ($Json) {
    $jsonResult = [ordered]@{
        status      = 'success'
        hours       = $Hours
        start_time  = $script:StartTime.ToString('o')
        end_time    = $script:EndTime.ToString('o')
        output_path = $resolvedOutputPath
        exported    = @($exportedPaths)
    }
    [pscustomobject]$jsonResult | ConvertTo-Json -Compress -Depth 4
} else {
    Write-Output ''
    Write-Output '------------------------------------------------------------'
    Write-Output 'Export complete.'
    Write-Output ('Window : last {0} hour(s)  ({1:o}  to  {2:o})' -f $Hours, $script:StartTime, $script:EndTime)
    Write-Output ''
    Write-Output 'Exported log location(s):'
    foreach ($p in $exportedPaths) {
        Write-Output ('  {0}' -f $p)
    }
    Write-Output '------------------------------------------------------------'
}
