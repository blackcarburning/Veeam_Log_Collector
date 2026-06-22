# Veeam_Log_Collector

A focused PowerShell script that reports the **last error text** from the most recent session for every Veeam backup, replication, backup-copy, agent, SOBR capacity-tier offload, configuration backup, and repository offload job.

## What it does

For every backup/replication/offload/housekeeping job it finds the **most recent session within the last N hours**, extracts the last error/warning text, and produces a compact, LLM-friendly report.  No log bundles are created — the output is small enough to paste directly into an LLM prompt or pipe to `jq`.

In normal text mode the report begins with a **Defined Jobs baseline** section showing all currently defined backup jobs together with their schedule, enabled status, next scheduled run, last run time, and last result.  It is followed immediately by a **Defined Repository** utilisation block showing repository, tier, parent, status, total, used, free, and used-percent columns.  These sections give operators (and LLMs) immediate context for which jobs and repositories exist and their current baseline state.  The sections are clearly delimited:

```
############### Defined Jobs BEGIN ###################
Job                                    Type        On  Next / schedule    Last run         Status      Last Result
------------------------------------------------------------------------------------------------------------------
VMware_Daily_Backup                    VM          Yes Daily 23:00        16/06/2026 02:26 Stopped     Warning
VMware_Daily_CATCHALL                  VM          Yes Daily 04:00        16/06/2026 02:45 Stopped     Warning
############### Defined Jobs END ###################
############### Defined Repository BEGIN ###################
Repository                     Tier             Parent               Status       Total       Used        Free        Used %
...
############### Defined Repository END ###################
```

The Defined Jobs and Defined Repository blocks are **omitted in `-Json` mode** so that stdout remains a pure JSON array.

Housekeeping-style processes (configuration backup and repository offload/extent-sync sessions) are included when their Veeam PowerShell cmdlets are available. If dedicated housekeeping session cmdlets are unavailable, the script also uses a `Get-VBRSession` fallback pass to discover relevant offload/repository/configuration sessions.

For in-progress offload sessions, the report includes how long the session has been running (`running_for`) and how much data has been processed so far (`data_processed`) when Veeam exposes it. Running sessions are always retained even when `-OnlyFailures` is used.

## Requirements

- PowerShell 7 (recommended) or Windows PowerShell 5.1 with VeeamPSSnapIn
- Veeam Backup & Replication console / PowerShell components installed
- Run as an elevated session on the VBR server (or a host with the VBR console installed)

## Usage

```powershell
# Default: show all jobs' last session in the last 24 hours
# Output starts with the Defined Jobs baseline, followed by per-job session details
.\Veeam_Collector.ps1

# Show only failed/warning jobs in the last 48 hours
.\Veeam_Collector.ps1 -Hours 48 -OnlyFailures

# Emit JSON (parseable by ConvertFrom-Json / jq); progress goes to Warning stream
# The Defined Jobs text block is skipped so stdout stays valid JSON
.\Veeam_Collector.ps1 -Json
```

## Parameters

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `-Hours` | int | 24 | Time window in hours (1–8760). Sessions outside this window are ignored. |
| `-Json` | switch | — | Emit a JSON array on stdout. Progress messages go to the Warning stream. |
| `-OnlyFailures` | switch | — | Only include jobs with a Failed, Warning, Error, or Stopped last session; currently running sessions are always included. |

## Output

**Text mode** (default): the report starts with a Defined Jobs baseline table, followed by one concise block per job showing name, type, result, end time, and last error text (empty when successful).

**JSON mode** (`-Json`): a single JSON array, one object per job, with fields:
- `job_name`, `job_type`, `result`
- `start_time`, `end_time` (ISO 8601)
- `running_for` (elapsed runtime for in-progress sessions; empty string otherwise)
- `data_processed` (formatted processed bytes for in-progress sessions when available; empty string otherwise)
- `last_error` (empty string when there is no error)
- `warning_details` (deeper per-task/session warning text when available)
- `source` (which cmdlet enumerated the job)

Output is sorted: Failed jobs first, then Warning, then others; within each group sorted by end time (most recent first).

A summary line is printed at the end (to Warning stream in `-Json` mode): jobs scanned, failed, warning, success, and how many had error text.

## Defined Jobs baseline

The Defined Jobs section is built from four job sources (collected in this order with case-insensitive de-duplication):

1. **Agent/computer backup jobs** (`Get-VBRComputerBackupJob`) — shown as type `Agent`
2. **Application backup jobs** (`Get-VBRApplicationBackupJob`) — shown as type `Application`
3. **Unstructured backup jobs** (`Get-VBRUnstructuredBackupJob`) — shown as type `File/NAS`
4. **Standard VBR backup jobs** (`Get-VBRJob`) — shown with the Veeam job type string; replication, backup-copy, tape, and SureBackup jobs are excluded from this section

Columns (fixed width):

| Column | Width | Description |
|--------|-------|-------------|
| Job | 38 | Job name |
| Type | 11 | Job type |
| On | 3 | `Yes` if scheduling is enabled, `No` otherwise |
| Next / schedule | 18 | Next scheduled run (if in the future) or a schedule description such as `Daily 22:00` |
| Last run | 16 | End time of the most recent session (`dd/MM/yyyy HH:mm`) |
| Status | 11 | Current session state (e.g. `Running`, `Stopped`, `Idle`) |
| Last Result | 11 | Actual result of the most recent session (e.g. `Success`, `Warning`, `Failed`) |

Session retrieval per job type:

| Job type | Session cmdlet used |
|----------|---------------------|
| Agent | `Get-VBRComputerBackupJobSession` (pre-fetched once, matched by ID/name) |
| Application | `Get-VBRApplicationBackupJobSession` (pre-fetched once, matched by ID/name) |
| File/NAS | `Get-VBRUnstructuredBackupSession -Name "$($Job.Name)*"` |
| Standard VBR | `FindLastSession()` on the job object; falls back to `Get-VBRBackupSession` |

## How errors are extracted

For each session the script tries, in order:
1. `$session.GetLastError()` — the primary Veeam API
2. `$session.GetTaskSessions()` and `Get-VBRTaskSession -Session` — per-task details
3. Task `GetLastError()` / `GetDetails()` and task logger records
4. Session logger records (`UpdatedRecords` or `Records`) filtered to warning/error states

All method/property accesses are guarded defensively; missing methods are silently skipped.
