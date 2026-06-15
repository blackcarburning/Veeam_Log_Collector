# Veeam_Log_Collector

A focused PowerShell script that reports the **last error text** from the most recent session for every Veeam backup, replication, backup-copy, agent, and SOBR capacity-tier offload job.

## What it does

For every backup/replication/offload job it finds the **most recent session within the last N hours**, extracts the last error/warning text, and produces a compact, LLM-friendly report.  No log bundles are created — the output is small enough to paste directly into an LLM prompt or pipe to `jq`.

## Requirements

- PowerShell 7 (recommended) or Windows PowerShell 5.1 with VeeamPSSnapIn
- Veeam Backup & Replication console / PowerShell components installed
- Run as an elevated session on the VBR server (or a host with the VBR console installed)

## Usage

```powershell
# Default: show all jobs' last session in the last 24 hours
.\Veeam_Collector.ps1

# Show only failed/warning jobs in the last 48 hours
.\Veeam_Collector.ps1 -Hours 48 -OnlyFailures

# Emit JSON (parseable by ConvertFrom-Json / jq); progress goes to Warning stream
.\Veeam_Collector.ps1 -Json
```

## Parameters

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `-Hours` | int | 24 | Time window in hours (1–8760). Sessions outside this window are ignored. |
| `-Json` | switch | — | Emit a JSON array on stdout. Progress messages go to the Warning stream. |
| `-OnlyFailures` | switch | — | Only include jobs with a Failed or Warning last session. |

## Output

**Text mode** (default): one concise block per job showing name, type, result, end time, and last error text (empty when successful).

**JSON mode** (`-Json`): a single JSON array, one object per job, with fields:
- `job_name`, `job_type`, `result`
- `start_time`, `end_time` (ISO 8601)
- `last_error` (empty string when there is no error)
- `warning_details` (deeper per-task/session warning text when available)
- `source` (which cmdlet enumerated the job)

Output is sorted: Failed jobs first, then Warning, then others; within each group sorted by end time (most recent first).

A summary line is printed at the end (to Warning stream in `-Json` mode): jobs scanned, failed, warning, success, and how many had error text.

## How errors are extracted

For each session the script tries, in order:
1. `$session.GetLastError()` — the primary Veeam API
2. `$session.GetTaskSessions()` and `Get-VBRTaskSession -Session` — per-task details
3. Task `GetLastError()` / `GetDetails()` and task logger records
4. Session logger records (`UpdatedRecords` or `Records`) filtered to warning/error states

All method/property accesses are guarded defensively; missing methods are silently skipped.
