#Requires -Version 5.1
[CmdletBinding()]
param(
    [string]$Environment  # e.g. -Environment PRD  (skips interactive menu)
)
. (Join-Path $PSScriptRoot "config.ps1")

# ---------------------------------------------------------------------------
# Environment selection
# ---------------------------------------------------------------------------
$envKeys = @($Environments.Keys)

if ($Environment) {
    if (-not $Environments.Contains($Environment)) {
        Write-Error "Unknown environment '$Environment'. Valid options: $($envKeys -join ', ')"
        exit 1
    }
    $selectedEnv = $Environment
} elseif ($envKeys.Count -eq 1) {
    $selectedEnv = $envKeys[0]
} else {
    Write-Host ""
    Write-Host "Select environment:"
    for ($i = 0; $i -lt $envKeys.Count; $i++) {
        $key = $envKeys[$i]
        Write-Host "  [$($i + 1)] $key - $($Environments[$key].Label)"
    }
    Write-Host ""
    do {
        $choice = Read-Host "Enter number (1-$($envKeys.Count))"
    } while ($choice -notmatch '^\d+$' -or [int]$choice -lt 1 -or [int]$choice -gt $envKeys.Count)
    $selectedEnv = $envKeys[[int]$choice - 1]
}

$Config.EtlSchema  = $Environments[$selectedEnv].EtlSchema
$Config.OracleDSN  = $Environments[$selectedEnv].OracleDSN
Write-Host "Environment: $selectedEnv - $($Environments[$selectedEnv].Label)"

# ---------------------------------------------------------------------------
# Benchmark timings (minutes) — extracted from Benchmark sheet, premier ETL comparison 265.xlsx
# ---------------------------------------------------------------------------
$Benchmark = @{
    'c_Delete fact records'               = 2084
    'c_SUM_PR_Do_MERGE'                   = 10
    'c_SUM_PR_Create_Staging'             = 85
    'c_SUM_PR_Drop_Staging'               = 238
    'c_UpdateTaskID'                      = 6
    'Course Entry'                        = 0
    'Evaluation'                          = 0
    'Property'                            = 0
    'Skill Entry'                         = 0
    'Compensation Agreement'              = 0
    'Employee Position'                   = 0
    'Stock Option'                        = 0
    'Recruitment'                         = 0
    'Disciplinary Events'                 = 0
    'Skill'                               = 0
    'Level'                               = 0
    'Stock Option Agr'                    = 0
    'Position'                            = 0
    'Vacancy'                             = 0
    'c_OpportunityEntry'                  = 0
    'Event'                               = 0
    'Event Flow'                          = 0
    'Campaign'                            = 0
    'c_NoteLine'                          = 0
    'c_D_FINANCIAL_BILLABILITY V2'        = 0
    'c_WPP Plan Rate'                     = 0
    'c_ContactPerson'                     = 2
    'c_ContactCompany'                    = 6
    'c_Supply values for generated rows'  = 3
    'c_Cleanup Deleted Records Table'     = 0
    'c_Delete KeyMeasures'                = 0
    'c_DWH Deleted Records'               = 0
    'c_Cleanup Deleted Jobs'              = 0
    'c_BudgetEntry'                       = 180
    'c_VendorEntry'                       = 29
    'c_CustomerEntry'                     = 11
    'c_ExpenseSheetLine'                  = 0
    'c_ExpenseSheetHeader'                = 0
    'c_JobProgressEntry'                  = 0
    'c_JobBudgetLine'                     = 0
    'c_JobForecastBudgetEntry'            = 0
    'c_InvoiceAllocationLine'             = 2
    'c_JobInvoiceOnAccount'               = 5
    'c_JobInvoiceOnAccountReconciliation' = 5
    'c_JobInvoiceLine'                    = 600
    'c_JobEntry'                          = 1440
    'c_FinanceEntry'                      = 1800
    'c_KeyMeasures X1'                    = 5
    'c_S_FixedHoursTimeSheet'             = 5
    'c_S_FixedHoursWeekCalendar'          = 6
    'c_Fixed Employee Time'               = 80
    'c_PurchaseOrderLine'                 = 1646
    'c_VendorEntryAccess'                 = 20
    'c_CustomerEntryAccess'               = 3
    'c_Invoice'                           = 4
    'c_JobEntryType'                      = 7
    'c_VendorEntryType'                   = 0
    'c_CustomerEntryType'                 = 0
    'c_TransactionType'                   = 17
    'c_Update_JobBudget'                  = 0
    'c_S_TaskProgressEntry'               = 0
    'c_Budget'                            = 0
    'c_Activity'                          = 0
    'Opportunity'                         = 0
    'c_JobPriceListLine'                  = 1
    'c_DimensionCombinationLine'          = 0
    'c_InvoiceExt_PR'                     = 0
    'c_VendorInvoiceJournal'              = 0
    'c_Task'                              = 19
    'c_JobBudget'                         = 3
    'c_Job'                               = 3
    'c_JobAccess'                         = 0
    'c_Customer'                          = 30
    'c_Parent Customer'                   = 0
    'c_Vendor'                            = 29
    'c_LocalSpecification3'               = 0
    'c_LocalSpecification2'               = 0
    'c_LocalSpecification1'               = 0
    'Specification 3'                     = 0
    'Specification 2'                     = 0
    'Specification 1'                     = 0
    'c_Purpose2'                          = 0
    'Project'                             = 0
    'Location'                            = 0
    'c_PurchaseOrder'                     = 0
    'c_EmployeeRevision V2'               = 0
    'c_Employee'                          = 2
    'c_Entity'                            = 0
    'c_EmployeeCategory'                  = 0
    'c_Local Account'                     = 2
    'c_FiscalPeriod'                      = 0
    'c_Company'                           = 0
    'c_Dimension Grouping'                = 0
    'Forecast Activity Type'              = 0
    'c_Account'                           = 0
    'c_ExchangeRate_PR'                   = 41
    'User Access'                         = 0
    'c_CompanyAccessHierarchy'            = 0
    'c_User'                              = 0
    'c_AdjustCachingOptions'              = 0
    'c_Create GatherStats SP'             = 0
    'c_Date'                              = 0
    'Time'                                = 0
    'Date'                                = 0
}

# ---------------------------------------------------------------------------
# Oracle query execution via sqlplus
# ---------------------------------------------------------------------------
function Invoke-OracleQuery {
    param([string]$Sql)
    $user    = $script:OracleCred.UserName
    $pass    = $script:OracleCred.GetNetworkCredential().Password
    $passSafe = $pass.Replace('"', '""')
    $script = @"
CONNECT $user/"$passSafe"@$($Config.OracleDSN)
SET PAGESIZE 50000
SET LINESIZE 500
SET COLSEP '|'
SET FEEDBACK OFF
SET HEADING ON
SET TRIMSPOOL ON
SET NUMWIDTH 20
$Sql
EXIT;
"@
    try {
        $result = $script | & $Config.SqlPlusPath -S /NOLOG 2>&1
        return $result
    } catch {
        return @("ERROR: $_")
    }
}

function Invoke-OraclePlainQuery {
    param([string]$Sql, [int]$TimeoutSeconds = 30)
    $user     = $script:OracleCred.UserName
    $pass     = $script:OracleCred.GetNetworkCredential().Password
    $passSafe = $pass.Replace('"', '""')
    $script   = @"
CONNECT $user/"$passSafe"@$($Config.OracleDSN)
SET PAGESIZE 50000
SET LINESIZE 200
SET FEEDBACK OFF
SET HEADING OFF
SET TRIMSPOOL ON
$Sql
EXIT;
"@
    try {
        $tempIn  = [System.IO.Path]::GetTempFileName()
        $tempOut = [System.IO.Path]::GetTempFileName()
        [System.IO.File]::WriteAllText($tempIn, $script, [System.Text.UTF8Encoding]::new($false))
        $proc = Start-Process -FilePath $Config.SqlPlusPath -ArgumentList '-S','/NOLOG' `
                    -RedirectStandardInput  $tempIn `
                    -RedirectStandardOutput $tempOut `
                    -NoNewWindow -PassThru
        $proc.WaitForExit($TimeoutSeconds * 1000) | Out-Null
        if (-not $proc.HasExited) { $proc.Kill() ; return 'Query timed out.' }
        $out = Get-Content $tempOut -Raw
        Remove-Item $tempIn, $tempOut -Force -ErrorAction SilentlyContinue
        return $out
    } catch {
        return "ERROR: $_"
    }
}

function ConvertFrom-SqlPlusCsv {
    param([string[]]$Lines)
    $rows = @()
    $headers = $null
    foreach ($line in $Lines) {
        $line = $line.Trim()
        if (-not $line -or $line -match '^(ERROR|ORA-|SP2-)') { continue }
        if (-not $line.Contains('|')) { continue }        # skip "Connected.", etc.
        if ($line -match '^[\s\-\|]+$') { continue }   # skip dashes separator line
        $cols = $line -split '\|'
        if ($null -eq $headers) { $headers = $cols; continue }
        if ($cols.Count -ne $headers.Count) { continue }
        $obj = [ordered]@{}
        for ($i = 0; $i -lt $headers.Count; $i++) {
            $obj[$headers[$i].Trim()] = $cols[$i].Trim()
        }
        $rows += [PSCustomObject]$obj
    }
    return $rows
}

function Format-Duration {
    param([double]$Minutes)
    if ($Minutes -lt 1) { return '< 1m' }
    $d = [math]::Floor($Minutes / 1440)
    $h = [math]::Floor(($Minutes % 1440) / 60)
    $m = [math]::Floor($Minutes % 60)
    $parts = @()
    if ($d -gt 0) { $parts += "${d}d" }
    if ($h -gt 0) { $parts += "${h}h" }
    if ($m -gt 0 -or $parts.Count -eq 0) { $parts += "${m}m" }
    return $parts -join ' '
}

function Format-DiffBadge {
    param([double]$DiffMins)
    if ([math]::Abs($DiffMins) -lt 10) { return '<span class="badge neutral">&#177;10 min</span>' }
    if ($DiffMins -lt 0) {
        return "<span class='badge ahead'>-$(Format-Duration ([math]::Abs($DiffMins))) ahead</span>"
    }
    return "<span class='badge behind'>+$(Format-Duration $DiffMins) behind</span>"
}

# ---------------------------------------------------------------------------
# Data collection
# ---------------------------------------------------------------------------
function Get-EtlData {
    $etlLog = if ($Config.EtlSchema) { "$($Config.EtlSchema).ETL_LOG" } else { "ETL_LOG" }

    # 1. ETL_LOG — all steps for current batch, newest first
    $logSql = @"
SELECT ID_BATCH, TRANSNAME, STATUS, ERRORS,
       TO_CHAR(STARTDATE,'YYYY-MM-DD HH24:MI:SS') AS STARTDATE,
       TO_CHAR(ENDDATE,'YYYY-MM-DD HH24:MI:SS') AS ENDDATE,
       TO_CHAR(LOGDATE,'YYYY-MM-DD HH24:MI:SS') AS LOGDATE
FROM $etlLog
WHERE ID_BATCH >= (SELECT MAX(ID_BATCH) - 300 FROM $etlLog)
ORDER BY ID_BATCH DESC;
"@
    $rawLog = Invoke-OracleQuery $logSql
    Write-Host "--- RAW ETL_LOG OUTPUT ---"
    $rawLog | Select-Object -First 10 | ForEach-Object { Write-Host $_ }
    Write-Host "--- END RAW OUTPUT ---"
    $logRows = @(ConvertFrom-SqlPlusCsv $rawLog)
    Write-Host "ETL_LOG rows returned: $($logRows.Count)"

    # 2. v$session_longops — any currently active long-running operation
    $longopsSql = @"
SELECT sl.sql_id, sl.sid, sl.opname, sl.target,
       sl.sofar, sl.totalwork,
       ROUND(sl.sofar / NULLIF(sl.totalwork, 0) * 100, 2) AS pct_complete,
       sl.elapsed_seconds,
       sl.time_remaining,
       TO_CHAR(SYSDATE + NUMTODSINTERVAL(sl.time_remaining, 'SECOND'),
               'DD-MON-YYYY HH24:MI:SS') AS estimated_finish
FROM v`$session_longops sl
WHERE sl.totalwork > 0
  AND sl.sofar < sl.totalwork
  AND sl.elapsed_seconds > 60
ORDER BY sl.elapsed_seconds DESC;
"@
    $rawLongops = Invoke-OracleQuery $longopsSql
    Write-Host "--- RAW LONGOPS OUTPUT ---"
    $rawLongops | ForEach-Object { Write-Host $_ }
    Write-Host "--- END LONGOPS OUTPUT ---"
    $longops = @(ConvertFrom-SqlPlusCsv $rawLongops)
    Write-Host "Longops rows parsed: $($longops.Count)"


    # 3. Index presence, validity and visibility check
    $idxSql = @"
WITH expected AS (
  SELECT 'IDX_DBA1_C_D_AUDIT' AS expected_name, 'C_D_AUDIT' AS table_name, 'LOAD_END_TIME,LOAD_START_TIME' AS cols FROM dual UNION ALL
  SELECT 'IDX_DBA2_EXCUST', 'EXCUSTOMERRECONCILIATIONJOURNA', 'ORIGINALRECONCILIATIONNUMBER,POSTEDPN' FROM dual UNION ALL
  SELECT 'IDX_DBA3_EXCUSTENTRY', 'EXCUSTOMERENTRY', 'CUSTOMERNUMBER,LINENUMBER' FROM dual UNION ALL
  SELECT 'IDX_DBA4_CUSTOMER', 'CUSTOMER', 'CUSTOMERNUMBER' FROM dual UNION ALL
  SELECT 'IDX_DBA5_JOBHEADER', 'JOBHEADER', 'JOBNUMBER' FROM dual UNION ALL
  SELECT 'IDX_DBA6_EXINVOICE', 'EXINVOICE', 'INVOICENUMBER' FROM dual UNION ALL
  SELECT 'IDX_DBA7_EXJOBINVONACCTRECON', 'EXJOBINVOICEONACCOUNTRECONCILI', 'INVOICENUMBER,JOBNUMBER,JOBINVOICEONACCOUNTINSTANCEK' FROM dual UNION ALL
  SELECT 'IDX_DBA8_ECONJRNAL', 'EXVENDORRECONCILIATIONJOURNAL', 'VENDORNUMBER,MAINENTRYLINENUMBER' FROM dual UNION ALL
  SELECT 'IDX_DBA9_VENDORNO', 'VENDOR', 'VENDORNUMBER' FROM dual
),
actual AS (
  SELECT ic.index_owner, ic.table_name, ic.index_name,
         LISTAGG(ic.column_name, ',') WITHIN GROUP (ORDER BY ic.column_position) AS cols
  FROM   all_ind_columns ic
  WHERE  ic.table_name IN (SELECT table_name FROM expected)
    AND  ic.index_owner = '$($Config.EtlSchema)'
  GROUP  BY ic.index_owner, ic.table_name, ic.index_name
)
SELECT e.expected_name,
       e.table_name,
       a.index_name AS actual_index_name,
       ai.status    AS idx_status,
       ai.visibility,
       CASE
         WHEN a.index_name IS NULL           THEN 'MISSING'
         WHEN ai.status <> 'VALID'           THEN 'INVALID'
         WHEN ai.visibility <> 'VISIBLE'     THEN 'INVISIBLE'
         WHEN a.index_name = e.expected_name THEN 'OK'
         ELSE                                     'OK (renamed)'
       END AS status
FROM   expected e
LEFT JOIN actual a
       ON a.table_name = e.table_name AND a.cols = e.cols
LEFT JOIN all_indexes ai
       ON ai.owner = a.index_owner AND ai.index_name = a.index_name
ORDER  BY e.expected_name;
"@
    $indexCheck = @(ConvertFrom-SqlPlusCsv (Invoke-OracleQuery $idxSql))

    return @{ Log = $logRows; LongOps = $longops; IndexCheck = $indexCheck }
}

# ---------------------------------------------------------------------------
# HTML generation
# ---------------------------------------------------------------------------
function Build-Html {
    param($Data)

    $now        = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $logRows    = $Data.Log
    $longops    = $Data.LongOps
    $indexCheck = @($Data.IndexCheck)

    $idxAllOk  = ($indexCheck.Count -gt 0) -and ($indexCheck | Where-Object { $_.STATUS -notmatch '^OK' }).Count -eq 0
    $idxBadgeCls  = if ($idxAllOk) { 'idx-badge-ok' } else { 'idx-badge-warn' }
    $idxBadgeTxt  = if ($idxAllOk) { 'Indexes OK' } else { 'Index Issue' }
    $idxRows = ($indexCheck | ForEach-Object {
        $rowCls = if ($_.STATUS -match '^OK') { '' } else { ' class="idx-missing"' }
        "<tr$rowCls><td>$($_.expected_name)</td><td>$($_.table_name)</td><td>$($_.STATUS)</td></tr>"
    }) -join ''
    $idxPopupHtml = "<table class='idx-table'><thead><tr><th>Expected Index</th><th>Table</th><th>Status</th></tr></thead><tbody>$idxRows</tbody></table>"

    # Log buttons
    $logButtonsHtml = @('PDI', 'Load') | ForEach-Object {
        $key    = $_
        $label  = if ($key -eq 'PDI') { 'PDI Log' } else { 'Load Log' }
        $path   = $script:LogPaths[$key]
        if ($path -and (Test-Path $path)) {
            $tail    = (Get-LogTail $path 30) -replace '&','&amp;' -replace '<','&lt;' -replace '>','&gt;'
            $tooltip = "<div class='log-tooltip'>$tail</div>"
            "<a class='log-btn' href='/setlogpath?log=$key' onclick='event.preventDefault();fetch(`"/setlogpath?log=$key`")'>$label $tooltip</a>"
        } else {
            "<a class='log-btn disabled' href='/setlogpath?log=$key' onclick='event.preventDefault();fetch(`"/setlogpath?log=$key`").then(()=>location.reload())' title='Click to set log path'>$label</a>"
        }
    }
    $logButtonsHtml = $logButtonsHtml -join ' '

    # ID_BATCH -> LOGDATE map for duration calculations (STARTDATE/ENDDATE are unreliable)
    $logDateById = @{}
    foreach ($row in @($logRows)) {
        if ($row.ID_BATCH) { $logDateById[[int]$row.ID_BATCH] = $row.LOGDATE }
    }

    # Identify current (running), failed, and finished steps
    $running  = $logRows | Where-Object { $_.STATUS -ieq 'running' -or $_.STATUS -ieq 'start' } | Select-Object -First 1
    $failed   = $logRows | Where-Object { $_.STATUS -ieq 'stop' } | Select-Object -First 1
    $finished = $logRows | Where-Object { $_.STATUS -ieq 'end' -or $_.STATUS -ieq 'finish' }

    $finishedCount = ($finished | Measure-Object).Count
    $totalSteps    = $Benchmark.Count
    $pct           = if ($totalSteps -gt 0) { [math]::Round($finishedCount / $totalSteps * 100, 1) } else { 0 }

    # ETA calculation
    $etaHtml = '--'
    $etaLabel = 'ETA Remaining'
    $paceHtml = ''
    $benchFinished = ($finished | ForEach-Object {
        if ($Benchmark.ContainsKey($_.TRANSNAME)) { $Benchmark[$_.TRANSNAME] } else { 0 }
    } | Measure-Object -Sum).Sum

    if ($benchFinished -gt 0 -and $finished.Count -gt 0) {
        # Sum actual minutes for finished steps using LOGDATE diff (STARTDATE/ENDDATE unreliable)
        $actualFinishedMins = ($finished | ForEach-Object {
            $sid = [int]$_.ID_BATCH
            $prevLd = $logDateById[($sid - 1)]
            if ($prevLd) {
                try {
                    $s = [datetime]::Parse($prevLd)
                    $e = [datetime]::Parse($_.LOGDATE)
                    [math]::Max(0, ($e - $s).TotalMinutes)
                } catch { 0 }
            } else { 0 }
        } | Measure-Object -Sum).Sum

        $pace = $actualFinishedMins / $benchFinished
        $paceHtml = "(pace: $([math]::Round($pace,2))&times; benchmark)"

        $isRunning = ($null -ne $running) -or ($null -ne $failed)
        if ($isRunning) {
            $remainingBench = ($Benchmark.Values | Measure-Object -Sum).Sum - $benchFinished
            $etaMins  = $remainingBench * $pace
            $etaHtml  = '~' + (Format-Duration $etaMins)
            $etaLabel = 'ETA Remaining'
        } else {
            $totalBench = ($Benchmark.Values | Measure-Object -Sum).Sum
            $diffMins   = $actualFinishedMins - $totalBench
            $etaHtml    = Format-DiffBadge $diffMins
            $etaLabel   = 'Against Benchmark'
        }
    }

    # Elapsed since batch start — use LOGDATE of earliest step (STARTDATE is unreliable)
    $batchElapsedHtml = '--'
    $batchStart = $logRows | Select-Object -Last 1
    if ($batchStart) {
        try {
            $start = [datetime]::Parse($batchStart.LOGDATE)
            $elapsedMins = ([datetime]::Now - $start).TotalMinutes
            $batchElapsedHtml = Format-Duration $elapsedMins
        } catch {}
    }

    # LRQ panel — built independently so it shows even if ETL_LOG has no running status
    $lrqHtml = ''
    if ($longops.Count -gt 0) {
        $op     = $longops[0]
        $pctVal = [double]$op.pct_complete
        $barW   = [math]::Round($pctVal)
        $elStr  = Format-Duration ([double]$op.elapsed_seconds / 60)
        $remStr = Format-Duration ([double]$op.time_remaining / 60)
        $sofar  = [int64]$op.sofar
        $total  = [int64]$op.totalwork
        $opDesc = if ($op.target) {
            "$(HtmlEncode($op.opname)) on $(HtmlEncode($op.target))"
        } else {
            HtmlEncode($op.opname)
        }
        $bmMinsForLrq = if ($running -and $Benchmark.ContainsKey($running.TRANSNAME)) { $Benchmark[$running.TRANSNAME] } else { $null }
        $slowWarning = ''
        $barColour = 'var(--green)'
        if ($null -ne $bmMinsForLrq -and $bmMinsForLrq -gt 0) {
            $projectedMins = ([double]$op.elapsed_seconds / 60) + ([double]$op.time_remaining / 60)
            if ($projectedMins -gt $bmMinsForLrq) {
                $barColour = 'var(--red)'
                $overBy = Format-Duration ($projectedMins - $bmMinsForLrq)
                $slowWarning = "<div class='lrq-warning'>&#9888; Projected finish is $overBy over benchmark</div>"
            } elseif ($projectedMins -gt $bmMinsForLrq * 0.8) {
                $barColour = 'var(--amber)'
            }
        }
        $lrqHtml = @"
<div class="lrq-panel">
  <div class="lrq-title">Active Query (LRQ)</div>
  <div class="lrq-opname">$opDesc</div>
  <div class="lrq-pct-row">
    <span class="lrq-pct-num">${pctVal}%</span>
    <span class="muted" style="font-size:12px">$("{0:N0}" -f $sofar) / $("{0:N0}" -f $total) rows</span>
  </div>
  <div class="progress-bar-bg" style="margin:6px 0 10px">
    <div class="progress-bar-fill" style="width:${barW}%;background:${barColour}"></div>
  </div>
  <table class="info-table">
    <tr><th>Elapsed</th><td>$elStr</td></tr>
    <tr><th>Remaining</th><td>$remStr</td></tr>
    <tr><th>ETA</th><td>$($op.estimated_finish)</td></tr>
    <tr><th>SQL ID</th><td class="mono muted" style="font-size:11px">$($op.sql_id)</td></tr>
  </table>
  $slowWarning
  $(if ($op.sql_id) {
    "<button class='exec-plan-btn' onclick='copyExecPlan(`"$($op.sql_id)`",this)'>Copy Execution Plan</button>"
  })
</div>
"@
    }

    # Current step card
    $currentHtml = '<p class="muted">No step currently running.</p>'
    if ($running) {
        $bmMins = if ($Benchmark.ContainsKey($running.TRANSNAME)) { $Benchmark[$running.TRANSNAME] } else { $null }
        $elapsedMins = 0
        $runStartDisplay = 'Unknown'
        $runningId = [int]$running.ID_BATCH
        $prevStep = $logRows | Where-Object { [int]$_.ID_BATCH -eq ($runningId - 1) } | Select-Object -First 1
        if ($prevStep) {
            try {
                $s = [datetime]::Parse($prevStep.LOGDATE)
                $elapsedMins = ([datetime]::Now - $s).TotalMinutes
                $runStartDisplay = $prevStep.LOGDATE
            } catch {}
        }
        $elapsedStr = Format-Duration $elapsedMins
        $bmStr  = if ($null -ne $bmMins) { Format-Duration $bmMins } else { 'N/A' }
        $diffHtml = if ($null -ne $bmMins -and $bmMins -gt 0) {
            Format-DiffBadge ($elapsedMins - $bmMins)
        } elseif ($null -ne $bmMins) {
            '<span class="badge neutral">&#177;0m</span>'
        } else {
            '<span class="badge neutral">N/A</span>'
        }
        $currentStepTable = @"
<table class="info-table">
  <tr><th>Step</th><td><strong>$(HtmlEncode($running.TRANSNAME))</strong></td></tr>
  <tr><th>Status</th><td><span class="status-running">&#9679; RUNNING</span></td></tr>
  <tr><th>Started</th><td>$runStartDisplay</td></tr>
  <tr><th>Elapsed</th><td>$elapsedStr</td></tr>
  <tr><th>Benchmark</th><td>$bmStr</td></tr>
  <tr><th>vs Benchmark</th><td>$diffHtml</td></tr>
</table>
"@
        $currentHtml = if ($lrqHtml) {
            "<div class='current-step-grid'>$currentStepTable$lrqHtml</div>"
        } else {
            $currentStepTable
        }
    } elseif ($lrqHtml) {
        # Step not detected in ETL_LOG but Oracle reports an active LRQ — show it anyway
        $currentHtml = $lrqHtml
    } elseif ($failed) {
        $errCount = $failed.ERRORS
        $failedAt = $failed.LOGDATE
        $currentHtml = @"
<table class="info-table">
  <tr><th>Step</th><td><strong>$(HtmlEncode($failed.TRANSNAME))</strong></td></tr>
  <tr><th>Status</th><td><span class="status-failed">&#9679; FAILED</span></td></tr>
  <tr><th>Errors</th><td>$errCount error(s)</td></tr>
  <tr><th>Stopped at</th><td>$failedAt</td></tr>
</table>
"@
    }

    # Recent steps table — all rows, split into visible (8) and hidden (rest)
    $recentRows = $finished
    $recentHtml = '<p class="muted">No completed steps yet.</p>'
    if ($recentRows.Count -gt 0) {
        $allRows = $recentRows | ForEach-Object {
            $dur = '--'
            $diffBadge = ''
            $startDisplay = '--'
            $endDisplay = '--'
            $sid = [int]$_.ID_BATCH
            $prevLd = $logDateById[($sid - 1)]
            if ($prevLd) {
                try {
                    $s = [datetime]::Parse($prevLd)
                    $e = [datetime]::Parse($_.LOGDATE)
                    $durMins = ($e - $s).TotalMinutes
                    $dur = Format-Duration $durMins
                    $startDisplay = $s.ToString('dd-MMM HH:mm')
                    $endDisplay   = $e.ToString('dd-MMM HH:mm')
                    if ($Benchmark.ContainsKey($_.TRANSNAME)) {
                        $diffBadge = Format-DiffBadge ($durMins - $Benchmark[$_.TRANSNAME])
                    }
                } catch {}
            }
            $errBadge = if ([int]$_.ERRORS -gt 0) { "<span class='badge behind'>$($_.ERRORS) err</span>" } else { '' }
            [PSCustomObject]@{ Html = "<tr>
              <td class='muted' style='white-space:nowrap'>$startDisplay</td>
              <td class='muted' style='white-space:nowrap'>$endDisplay</td>
              <td>$(HtmlEncode($_.TRANSNAME))</td>
              <td><span class='status-done'>&#10003;</span></td>
              <td>$dur</td>
              <td>$diffBadge</td>
              <td class='muted'>$errBadge</td>
            </tr>" }
        }
        $visibleRows = ($allRows | Select-Object -First 8).Html -join ''
        $hiddenRows  = ($allRows | Select-Object -Skip 8).Html -join ''
        $hiddenSection = if ($hiddenRows) {
            "<tbody id='steps-extra' style='display:none'>$hiddenRows</tbody>
            <tfoot><tr><td colspan='7' style='padding:10px 0 0'>
              <button class='expand-btn' onclick='toggleSteps(this)'>Show all $($recentRows.Count) steps</button>
            </td></tr></tfoot>"
        } else { '' }
        $recentHtml = @"
<table class="data-table" id="steps-table">
  <thead><tr><th>Started</th><th>Finished</th><th>Step</th><th>Status</th><th>Duration</th><th>vs Benchmark</th><th></th></tr></thead>
  <tbody id="steps-visible">$visibleRows</tbody>
  $hiddenSection
</table>
"@
    }

    return @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta http-equiv="refresh" content="$($Config.RefreshSeconds)">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>ETL Monitor</title>
<style>
  :root {
    --bg: #f5f6f8; --card: #fff; --border: #e2e5ea;
    --text: #1a1d23; --muted: #6b7280;
    --green: #16a34a; --red: #dc2626; --amber: #d97706; --blue: #2563eb;
    --green-bg: #dcfce7; --red-bg: #fee2e2; --amber-bg: #fef3c7;
  }
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body { font-family: system-ui, -apple-system, sans-serif; background: var(--bg);
         color: var(--text); font-size: 14px; padding: 16px; }
  header { display: flex; justify-content: space-between; align-items: baseline;
           margin-bottom: 20px; flex-wrap: wrap; gap: 8px; }
  header h1 { font-size: 20px; font-weight: 700; }
  .meta { color: var(--muted); font-size: 13px; }
  .cards { display: grid; grid-template-columns: repeat(auto-fit, minmax(160px,1fr));
           gap: 12px; margin-bottom: 20px; }
  .card { background: var(--card); border: 1px solid var(--border); border-radius: 8px;
          padding: 14px 16px; }
  .card-label { font-size: 11px; text-transform: uppercase; letter-spacing: .05em;
                color: var(--muted); margin-bottom: 4px; }
  .card-value { font-size: 22px; font-weight: 700; }
  .progress-wrap { background: var(--card); border: 1px solid var(--border);
                   border-radius: 8px; padding: 14px 16px; margin-bottom: 20px; }
  .progress-label { display: flex; justify-content: space-between; font-size: 13px;
                    color: var(--muted); margin-bottom: 8px; }
  .progress-bar-bg { background: #e5e7eb; border-radius: 99px; height: 12px; overflow: hidden; }
  .progress-bar-fill { background: var(--blue); height: 100%; border-radius: 99px;
                       transition: width .3s; }
  .progress-detail { font-size: 12px; color: var(--muted); margin-top: 6px; }
  section { background: var(--card); border: 1px solid var(--border); border-radius: 8px;
            padding: 16px; margin-bottom: 16px; }
  section h2 { font-size: 14px; font-weight: 600; margin-bottom: 12px;
               padding-bottom: 8px; border-bottom: 1px solid var(--border); }
  .info-table { width: 100%; border-collapse: collapse; }
  .info-table th { width: 130px; text-align: left; color: var(--muted);
                   font-weight: 500; padding: 5px 0; vertical-align: top; }
  .info-table td { padding: 5px 0 5px 12px; }
  .data-table { width: 100%; border-collapse: collapse; font-size: 13px; }
  .data-table th { text-align: left; color: var(--muted); font-weight: 500;
                   font-size: 11px; text-transform: uppercase; letter-spacing: .05em;
                   padding: 0 8px 8px 0; border-bottom: 1px solid var(--border); }
  .data-table td { padding: 7px 8px 7px 0; border-bottom: 1px solid #f0f2f5;
                   vertical-align: middle; }
  .data-table tr:last-child td { border-bottom: none; }
  .badge { display: inline-block; padding: 2px 7px; border-radius: 4px; font-size: 12px; font-weight: 500; }
  .ahead   { background: var(--green-bg); color: var(--green); }
  .behind  { background: var(--red-bg);   color: var(--red); }
  .neutral { background: #f3f4f6; color: var(--muted); }
  .status-running { color: var(--amber); font-weight: 600; }
  .status-done    { color: var(--green); }
  .muted { color: var(--muted); }
  .mono  { font-family: monospace; font-size: 12px; }
  .mini-bar-wrap { display: inline-block; width: 80px; height: 8px; background: #e5e7eb;
                   border-radius: 99px; overflow: hidden; vertical-align: middle;
                   margin-right: 6px; }
  .mini-bar { height: 100%; background: var(--blue); border-radius: 99px; }
  .longop-card { background: var(--card); border: 1px solid var(--border); border-radius: 8px;
                 padding: 12px 16px; margin-bottom: 16px; }
  .longop-header { display: flex; justify-content: space-between; align-items: baseline; }
  .longop-name { font-size: 13px; font-weight: 500; }
  .longop-pct { font-size: 20px; font-weight: 700; color: var(--amber); }
  .longop-stats { display: flex; gap: 24px; font-size: 12px; color: var(--muted); flex-wrap: wrap; }
  .longop-stats strong { color: var(--text); }
  .idx-wrap { position: relative; display: inline-block; }
  .idx-badge-ok, .idx-badge-warn {
    display: inline-block; padding: 3px 10px; border-radius: 4px;
    font-size: 12px; font-weight: 600; cursor: default; }
  .idx-badge-ok   { background: var(--green-bg); color: var(--green); }
  .idx-badge-warn { background: var(--red-bg);   color: var(--red); }
  .idx-popup { display: none; position: absolute; right: 0; top: calc(100% + 6px);
               background: var(--card); border: 1px solid var(--border);
               border-radius: 8px; box-shadow: 0 4px 16px rgba(0,0,0,.12);
               z-index: 100; min-width: 480px; padding: 12px; }
  .idx-wrap:hover .idx-popup { display: block; }
  .idx-table { width: 100%; border-collapse: collapse; font-size: 12px; }
  .idx-table th { text-align: left; color: var(--muted); font-weight: 500;
                  font-size: 11px; text-transform: uppercase; letter-spacing: .05em;
                  padding: 0 8px 6px 0; border-bottom: 1px solid var(--border); }
  .idx-table td { padding: 5px 8px 5px 0; border-bottom: 1px solid #f0f2f5; }
  .idx-table tr:last-child td { border-bottom: none; }
  .idx-missing td { color: var(--red); }
  .current-step-grid { display:grid; grid-template-columns:1fr 1fr; gap:20px; align-items:start; }
  .lrq-panel { background:var(--bg); border:1px solid var(--border); border-radius:8px; padding:14px 16px; }
  .lrq-title { font-size:11px; text-transform:uppercase; letter-spacing:.06em; color:var(--muted); margin-bottom:6px; }
  .lrq-opname { font-weight:600; font-size:14px; margin-bottom:8px; word-break:break-word; }
  .lrq-pct-row { display:flex; align-items:baseline; gap:10px; }
  .lrq-pct-num { font-size:28px; font-weight:700; line-height:1; }
  .status-failed { color:var(--red); font-weight:700; }
  .lrq-warning { margin-top:10px; padding:7px 10px; background:#fef9ec; border:1px solid #f5d87a;
                 border-radius:6px; color:#92610a; font-size:13px; }
  .log-btn { position:relative; display:inline-block; padding:4px 10px; border-radius:12px;
             font-size:12px; font-weight:600; cursor:pointer; border:none; text-decoration:none;
             background:var(--blue); color:#fff; }
  .log-btn.disabled { background:var(--border); color:var(--muted); cursor:pointer; }
  .log-btn .log-tooltip { display:none; position:absolute; top:110%; right:0; z-index:100;
             background:#1a1d23; color:#e5e7eb; font-family:monospace; font-size:11px;
             white-space:pre; padding:10px 12px; border-radius:6px; min-width:520px;
             max-width:700px; max-height:340px; overflow-y:auto;
             box-shadow:0 4px 16px rgba(0,0,0,.35); line-height:1.5; }
  .log-btn:hover .log-tooltip { display:block; }
  .exec-plan-btn { margin-top:10px; padding:5px 12px; border-radius:6px; border:none;
                   background:var(--blue); color:#fff; font-size:12px; font-weight:600;
                   cursor:pointer; }
  .exec-plan-btn:hover { opacity:.85; }
  .exec-plan-btn.copied { background:var(--green); }
  .expand-btn { padding:5px 14px; border-radius:6px; border:1px solid var(--border);
                background:var(--card); color:var(--text); font-size:12px; cursor:pointer; }
  .expand-btn:hover { background:var(--bg); }
  .pdf-btn { padding:4px 12px; border-radius:12px; border:none; background:var(--blue);
             color:#fff; font-size:12px; font-weight:600; cursor:pointer; }
  .pdf-btn:hover { opacity:.85; }
  @media print {
    .no-print, .log-btn, .pdf-btn, .expand-btn, .idx-wrap,
    meta[http-equiv="refresh"] { display:none !important; }
    .print-only { display:inline !important; }
    body { background:#fff; }
    header { border-bottom:2px solid #000; padding-bottom:8px; margin-bottom:16px; }
    .cards, .progress-wrap { page-break-inside:avoid; }
    section:first-of-type { page-break-after:auto; }
    #steps-extra { display:table-row-group !important; }
    #steps-table tr { page-break-inside:avoid; }
    .data-table th, .data-table td { font-size:10px; padding:4px 6px; }
    .badge { font-size:9px; padding:1px 4px; }
    tfoot { display:none; }
  }
</style>
</head>
<body>

<header>
  <div style="font-size:11px;text-transform:uppercase;letter-spacing:.08em;color:var(--muted);margin-bottom:2px">$selectedEnv - $($Environments[$selectedEnv].Label)</div>
  <h1>ETL Monitor</h1>
  <div class="meta" style="display:flex;align-items:center;gap:12px;flex-wrap:wrap">
    <span class="no-print">Last refresh: $now &nbsp;|&nbsp; Auto-refresh: every $($Config.RefreshSeconds)s</span>
    <span class="print-only" style="display:none">Report generated: $now</span>
    $logButtonsHtml
    <button class="pdf-btn no-print" onclick="exportPdf()">Export PDF</button>
    <div class="idx-wrap">
      <span class="$idxBadgeCls">$idxBadgeTxt</span>
      <div class="idx-popup">$idxPopupHtml</div>
    </div>
  </div>
</header>

<div class="cards">
  <div class="card">
    <div class="card-label">Steps Complete</div>
    <div class="card-value">$finishedCount <span style="font-size:14px;color:var(--muted)">/ $totalSteps</span></div>
  </div>
  <div class="card">
    <div class="card-label">Progress</div>
    <div class="card-value">$pct<span style="font-size:14px">%</span></div>
  </div>
  <div class="card">
    <div class="card-label">Elapsed</div>
    <div class="card-value" style="font-size:16px">$batchElapsedHtml</div>
  </div>
  <div class="card">
    <div class="card-label">$etaLabel</div>
    <div class="card-value" style="font-size:16px">$etaHtml</div>
  </div>
</div>

<div class="progress-wrap">
  <div class="progress-label">
    <span>Overall Progress</span>
    <span>$pct%</span>
  </div>
  <div class="progress-bar-bg">
    <div class="progress-bar-fill" style="width:${pct}%"></div>
  </div>
  <div class="progress-detail">$paceHtml</div>
</div>

<section>
  <h2>Current Step</h2>
  $currentHtml
</section>

<section>
  <h2>Recent Steps (last 20)</h2>
  $recentHtml
</section>

<script>
function copyExecPlan(sqlId, btn) {
  btn.textContent = 'Fetching...';
  fetch('/execplan?sql_id=' + encodeURIComponent(sqlId))
    .then(r => r.text())
    .then(t => {
      navigator.clipboard.writeText(t).then(() => {
        btn.textContent = 'Copied!';
        btn.classList.add('copied');
        setTimeout(() => { btn.textContent = 'Copy Execution Plan'; btn.classList.remove('copied'); }, 2000);
      });
    })
    .catch(() => { btn.textContent = 'Error'; });
}

function toggleSteps(btn) {
  var extra = document.getElementById('steps-extra');
  var hidden = extra.style.display === 'none';
  extra.style.display = hidden ? '' : 'none';
  btn.textContent = hidden ? 'Show fewer steps' : btn.textContent.replace('Show fewer', 'Show all');
}

function exportPdf() {
  // Expand all steps before printing
  var extra = document.getElementById('steps-extra');
  var wasHidden = extra && extra.style.display === 'none';
  if (wasHidden) extra.style.display = '';
  window.print();
  if (wasHidden) extra.style.display = 'none';
}
</script>
</body>
</html>
"@
}

# ---------------------------------------------------------------------------
# Log paths — stored in paths.json, keyed by environment
# ---------------------------------------------------------------------------
$pathsFile = Join-Path $PSScriptRoot 'paths.json'
$script:LogPaths = @{ PDI = ''; Load = '' }
if (Test-Path $pathsFile) {
    try {
        $allPaths = Get-Content $pathsFile -Raw | ConvertFrom-Json
        if ($allPaths.$selectedEnv) {
            $ep = $allPaths.$selectedEnv
            $script:LogPaths.PDI  = if ($ep.PDI)  { $ep.PDI }  else { '' }
            $script:LogPaths.Load = if ($ep.Load) { $ep.Load } else { '' }
        }
    } catch {}
}

function Save-LogPath {
    param([string]$LogKey, [string]$Path)
    # Read existing JSON as PSCustomObject then convert to nested hashtables (PS 5.1 compatible)
    $allPaths = @{}
    if (Test-Path $pathsFile) {
        try {
            $json = Get-Content $pathsFile -Raw | ConvertFrom-Json
            foreach ($env in $json.PSObject.Properties) {
                $allPaths[$env.Name] = @{}
                foreach ($key in $env.Value.PSObject.Properties) {
                    $allPaths[$env.Name][$key.Name] = $key.Value
                }
            }
        } catch {}
    }
    if (-not $allPaths[$selectedEnv]) { $allPaths[$selectedEnv] = @{} }
    $allPaths[$selectedEnv][$LogKey] = $Path
    $allPaths | ConvertTo-Json -Depth 4 | Set-Content $pathsFile
    $script:LogPaths[$LogKey] = $Path
}

function Open-FilePicker {
    param([string]$Title)
    Add-Type -AssemblyName System.Windows.Forms
    # Invisible top-most owner forces dialog to appear in front of the browser
    $owner = [System.Windows.Forms.Form]::new()
    $owner.TopMost = $true
    $owner.StartPosition = [System.Windows.Forms.FormStartPosition]::Manual
    $owner.Location = [System.Drawing.Point]::new(-32000, -32000)
    $owner.Size = [System.Drawing.Size]::new(1, 1)
    $owner.Show()
    $dlg = [System.Windows.Forms.OpenFileDialog]::new()
    $dlg.Title  = $Title
    $dlg.Filter = 'Log files (*.log;*.txt)|*.log;*.txt|All files (*.*)|*.*'
    $result = $dlg.ShowDialog($owner)
    $owner.Dispose()
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) { return $dlg.FileName } else { return '' }
}

function Get-LogTail {
    param([string]$Path, [int]$Lines = 30)
    if (-not $Path -or -not (Test-Path $Path)) { return $null }
    try {
        return (Get-Content $Path -Tail $Lines -ErrorAction Stop) -join "`n"
    } catch { return $null }
}

# ---------------------------------------------------------------------------
# Credentials — encrypted with Windows DPAPI, tied to this user/machine
# ---------------------------------------------------------------------------
$credFile = Join-Path $env:APPDATA "ETLpulse\oracle.$selectedEnv.cred"
if (-not (Test-Path (Split-Path $credFile))) { New-Item -ItemType Directory -Path (Split-Path $credFile) | Out-Null }

if (Test-Path $credFile) {
    try {
        $script:OracleCred = Import-Clixml $credFile
        Write-Host "Loaded saved credentials for: $($script:OracleCred.UserName)"
    } catch {
        Write-Warning "Saved credentials could not be decrypted (wrong user/machine). Re-entering credentials."
        Remove-Item $credFile -Force
        $script:OracleCred = $null
    }
}
if (-not $script:OracleCred) {
    $script:OracleCred = Get-Credential -Message "Enter Oracle credentials for $selectedEnv ($($Config.OracleDSN))"
    if (-not $script:OracleCred) { Write-Error "No credentials provided."; exit 1 }
    $script:OracleCred | Export-Clixml $credFile
    Write-Host "Credentials saved (encrypted) to $credFile"
    Write-Host "NOTE: only this Windows account on this machine can decrypt them."
}

Write-Host "Testing Oracle connection..."
$testPass   = $script:OracleCred.GetNetworkCredential().Password
$testSafe   = $testPass.Replace('"', '""')
$testUser   = $script:OracleCred.UserName
$testDSN    = $Config.OracleDSN
$testScript = @"
CONNECT $testUser/"$testSafe"@$testDSN
SELECT 'connected' FROM dual;
EXIT;
"@
$testResult = $testScript | & $Config.SqlPlusPath -S /NOLOG 2>&1
Write-Host "sqlplus response: $($testResult -join ' | ')"
if ($testResult -match 'ORA-|SP2-|TNS-') {
    Write-Warning "Connection failed. To re-enter credentials, delete the saved file and restart:"
    Write-Warning "  Remove-Item '$credFile'"
}

# ---------------------------------------------------------------------------
# HTTP listener — serves dashboard on every GET request
# ---------------------------------------------------------------------------
function HtmlEncode([string]$s) {
    $s.Replace('&','&amp;').Replace('<','&lt;').Replace('>','&gt;').Replace('"','&quot;').Replace("'",'&#39;')
}

$listener = [System.Net.HttpListener]::new()
$listener.Prefixes.Add("http://+:$($Config.Port)/")

try {
    $listener.Start()
} catch {
    Write-Error "Cannot start listener on port $($Config.Port): $_"
    Write-Host "Try running PowerShell as Administrator, or change the Port in config.ps1."
    exit 1
}

$hostname = [System.Net.Dns]::GetHostName()
$dashboardUrl = "http://localhost:$($Config.Port)/"
Write-Host "ETL Monitor started."
Write-Host "Dashboard: http://${hostname}:$($Config.Port)/"
Write-Host "Press Ctrl+C to stop.`n"
$chromePaths = @(
    "$env:ProgramFiles\Google\Chrome\Application\chrome.exe",
    "${env:ProgramFiles(x86)}\Google\Chrome\Application\chrome.exe",
    "$env:LocalAppData\Google\Chrome\Application\chrome.exe"
)
$chrome = $chromePaths | Where-Object { Test-Path $_ } | Select-Object -First 1
if ($chrome) {
    Start-Process $chrome $dashboardUrl
} else {
    Start-Process $dashboardUrl  # fallback to default browser
}

try {
    while ($listener.IsListening) {
        $ctx = $listener.GetContext()
        $req = $ctx.Request
        $url = $req.Url.AbsolutePath
        Write-Host "$(Get-Date -Format 'HH:mm:ss')  $($req.HttpMethod) $url"

        $responseText = $null
        $contentType  = 'text/plain; charset=utf-8'

        if ($url -eq '/setlogpath') {
            $logKey = $req.QueryString['log']
            # Respond immediately so the browser doesn't time out, then open the dialog
            $responseText = 'ok'
            $bytes = [System.Text.Encoding]::UTF8.GetBytes($responseText)
            $ctx.Response.ContentType     = 'text/plain; charset=utf-8'
            $ctx.Response.ContentLength64 = $bytes.LongLength
            $ctx.Response.OutputStream.Write($bytes, 0, $bytes.Length)
            $ctx.Response.OutputStream.Close()
            if ($logKey -eq 'PDI' -or $logKey -eq 'Load') {
                $label  = if ($logKey -eq 'PDI') { 'PDI Log' } else { 'Load Log' }
                $picked = Open-FilePicker -Title "Select $label file for $selectedEnv"
                if ($picked) {
                    Save-LogPath -LogKey $logKey -Path $picked
                    Write-Host "Saved $logKey log path: $picked"
                }
            }
            continue

        } elseif ($url -eq '/logtail') {
            $logKey = $req.QueryString['log']
            $path   = $script:LogPaths[$logKey]
            $tail   = Get-LogTail $path 30
            $responseText = if ($tail) { $tail } else { '(no data)' }

        } elseif ($url -eq '/execplan') {
            $sqlId = $req.QueryString['sql_id']
            if ($sqlId -match '^[a-z0-9]+$') {
                $planSql = @"
SELECT operation || ' ' || options || CASE WHEN object_name IS NOT NULL THEN ' [' || object_name || ']' ELSE '' END ||
       CASE WHEN cardinality IS NOT NULL THEN '  rows=' || cardinality ELSE '' END ||
       CASE WHEN cost IS NOT NULL THEN '  cost=' || cost ELSE '' END AS plan_line
FROM v`$sql_plan
WHERE sql_id = '$sqlId'
ORDER BY id;
"@
                $responseText = Invoke-OraclePlainQuery $planSql -TimeoutSeconds 20
            } else {
                $responseText = 'Invalid SQL ID'
            }

        } else {
            # Default: serve dashboard
            try {
                $data = Get-EtlData
                $html = Build-Html -Data $data
            } catch {
                $html = "<html><body><h2>Error fetching data</h2><pre>$_</pre></body></html>"
            }
            $contentType  = 'text/html; charset=utf-8'
            $responseText = $html
        }

        try {
            $bytes = [System.Text.Encoding]::UTF8.GetBytes($responseText)
            $ctx.Response.ContentType     = $contentType
            $ctx.Response.ContentLength64 = $bytes.LongLength
            $ctx.Response.OutputStream.Write($bytes, 0, $bytes.Length)
            $ctx.Response.OutputStream.Close()
        } catch {
            Write-Host "Response write failed (client disconnected): $_"
        }
    }
} finally {
    $listener.Stop()
    Write-Host "Listener stopped."
}
