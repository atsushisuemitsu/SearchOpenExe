<#
.SYNOPSIS
    SearchOpenExe - Search and manage processes running from a specific folder or locking a file.
.DESCRIPTION
    Right-click context menu tool for Windows 11.
    Searches for:
      - EXE processes running from the target folder
      - Processes that loaded DLLs from the target folder
      - Processes locking files (via Restart Manager API)
    Displays results in a WinForms GUI with the ability to terminate selected processes.
.PARAMETER TargetPath
    The folder or file path to investigate.
#>
param(
    [Parameter(Mandatory = $true)]
    [string]$TargetPath
)

# --- Restart Manager P/Invoke definitions ---
Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;

public struct RM_UNIQUE_PROCESS {
    public int dwProcessId;
    public System.Runtime.InteropServices.ComTypes.FILETIME ProcessStartTime;
}

public enum RM_APP_TYPE {
    RmUnknownApp = 0,
    RmMainWindow = 1,
    RmOtherWindow = 2,
    RmService = 3,
    RmExplorer = 4,
    RmConsole = 5,
    RmCritical = 1000
}

[StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
public struct RM_PROCESS_INFO {
    public RM_UNIQUE_PROCESS Process;
    [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 256)]
    public string strAppName;
    [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 64)]
    public string strServiceShortName;
    public RM_APP_TYPE ApplicationType;
    public int AppStatus;
    public int TSSessionId;
    [MarshalAs(UnmanagedType.Bool)]
    public bool bRestartable;
}

public static class RestartManager {
    [DllImport("rstrtmgr.dll", CharSet = CharSet.Unicode)]
    public static extern int RmStartSession(out uint pSessionHandle, int dwSessionFlags, string strSessionKey);

    [DllImport("rstrtmgr.dll")]
    public static extern int RmEndSession(uint pSessionHandle);

    [DllImport("rstrtmgr.dll", CharSet = CharSet.Unicode)]
    public static extern int RmRegisterResources(uint pSessionHandle,
        uint nFiles, string[] rgsFilenames,
        uint nApplications, [In] RM_UNIQUE_PROCESS[] rgApplications,
        uint nServices, string[] rgsServiceNames);

    [DllImport("rstrtmgr.dll")]
    public static extern int RmGetList(uint pSessionHandle,
        out uint pnProcInfoNeeded, ref uint pnProcInfo,
        [In, Out] RM_PROCESS_INFO[] rgAffectedApps, ref uint lpdwRebootReasons);
}
"@ -ErrorAction SilentlyContinue

# --- Load WinForms ---
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# =============================================================================
# Search Functions
# =============================================================================

function Get-RunningExeProcesses {
    <#
    .SYNOPSIS
        Find processes whose executable is inside the target folder.
    #>
    param([string]$FolderPath)

    $results = @()
    $normalizedFolder = $FolderPath.TrimEnd('\') + '\'

    try {
        $processes = Get-Process | Where-Object { $_.Path }
        foreach ($proc in $processes) {
            try {
                $procPath = $proc.Path
                if ($procPath -and $procPath.StartsWith($normalizedFolder, [StringComparison]::OrdinalIgnoreCase)) {
                    $results += [PSCustomObject]@{
                        ProcessName   = $proc.ProcessName
                        PID           = $proc.Id
                        ExePath       = $procPath
                        DetectionType = "Running EXE"
                        LockedFile    = ""
                    }
                }
            } catch {
                # Access denied - skip silently
            }
        }
    } catch {
        # Ignore enumeration errors
    }

    return $results
}

function Get-LoadedModuleProcesses {
    <#
    .SYNOPSIS
        Find processes that have loaded DLL modules from the target folder.
    #>
    param([string]$FolderPath)

    $results = @()
    $normalizedFolder = $FolderPath.TrimEnd('\') + '\'

    try {
        $processes = Get-Process | Where-Object { $_.Path }
        foreach ($proc in $processes) {
            try {
                $modules = $proc.Modules
                foreach ($mod in $modules) {
                    $modPath = $mod.FileName
                    if ($modPath -and $modPath.StartsWith($normalizedFolder, [StringComparison]::OrdinalIgnoreCase)) {
                        # Exclude the EXE itself (already caught by Get-RunningExeProcesses)
                        if ($modPath -ne $proc.Path) {
                            $results += [PSCustomObject]@{
                                ProcessName   = $proc.ProcessName
                                PID           = $proc.Id
                                ExePath       = $proc.Path
                                DetectionType = "Loaded DLL"
                                LockedFile    = $modPath
                            }
                        }
                    }
                }
            } catch {
                # Access denied or 32/64-bit mismatch - skip silently
            }
        }
    } catch {
        # Ignore enumeration errors
    }

    return $results
}

function Get-FileLockProcesses-Single {
    <#
    .SYNOPSIS
        Use Restart Manager API to find processes locking a single file.
        Returns array of PSCustomObjects with process info.
    #>
    param([string]$FilePath)

    $results = @()
    $sessionHandle = [uint32]0
    $sessionKey = [Guid]::NewGuid().ToString()

    $startResult = [RestartManager]::RmStartSession([ref]$sessionHandle, 0, $sessionKey)
    if ($startResult -ne 0) { return $results }

    try {
        $fileArray = @($FilePath)
        $regResult = [RestartManager]::RmRegisterResources(
            $sessionHandle,
            [uint32]1, $fileArray,
            0, $null,
            0, $null
        )
        if ($regResult -ne 0) { return $results }

        $pnProcInfoNeeded = [uint32]0
        $pnProcInfo = [uint32]0
        $rebootReasons = [uint32]0

        # First call to get count
        [RestartManager]::RmGetList($sessionHandle,
            [ref]$pnProcInfoNeeded, [ref]$pnProcInfo, $null, [ref]$rebootReasons) | Out-Null

        if ($pnProcInfoNeeded -gt 0) {
            $pnProcInfo = $pnProcInfoNeeded
            $processInfoArray = New-Object RM_PROCESS_INFO[] $pnProcInfo

            $getResult = [RestartManager]::RmGetList($sessionHandle,
                [ref]$pnProcInfoNeeded, [ref]$pnProcInfo, $processInfoArray, [ref]$rebootReasons)

            if ($getResult -eq 0) {
                foreach ($pi in $processInfoArray) {
                    if ($pi.Process.dwProcessId -eq 0) { continue }
                    try {
                        $proc = Get-Process -Id $pi.Process.dwProcessId -ErrorAction SilentlyContinue
                        $exePath = if ($proc -and $proc.Path) { $proc.Path } else { "" }

                        $results += [PSCustomObject]@{
                            ProcessName   = $pi.strAppName
                            PID           = $pi.Process.dwProcessId
                            ExePath       = $exePath
                            DetectionType = "File Lock (RM)"
                            LockedFile    = $FilePath
                        }
                    } catch {
                        # Skip processes we can't access
                    }
                }
            }
        }
    } finally {
        [RestartManager]::RmEndSession($sessionHandle) | Out-Null
    }

    return $results
}

function Get-FileLockProcesses {
    <#
    .SYNOPSIS
        Use Restart Manager API to find processes locking files.
        Checks each file individually to identify exactly which file is locked.
    #>
    param(
        [string]$Path,
        [bool]$IsFile = $false
    )

    $results = @()

    if ($IsFile) {
        $results += Get-FileLockProcesses-Single -FilePath $Path
    } else {
        # Recursively enumerate files with a cap of 1000
        $files = @()
        try {
            $allFiles = Get-ChildItem -Path $Path -Recurse -File -ErrorAction SilentlyContinue | Select-Object -First 1000
            $files = @($allFiles | ForEach-Object { $_.FullName })
        } catch {
            return $results
        }

        # Check each file individually so we know exactly which file is locked
        foreach ($file in $files) {
            $fileResults = Get-FileLockProcesses-Single -FilePath $file
            if ($fileResults) {
                $results += $fileResults
            }
        }
    }

    return $results
}

function Get-AllProcessInfo {
    <#
    .SYNOPSIS
        Aggregate results from all detection methods and deduplicate by PID.
    #>
    param(
        [string]$Path,
        [bool]$IsFile = $false
    )

    $allResults = @()

    if ($IsFile) {
        # For a file, use RM file lock detection
        $rmResults = Get-FileLockProcesses -Path $Path -IsFile $true
        $allResults += $rmResults

        # Also check if the file itself is an EXE being run
        if ($Path -match '\.exe$') {
            try {
                $procs = Get-Process | Where-Object { $_.Path -eq $Path }
                foreach ($proc in $procs) {
                    $allResults += [PSCustomObject]@{
                        ProcessName   = $proc.ProcessName
                        PID           = $proc.Id
                        ExePath       = $proc.Path
                        DetectionType = "Running EXE"
                        LockedFile    = ""
                    }
                }
            } catch {}
        }
    } else {
        # For a folder, use all three methods
        $exeResults = Get-RunningExeProcesses -FolderPath $Path
        $dllResults = Get-LoadedModuleProcesses -FolderPath $Path
        $rmResults  = Get-FileLockProcesses -Path $Path -IsFile $false

        $allResults += $exeResults
        $allResults += $dllResults
        $allResults += $rmResults
    }

    # Deduplicate by PID + DetectionType (keep unique detection entries)
    $seen = @{}
    $unique = @()
    foreach ($item in $allResults) {
        $key = "$($item.PID)|$($item.DetectionType)"
        if (-not $seen.ContainsKey($key)) {
            $seen[$key] = $true
            $unique += $item
        }
    }

    return $unique
}

# =============================================================================
# GUI
# =============================================================================

function Show-ProcessGui {
    param(
        [string]$Path,
        [bool]$IsFile = $false
    )

    # --- Choose font ---
    $fontFamily = "Yu Gothic UI"
    $testFont = New-Object System.Drawing.FontFamily($fontFamily) -ErrorAction SilentlyContinue
    if (-not $testFont) { $fontFamily = "Segoe UI" }

    $font      = New-Object System.Drawing.Font($fontFamily, 9)
    $fontBold  = New-Object System.Drawing.Font($fontFamily, 9, [System.Drawing.FontStyle]::Bold)
    $fontTitle = New-Object System.Drawing.Font($fontFamily, 11, [System.Drawing.FontStyle]::Bold)

    # --- Form ---
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "SearchOpenExe - Process Finder"
    $form.Size = New-Object System.Drawing.Size(950, 560)
    $form.StartPosition = "CenterScreen"
    $form.Font = $font
    $form.MinimumSize = New-Object System.Drawing.Size(750, 400)

    # --- Title label ---
    $lblTitle = New-Object System.Windows.Forms.Label
    $lblTitle.Font = $fontTitle
    $lblTitle.AutoSize = $false
    $lblTitle.Size = New-Object System.Drawing.Size(910, 22)
    $lblTitle.Location = New-Object System.Drawing.Point(10, 8)
    $lblTitle.Anchor = "Top, Left, Right"
    $displayPath = if ($Path.Length -gt 80) { $Path.Substring(0, 77) + "..." } else { $Path }
    $typeLabel = if ($IsFile) { "File" } else { "Folder" }
    $lblTitle.Text = "Target ($typeLabel): $displayPath"
    $form.Controls.Add($lblTitle)

    # --- Status label ---
    $lblStatus = New-Object System.Windows.Forms.Label
    $lblStatus.AutoSize = $false
    $lblStatus.Size = New-Object System.Drawing.Size(910, 20)
    $lblStatus.Location = New-Object System.Drawing.Point(10, 32)
    $lblStatus.Anchor = "Top, Left, Right"
    $lblStatus.ForeColor = [System.Drawing.Color]::Gray
    $lblStatus.Text = "Searching..."
    $form.Controls.Add($lblStatus)

    # --- DataGridView ---
    $script:dgv = New-Object System.Windows.Forms.DataGridView
    $script:dgv.Location = New-Object System.Drawing.Point(10, 56)
    $script:dgv.Size = New-Object System.Drawing.Size(910, 410)
    $script:dgv.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
    $script:dgv.AllowUserToAddRows = $false
    $script:dgv.AllowUserToDeleteRows = $false
    $script:dgv.ReadOnly = $true
    $script:dgv.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
    $script:dgv.MultiSelect = $true
    $script:dgv.AutoGenerateColumns = $false
    $script:dgv.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
    $script:dgv.RowHeadersVisible = $false
    $script:dgv.BackgroundColor = [System.Drawing.Color]::White
    $script:dgv.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $script:dgv.Font = $font
    $script:dgv.ColumnHeadersDefaultCellStyle.Font = $fontBold

    # Columns with DataPropertyName for DataTable binding
    $colName = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colName.HeaderText = "Process Name"
    $colName.Name = "ProcessName"
    $colName.DataPropertyName = "ProcessName"
    $colName.FillWeight = 15

    $colPID = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colPID.HeaderText = "PID"
    $colPID.Name = "PID"
    $colPID.DataPropertyName = "PID"
    $colPID.FillWeight = 8

    $colPath = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colPath.HeaderText = "Exe Path"
    $colPath.Name = "ExePath"
    $colPath.DataPropertyName = "ExePath"
    $colPath.FillWeight = 30

    $colType = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colType.HeaderText = "Detection Type"
    $colType.Name = "DetectionType"
    $colType.DataPropertyName = "DetectionType"
    $colType.FillWeight = 12

    $colLocked = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colLocked.HeaderText = "Locked File / Module"
    $colLocked.Name = "LockedFile"
    $colLocked.DataPropertyName = "LockedFile"
    $colLocked.FillWeight = 35

    $script:dgv.Columns.Add($colName) | Out-Null
    $script:dgv.Columns.Add($colPID) | Out-Null
    $script:dgv.Columns.Add($colPath) | Out-Null
    $script:dgv.Columns.Add($colType) | Out-Null
    $script:dgv.Columns.Add($colLocked) | Out-Null
    $form.Controls.Add($script:dgv)

    # --- Button panel ---
    $panelButtons = New-Object System.Windows.Forms.FlowLayoutPanel
    $panelButtons.FlowDirection = "RightToLeft"
    $panelButtons.Size = New-Object System.Drawing.Size(910, 40)
    $panelButtons.Location = New-Object System.Drawing.Point(10, 472)
    $panelButtons.Anchor = "Bottom, Left, Right"

    $btnClose = New-Object System.Windows.Forms.Button
    $btnClose.Text = "Close"
    $btnClose.Size = New-Object System.Drawing.Size(90, 30)
    $btnClose.Font = $font

    $btnKill = New-Object System.Windows.Forms.Button
    $btnKill.Text = "Terminate Selected"
    $btnKill.Size = New-Object System.Drawing.Size(140, 30)
    $btnKill.Font = $fontBold
    $btnKill.BackColor = [System.Drawing.Color]::FromArgb(220, 53, 69)
    $btnKill.ForeColor = [System.Drawing.Color]::White
    $btnKill.FlatStyle = "Flat"

    $btnRefresh = New-Object System.Windows.Forms.Button
    $btnRefresh.Text = "Refresh"
    $btnRefresh.Size = New-Object System.Drawing.Size(90, 30)
    $btnRefresh.Font = $font

    $panelButtons.Controls.Add($btnClose)
    $panelButtons.Controls.Add($btnKill)
    $panelButtons.Controls.Add($btnRefresh)
    $form.Controls.Add($panelButtons)

    # --- Refresh function ---
    $script:refreshData = {
        $script:dgv.DataSource = $null
        $lblStatus.Text = "Searching..."
        $lblStatus.ForeColor = [System.Drawing.Color]::Gray
        $form.Refresh()

        $data = Get-AllProcessInfo -Path $Path -IsFile $IsFile

        # Build DataTable
        $dt = New-Object System.Data.DataTable
        [void]$dt.Columns.Add("ProcessName", [string])
        [void]$dt.Columns.Add("PID", [string])
        [void]$dt.Columns.Add("ExePath", [string])
        [void]$dt.Columns.Add("DetectionType", [string])
        [void]$dt.Columns.Add("LockedFile", [string])

        foreach ($item in $data) {
            [void]$dt.Rows.Add(
                [string]$item.ProcessName,
                [string]$item.PID,
                [string]$item.ExePath,
                [string]$item.DetectionType,
                [string]$item.LockedFile
            )
        }

        $script:dgv.DataSource = $dt

        if ($dt.Rows.Count -eq 0) {
            $lblStatus.Text = "No processes found."
            $lblStatus.ForeColor = [System.Drawing.Color]::Green
        } else {
            $lblStatus.Text = "$($dt.Rows.Count) process(es) found."
            $lblStatus.ForeColor = [System.Drawing.Color]::DarkBlue
        }
    }

    # --- Event handlers ---
    $btnClose.Add_Click({ $form.Close() })

    $btnRefresh.Add_Click({ & $script:refreshData })

    $btnKill.Add_Click({
        $selectedRows = $script:dgv.SelectedRows
        if ($selectedRows.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show(
                "Please select one or more processes to terminate.",
                "No Selection",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            )
            return
        }

        # Use $targetPid / $procName to avoid conflict with automatic variable $PID
        $targetPids = @()
        $nameList = @()
        foreach ($row in $selectedRows) {
            $cellPid = $row.Cells["PID"].Value
            $procName = $row.Cells["ProcessName"].Value
            if ($cellPid -and $cellPid -ne "") {
                $targetPids += [int]$cellPid
                $nameList += "$procName (PID: $cellPid)"
            }
        }

        if ($targetPids.Count -eq 0) { return }

        $msg = "Are you sure you want to terminate the following process(es)?`n`n"
        $msg += ($nameList -join "`n")
        $msg += "`n`nThis action cannot be undone."

        $confirm = [System.Windows.Forms.MessageBox]::Show(
            $msg,
            "Confirm Termination",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )

        if ($confirm -eq [System.Windows.Forms.DialogResult]::Yes) {
            $successCount = 0
            $failCount = 0
            $failDetails = @()
            foreach ($targetPid in $targetPids) {
                try {
                    Stop-Process -Id $targetPid -Force -ErrorAction Stop
                    # Verify the process is actually gone
                    Start-Sleep -Milliseconds 500
                    $stillAlive = Get-Process -Id $targetPid -ErrorAction SilentlyContinue
                    if ($stillAlive) {
                        $failCount++
                        $failDetails += "PID $targetPid : process still running after Stop-Process"
                    } else {
                        $successCount++
                    }
                } catch {
                    $failCount++
                    $failDetails += "PID ${targetPid}: $($_.Exception.Message)"
                }
            }

            $resultMsg = "Terminated: $successCount"
            if ($failCount -gt 0) {
                $resultMsg += "`nFailed: $failCount"
                $resultMsg += "`n`nDetails:`n"
                $resultMsg += ($failDetails -join "`n")
                $resultMsg += "`n`nTry running as Administrator for full access."
            }
            [System.Windows.Forms.MessageBox]::Show(
                $resultMsg,
                "Result",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                $(if ($failCount -gt 0) { [System.Windows.Forms.MessageBoxIcon]::Warning } else { [System.Windows.Forms.MessageBoxIcon]::Information })
            )

            # Auto-refresh after termination
            & $script:refreshData
        }
    })

    # --- Initial load ---
    $form.Add_Shown({ & $script:refreshData })

    # --- Show ---
    [System.Windows.Forms.Application]::EnableVisualStyles()
    $form.ShowDialog() | Out-Null

    # --- Cleanup ---
    $form.Dispose()
}

# =============================================================================
# Main Entry Point
# =============================================================================

# Validate path
if (-not (Test-Path $TargetPath)) {
    [System.Windows.Forms.MessageBox]::Show(
        "The specified path does not exist:`n$TargetPath",
        "SearchOpenExe - Error",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    )
    exit 1
}

$resolvedPath = (Resolve-Path $TargetPath).Path
$isFile = -not (Test-Path $resolvedPath -PathType Container)

Show-ProcessGui -Path $resolvedPath -IsFile $isFile
