# =====================================================================
#  Deploy-FolderToServers.ps1  -  v1.0.0
#  ------------------------------------------------------------------
#  Copies a local folder to multiple remote servers (fast, robocopy /MT)
#  Before each copy: backs up the remote target folder to d$\backup
#  on the SAME remote server, named <FolderName>_yyyyMMdd_HHmm
#  All settings are stored in a JSON file next to the script.
#  Servers are selected with checkboxes and run IN PARALLEL (runspaces).
#
#  נכתב ע"י רועי חדד
# =====================================================================

#Requires -Version 5.1

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# ---------------------------------------------------------------------
#  Config
# ---------------------------------------------------------------------
$ScriptDir  = Split-Path -Parent $MyInvocation.MyCommand.Path
$ConfigPath = Join-Path $ScriptDir 'DeployTool-Config.json'

if (-not (Test-Path $ConfigPath)) {
    $defaultConfig = [ordered]@{
        SourceFolder       = 'C:\Deploy\wwwroot'
        RemoteTargetFolder = 'd$\Webroot\wwwroot'
        BackupFolder       = 'd$\backup'
        MaxParallel        = 5
        RobocopyThreads    = 32
        MirrorMode         = $false     # $true = /MIR (deletes extra files on target!), $false = /E
        StopIIS            = $false     # default state of the 'Stop IIS' checkbox
        UseWinRM           = $true      # backup runs locally on the remote server (much faster)
        Servers            = @(
            [ordered]@{ Name = 'NTAS102583A9F'; IP = '172.29.15.63'; Description = 'Kohav_Business_Prod_Web01'; Checked = $false }
        )
    }
    $defaultConfig | ConvertTo-Json -Depth 5 | Set-Content -Path $ConfigPath -Encoding UTF8
}

try {
    $Config = Get-Content $ConfigPath -Raw -Encoding UTF8 | ConvertFrom-Json
} catch {
    [System.Windows.Forms.MessageBox]::Show("Failed to read config file:`n$($_.Exception.Message)", 'Deploy Tool',
        [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
    exit 1
}

# ---------------------------------------------------------------------
#  Shared state (GUI <-> runspaces)
# ---------------------------------------------------------------------
$Sync = [hashtable]::Synchronized(@{
    LogQueue = [System.Collections.Queue]::Synchronized((New-Object System.Collections.Queue))
})
$script:Jobs         = @()
$script:RunspacePool = $null

# ---------------------------------------------------------------------
#  Colors / fonts  (styled like the Folder Comparison Tool)
# ---------------------------------------------------------------------
$ClrNavy   = [System.Drawing.Color]::FromArgb(31, 56, 100)
$ClrBlue   = [System.Drawing.Color]::FromArgb(41, 84, 173)
$ClrGreen  = [System.Drawing.Color]::FromArgb(46, 125, 50)
$ClrRed    = [System.Drawing.Color]::FromArgb(192, 57, 43)
$ClrGray   = [System.Drawing.Color]::FromArgb(120, 128, 136)
$ClrWhite  = [System.Drawing.Color]::White
$FontBase  = New-Object System.Drawing.Font('Segoe UI', 9)
$FontBold  = New-Object System.Drawing.Font('Segoe UI', 9,  [System.Drawing.FontStyle]::Bold)
$FontTitle = New-Object System.Drawing.Font('Segoe UI', 16, [System.Drawing.FontStyle]::Bold)
$FontMono  = New-Object System.Drawing.Font('Consolas', 9)

function New-ColorButton {
    param($Text, $X, $Y, $W, $H, $Color)
    $b = New-Object System.Windows.Forms.Button
    $b.Text      = $Text
    $b.Location  = New-Object System.Drawing.Point($X, $Y)
    $b.Size      = New-Object System.Drawing.Size($W, $H)
    $b.BackColor = $Color
    $b.ForeColor = $ClrWhite
    $b.FlatStyle = 'Flat'
    $b.FlatAppearance.BorderSize = 0
    $b.Font      = $FontBold
    $b.Cursor    = 'Hand'
    return $b
}

# ---------------------------------------------------------------------
#  Main form
# ---------------------------------------------------------------------
$Form                 = New-Object System.Windows.Forms.Form
$Form.Text            = 'Folder Deploy Tool'
$Form.Size            = New-Object System.Drawing.Size(900, 820)
$Form.StartPosition   = 'CenterScreen'
$Form.BackColor       = $ClrWhite
$Form.FormBorderStyle = 'FixedSingle'
$Form.MaximizeBox     = $false
$Form.Font            = $FontBase

# --- Header -----------------------------------------------------------
$lblTitle           = New-Object System.Windows.Forms.Label
$lblTitle.Text      = 'Folder Deploy Tool'
$lblTitle.Font      = $FontTitle
$lblTitle.ForeColor = $ClrNavy
$lblTitle.Location  = New-Object System.Drawing.Point(30, 20)
$lblTitle.AutoSize  = $true
$Form.Controls.Add($lblTitle)

$lblVer           = New-Object System.Windows.Forms.Label
$lblVer.Text      = 'v1.0.0'
$lblVer.ForeColor = $ClrGray
$lblVer.Location  = New-Object System.Drawing.Point(420, 30)
$lblVer.AutoSize  = $true
$Form.Controls.Add($lblVer)

$lblAuthor           = New-Object System.Windows.Forms.Label
$lblAuthor.Text      = 'נכתב ע"י רועי חדד'
$lblAuthor.ForeColor = $ClrNavy
$lblAuthor.Location  = New-Object System.Drawing.Point(700, 30)
$lblAuthor.AutoSize  = $true
$Form.Controls.Add($lblAuthor)

$sep           = New-Object System.Windows.Forms.Label
$sep.BorderStyle = 'Fixed3D'
$sep.Location  = New-Object System.Drawing.Point(30, 60)
$sep.Size      = New-Object System.Drawing.Size(820, 2)
$Form.Controls.Add($sep)

# --- Source folder ----------------------------------------------------
$lblSrc          = New-Object System.Windows.Forms.Label
$lblSrc.Text     = 'Source Folder:'
$lblSrc.Font     = $FontBold
$lblSrc.ForeColor= $ClrNavy
$lblSrc.Location = New-Object System.Drawing.Point(30, 85)
$lblSrc.AutoSize = $true
$Form.Controls.Add($lblSrc)

$txtSrc          = New-Object System.Windows.Forms.TextBox
$txtSrc.Location = New-Object System.Drawing.Point(180, 82)
$txtSrc.Size     = New-Object System.Drawing.Size(560, 24)
$txtSrc.Text     = $Config.SourceFolder
$Form.Controls.Add($txtSrc)

$btnBrowse = New-ColorButton '...' 750 80 60 26 $ClrBlue
$btnBrowse.Add_Click({
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    if ($txtSrc.Text -and (Test-Path $txtSrc.Text)) { $dlg.SelectedPath = $txtSrc.Text }
    if ($dlg.ShowDialog() -eq 'OK') { $txtSrc.Text = $dlg.SelectedPath }
})
$Form.Controls.Add($btnBrowse)

# --- Target servers ---------------------------------------------------
$lblSrv          = New-Object System.Windows.Forms.Label
$lblSrv.Text     = 'Target Servers:'
$lblSrv.Font     = $FontBold
$lblSrv.ForeColor= $ClrNavy
$lblSrv.Location = New-Object System.Drawing.Point(30, 125)
$lblSrv.AutoSize = $true
$Form.Controls.Add($lblSrv)

$lvServers               = New-Object System.Windows.Forms.ListView
$lvServers.Location      = New-Object System.Drawing.Point(180, 125)
$lvServers.Size          = New-Object System.Drawing.Size(560, 170)
$lvServers.View          = 'Details'
$lvServers.CheckBoxes    = $true
$lvServers.FullRowSelect = $true
$lvServers.GridLines     = $true
[void]$lvServers.Columns.Add('שם שרת',   130)
[void]$lvServers.Columns.Add('כתובת',    110)
[void]$lvServers.Columns.Add('תיאור',    200)
[void]$lvServers.Columns.Add('סטטוס',    100)
$Form.Controls.Add($lvServers)

function Add-ServerRow {
    param($Name, $IP, $Description, [bool]$Checked)
    $it = New-Object System.Windows.Forms.ListViewItem($Name)
    [void]$it.SubItems.Add($IP)
    [void]$it.SubItems.Add($Description)
    [void]$it.SubItems.Add('')
    $it.Checked = $Checked
    [void]$lvServers.Items.Add($it)
}
foreach ($s in $Config.Servers) { Add-ServerRow $s.Name $s.IP $s.Description ([bool]$s.Checked) }

$btnAdd    = New-ColorButton 'Add'    755 125 95 28 $ClrGreen
$btnRemove = New-ColorButton 'Remove' 755 160 95 28 $ClrRed
$btnAll    = New-ColorButton 'All'    755 195 95 28 $ClrBlue
$btnNone   = New-ColorButton 'None'   755 230 95 28 $ClrGray
$Form.Controls.AddRange(@($btnAdd, $btnRemove, $btnAll, $btnNone))

$btnAll.Add_Click({  foreach ($i in $lvServers.Items) { $i.Checked = $true  } })
$btnNone.Add_Click({ foreach ($i in $lvServers.Items) { $i.Checked = $false } })
$btnRemove.Add_Click({
    foreach ($i in @($lvServers.SelectedItems)) { $lvServers.Items.Remove($i) }
})
$btnAdd.Add_Click({
    $dlgF               = New-Object System.Windows.Forms.Form
    $dlgF.Text          = 'Add Server'
    $dlgF.Size          = New-Object System.Drawing.Size(360, 220)
    $dlgF.StartPosition = 'CenterParent'
    $dlgF.FormBorderStyle = 'FixedDialog'
    $dlgF.MaximizeBox   = $false; $dlgF.MinimizeBox = $false
    $labels = @('שם שרת:', 'כתובת IP:', 'תיאור:')
    $boxes  = @()
    for ($i = 0; $i -lt 3; $i++) {
        $l = New-Object System.Windows.Forms.Label
        $l.Text = $labels[$i]; $l.Location = New-Object System.Drawing.Point(15, (20 + $i * 35)); $l.AutoSize = $true
        $t = New-Object System.Windows.Forms.TextBox
        $t.Location = New-Object System.Drawing.Point(110, (17 + $i * 35)); $t.Size = New-Object System.Drawing.Size(215, 24)
        $dlgF.Controls.Add($l); $dlgF.Controls.Add($t); $boxes += $t
    }
    $ok = New-ColorButton 'OK' 110 135 100 28 $ClrGreen
    $ok.DialogResult = 'OK'
    $dlgF.Controls.Add($ok)
    $dlgF.AcceptButton = $ok
    if ($dlgF.ShowDialog($Form) -eq 'OK' -and $boxes[0].Text -and $boxes[1].Text) {
        Add-ServerRow $boxes[0].Text $boxes[1].Text $boxes[2].Text $true
    }
})

# --- Remote target + backup folder -----------------------------------
$lblTgt          = New-Object System.Windows.Forms.Label
$lblTgt.Text     = 'Remote Target:'
$lblTgt.Font     = $FontBold
$lblTgt.ForeColor= $ClrNavy
$lblTgt.Location = New-Object System.Drawing.Point(30, 315)
$lblTgt.AutoSize = $true
$Form.Controls.Add($lblTgt)

$txtTgt          = New-Object System.Windows.Forms.TextBox
$txtTgt.Location = New-Object System.Drawing.Point(180, 312)
$txtTgt.Size     = New-Object System.Drawing.Size(560, 24)
$txtTgt.Text     = $Config.RemoteTargetFolder
$Form.Controls.Add($txtTgt)

$lblBak          = New-Object System.Windows.Forms.Label
$lblBak.Text     = 'Backup Folder:'
$lblBak.Font     = $FontBold
$lblBak.ForeColor= $ClrNavy
$lblBak.Location = New-Object System.Drawing.Point(30, 350)
$lblBak.AutoSize = $true
$Form.Controls.Add($lblBak)

$txtBak          = New-Object System.Windows.Forms.TextBox
$txtBak.Location = New-Object System.Drawing.Point(180, 347)
$txtBak.Size     = New-Object System.Drawing.Size(560, 24)
$txtBak.Text     = $Config.BackupFolder
$Form.Controls.Add($txtBak)

# --- Stop IIS option ---------------------------------------------------
$chkIIS           = New-Object System.Windows.Forms.CheckBox
$chkIIS.Text      = 'Stop IIS before copy (start again when finished)'
$chkIIS.Font      = $FontBold
$chkIIS.ForeColor = $ClrNavy
$chkIIS.Location  = New-Object System.Drawing.Point(180, 380)
$chkIIS.AutoSize  = $true
$chkIIS.Checked   = [bool]$Config.StopIIS
$Form.Controls.Add($chkIIS)

# --- Log --------------------------------------------------------------
$lblLog          = New-Object System.Windows.Forms.Label
$lblLog.Text     = 'Log:'
$lblLog.Font     = $FontBold
$lblLog.ForeColor= $ClrNavy
$lblLog.Location = New-Object System.Drawing.Point(30, 408)
$lblLog.AutoSize = $true
$Form.Controls.Add($lblLog)

$rtbLog             = New-Object System.Windows.Forms.RichTextBox
$rtbLog.Location    = New-Object System.Drawing.Point(30, 430)
$rtbLog.Size        = New-Object System.Drawing.Size(820, 265)
$rtbLog.BackColor   = [System.Drawing.Color]::Black
$rtbLog.ForeColor   = [System.Drawing.Color]::Lime
$rtbLog.Font        = $FontMono
$rtbLog.ReadOnly    = $true
$rtbLog.BorderStyle = 'FixedSingle'
$Form.Controls.Add($rtbLog)

function Write-Log {
    param($Text, $Color = [System.Drawing.Color]::Lime)
    $rtbLog.SelectionStart  = $rtbLog.TextLength
    $rtbLog.SelectionColor  = $Color
    $line = "[{0}] {1}`r`n" -f (Get-Date -Format 'HH:mm:ss'), $Text
    $rtbLog.AppendText($line)
    $rtbLog.ScrollToCaret()
}

# --- Bottom buttons ----------------------------------------------------
$ClrOrange = [System.Drawing.Color]::FromArgb(211, 120, 0)
$btnRun     = New-ColorButton 'Run Deploy'    30  715 130 36 $ClrBlue
$btnStopSrv = New-ColorButton 'Stop Server'   170 715 115 36 $ClrOrange
$btnStopAll = New-ColorButton 'Stop All'      295 715 100 36 $ClrRed
$btnClear   = New-ColorButton 'Clear Log'     420 715 105 36 $ClrGray
$btnSave    = New-ColorButton 'Save Settings' 535 715 125 36 $ClrGreen
$btnClose   = New-ColorButton 'Close'         745 715 105 36 $ClrRed
$btnStopSrv.Enabled = $false
$btnStopAll.Enabled = $false
$Form.Controls.AddRange(@($btnRun, $btnStopSrv, $btnStopAll, $btnClear, $btnSave, $btnClose))

$btnClear.Add_Click({ $rtbLog.Clear() })
$btnClose.Add_Click({ $Form.Close() })

# --- Save settings -----------------------------------------------------
function Save-Settings {
    $servers = @()
    foreach ($i in $lvServers.Items) {
        $servers += [ordered]@{
            Name        = $i.SubItems[0].Text
            IP          = $i.SubItems[1].Text
            Description = $i.SubItems[2].Text
            Checked     = [bool]$i.Checked
        }
    }
    $out = [ordered]@{
        SourceFolder       = $txtSrc.Text
        RemoteTargetFolder = $txtTgt.Text
        BackupFolder       = $txtBak.Text
        MaxParallel        = [int]$Config.MaxParallel
        RobocopyThreads    = [int]$Config.RobocopyThreads
        MirrorMode         = [bool]$Config.MirrorMode
        StopIIS            = [bool]$chkIIS.Checked
        UseWinRM           = [bool]$Config.UseWinRM
        Servers            = $servers
    }
    $out | ConvertTo-Json -Depth 5 | Set-Content -Path $ConfigPath -Encoding UTF8
    Write-Log "Settings saved to $ConfigPath" ([System.Drawing.Color]::Cyan)
}
$btnSave.Add_Click({ Save-Settings })

# ---------------------------------------------------------------------
#  Worker scriptblock (runs inside a runspace, one per server)
# ---------------------------------------------------------------------
$Worker = {
    param($Server, $Cfg, $Sync)

    function Q { param($Kind, $Status, $Msg)
        $Sync.LogQueue.Enqueue([pscustomobject]@{
            Server = $Server.Name; Kind = $Kind; Status = $Status; Message = $Msg
        })
    }

    try {
        $ip        = $Server.IP
        $remoteRel = $Cfg.RemoteTargetFolder.Trim('\')          # e.g. d$\Webroot\wwwroot
        $backupRel = $Cfg.BackupFolder.Trim('\')                # e.g. d$\backup
        $uncTarget = "\\$ip\$remoteRel"
        $uncBackup = "\\$ip\$backupRel"
        $localTarget = ($remoteRel -replace '^([a-zA-Z])\$', '${1}:')   # D:\Webroot\wwwroot
        $localBackup = ($backupRel -replace '^([a-zA-Z])\$', '${1}:')   # D:\backup
        $leaf       = Split-Path $remoteRel -Leaf
        $stamp      = Get-Date -Format 'yyyyMMdd_HHmm'
        $backupName = "${leaf}_${stamp}"
        $mt         = [int]$Cfg.RobocopyThreads

        # ---------- helper: stop/start IIS on the remote server ----------
        function Set-RemoteIIS {
            param($Action)   # 'stop' or 'start'
            # Preferred: run iisreset locally on the server via WinRM (hostname -> Kerberos)
            try {
                $rc = Invoke-Command -ComputerName $Server.Name -ErrorAction Stop -ScriptBlock {
                    param($a)
                    & iisreset "/$a" 2>&1 | Out-Null
                    return $LASTEXITCODE
                } -ArgumentList $Action
                if ($rc -eq 0) { return 'WinRM' }
            } catch { }
            # Fallback: iisreset supports a remote computer name directly (RPC)
            & iisreset $Server.Name "/$Action" 2>&1 | Out-Null
            if ($LASTEXITCODE -eq 0) { return 'RPC' }
            return $null
        }

        # ---------- 1) Stop IIS (optional) - BEFORE backup, so the backup is consistent ----------
        $iisStopped = $false
        if ($Cfg.StopIIS) {
            Q 'IIS' 'Stopping IIS...' 'Stopping IIS...'
            $via = Set-RemoteIIS 'stop'
            if (-not $via) { throw 'Failed to stop IIS - aborting (no backup, target was NOT modified)' }
            $iisStopped = $true
            Q 'IIS' 'IIS is DOWN' "IIS stopped (via $via) - site is now offline"
        }

        try {
            # ---------- 2) Backup on the remote server ----------
            if (Test-Path $uncTarget) {
                Q 'Backup' 'Backing up...' "Starting backup -> $backupRel\$backupName"
                $backupOk = $false

                if ($Cfg.UseWinRM) {
                    try {
                        # Kerberos authentication requires a hostname, not an IP address
                        $rc = Invoke-Command -ComputerName $Server.Name -ErrorAction Stop -ScriptBlock {
                            param($src, $dst, $threads)
                            $parent = Split-Path $dst -Parent
                            if (-not (Test-Path $parent)) { New-Item -ItemType Directory -Path $parent -Force | Out-Null }
                            robocopy $src $dst /E /MT:$threads /R:1 /W:1 /NP /NFL /NDL /NJH /NJS | Out-Null
                            return $LASTEXITCODE
                        } -ArgumentList $localTarget, (Join-Path $localBackup $backupName), $mt
                        if ($rc -lt 8) {
                            $backupOk = $true
                            Q 'Backup' 'Backup done' "Backup completed locally on $($Server.Name) (WinRM disk-to-disk, exit $rc)"
                        } else {
                            Q 'Backup' 'Backing up (UNC)...' "Remote robocopy failed (exit $rc), falling back to UNC backup"
                        }
                    } catch {
                        Q 'Backup' 'Backing up (UNC)...' "WinRM unavailable ($($_.Exception.Message.Trim())), falling back to UNC backup"
                    }
                }

                if (-not $backupOk) {
                    if (-not (Test-Path $uncBackup)) { New-Item -ItemType Directory -Path $uncBackup -Force | Out-Null }
                    robocopy $uncTarget (Join-Path $uncBackup $backupName) /E /MT:$mt /R:1 /W:1 /NP /NFL /NDL /NJH /NJS | Out-Null
                    if ($LASTEXITCODE -ge 8) { throw "Backup failed (robocopy exit $LASTEXITCODE)" }
                    Q 'Backup' 'Backup done' "Backup completed via UNC (exit $LASTEXITCODE)"
                }
            } else {
                Q 'Backup' 'No backup' 'Remote target folder does not exist - skipping backup'
            }

            # ---------- 3) Stop IIS again right before the copy ----------
            # (safety net: monitoring / another admin may have started IIS back
            #  while the backup was running - make sure it is really down)
            if ($Cfg.StopIIS) {
                Q 'IIS' 'Stopping IIS...' 'Re-verifying IIS is stopped before copy...'
                $via = Set-RemoteIIS 'stop'
                if (-not $via) { throw 'Failed to stop IIS before copy - aborting (backup exists, target was NOT modified)' }
                $iisStopped = $true
                Q 'IIS' 'IIS is DOWN' "IIS confirmed stopped (via $via) - starting copy"
            }

            # ---------- 4) Fast copy local -> remote ----------
            Q 'Copy' 'Copying...' "Copying $($Cfg.SourceFolder) -> $uncTarget (MT:$mt)"
            $sw = [System.Diagnostics.Stopwatch]::StartNew()
            $mode = if ($Cfg.MirrorMode) { '/MIR' } else { '/E' }
            robocopy $Cfg.SourceFolder $uncTarget $mode /MT:$mt /R:2 /W:2 /NP /NFL /NDL /NJH /NJS | Out-Null
            $rcCopy = $LASTEXITCODE
            $sw.Stop()
            if ($rcCopy -ge 8) { throw "Copy failed (robocopy exit $rcCopy)" }
        }
        finally {
            # ---------- 5) Start IIS again - even if the backup or copy failed ----------
            if ($iisStopped) {
                Q 'IIS' 'Starting IIS...' 'Starting IIS...'
                $via = Set-RemoteIIS 'start'
                if ($via) { Q 'IIS' 'IIS is UP' "IIS started (via $via) - site is back online" }
                else      { Q 'Error' 'IIS STILL DOWN!' 'FAILED TO START IIS - start it manually on the server!' }
            }
        }

        Q 'Done' 'Done' ("Finished successfully in {0:mm\:ss} (robocopy exit {1})" -f $sw.Elapsed, $rcCopy)
    }
    catch {
        Q 'Error' 'Error' $_.Exception.Message
    }
}

# ---------------------------------------------------------------------
#  UI timer - drains the log queue + tracks job completion
# ---------------------------------------------------------------------
$Timer          = New-Object System.Windows.Forms.Timer
$Timer.Interval = 300
$Timer.Add_Tick({
    while ($Sync.LogQueue.Count -gt 0) {
        $e = $Sync.LogQueue.Dequeue()
        $color = switch ($e.Kind) {
            'Error'  { [System.Drawing.Color]::Red }
            'Done'   { [System.Drawing.Color]::LimeGreen }
            'Backup' { [System.Drawing.Color]::Yellow }
            'Copy'   { [System.Drawing.Color]::Cyan }
            'IIS'    { [System.Drawing.Color]::Orange }
            default  { [System.Drawing.Color]::Lime }
        }
        Write-Log "[$($e.Server)] $($e.Message)" $color

        foreach ($i in $lvServers.Items) {
            if ($i.SubItems[0].Text -eq $e.Server) { $i.SubItems[3].Text = $e.Status }
        }
    }

    if ($script:Jobs.Count -gt 0 -and -not ($script:Jobs | Where-Object { -not $_.Handle.IsCompleted })) {
        foreach ($j in $script:Jobs) {
            try { $j.PS.EndInvoke($j.Handle) } catch {}
            $j.PS.Dispose()
        }
        $script:Jobs = @()
        if ($script:RunspacePool) { $script:RunspacePool.Close(); $script:RunspacePool.Dispose(); $script:RunspacePool = $null }
        Write-Log '=== All servers finished ===' ([System.Drawing.Color]::White)
        $btnRun.Enabled     = $true
        $btnRun.Text        = 'Run Deploy'
        $btnStopSrv.Enabled = $false
        $btnStopAll.Enabled = $false
    }
})
$Timer.Start()

# ---------------------------------------------------------------------
#  Run Deploy
# ---------------------------------------------------------------------
$btnRun.Add_Click({
    $checked = @($lvServers.Items | Where-Object { $_.Checked })
    if ($checked.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show('לא נבחרו שרתים', 'Deploy Tool',
            [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
        return
    }
    if (-not (Test-Path $txtSrc.Text)) {
        [System.Windows.Forms.MessageBox]::Show('תיקיית המקור לא קיימת', 'Deploy Tool',
            [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
        return
    }

    $confirmMsg = "לבצע גיבוי + העתקה ל-$($checked.Count) שרתים?"
    if ($chkIIS.Checked) { $confirmMsg += "`n`nIIS ייעצר בכל שרת לפני ההעתקה ויופעל מחדש בסיום." }
    if ($Config.MirrorMode) { $confirmMsg += "`n`nשים לב: MirrorMode פעיל - קבצים עודפים ביעד יימחקו!" }
    $ans = [System.Windows.Forms.MessageBox]::Show($confirmMsg, 'Deploy Tool',
        [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
    if ($ans -ne 'Yes') { return }

    $btnRun.Enabled = $false
    $btnRun.Text    = 'Running...'
    foreach ($i in $lvServers.Items) { $i.SubItems[3].Text = '' }

    # runtime config snapshot from the GUI
    $cfgRun = [pscustomobject]@{
        SourceFolder       = $txtSrc.Text
        RemoteTargetFolder = $txtTgt.Text
        BackupFolder       = $txtBak.Text
        RobocopyThreads    = [int]$Config.RobocopyThreads
        MirrorMode         = [bool]$Config.MirrorMode
        StopIIS            = [bool]$chkIIS.Checked
        UseWinRM           = [bool]$Config.UseWinRM
    }

    Write-Log "=== Starting deploy to $($checked.Count) server(s), MaxParallel=$($Config.MaxParallel) ===" ([System.Drawing.Color]::White)

    $script:RunspacePool = [runspacefactory]::CreateRunspacePool(1, [int]$Config.MaxParallel)
    $script:RunspacePool.Open()
    $script:Jobs = @()

    foreach ($item in $checked) {
        $srv = [pscustomobject]@{
            Name = $item.SubItems[0].Text
            IP   = $item.SubItems[1].Text
        }
        $item.SubItems[3].Text = 'Queued'

        $ps = [powershell]::Create()
        $ps.RunspacePool = $script:RunspacePool
        [void]$ps.AddScript($Worker).AddArgument($srv).AddArgument($cfgRun).AddArgument($Sync)
        $script:Jobs += [pscustomobject]@{ PS = $ps; Handle = $ps.BeginInvoke(); Server = $srv.Name; Stopped = $false }
    }
    $btnStopSrv.Enabled = $true
    $btnStopAll.Enabled = $true
})

# ---------------------------------------------------------------------
#  Stop buttons
# ---------------------------------------------------------------------
function Stop-Job-ForServer {
    param($Job)
    if ($Job.Stopped -or $Job.Handle.IsCompleted) { return }
    $Job.Stopped = $true
    try { [void]$Job.PS.BeginStop($null, $null) } catch {}
    foreach ($i in $lvServers.Items) {
        if ($i.SubItems[0].Text -eq $Job.Server) { $i.SubItems[3].Text = 'Cancelled' }
    }
    Write-Log "[$($Job.Server)] Cancelled by user" ([System.Drawing.Color]::Orange)
    if ($chkIIS.Checked) {
        Write-Log "[$($Job.Server)] WARNING: if IIS was already stopped on this server, verify it and start it manually (iisreset $($Job.Server) /start)" ([System.Drawing.Color]::Red)
    }
}

$btnStopSrv.Add_Click({
    if ($lvServers.SelectedItems.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show('סמן שורה של שרת ברשימה (קליק על השורה, לא על ה-checkbox) ואז לחץ שוב', 'Deploy Tool',
            [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
        return
    }
    $name = $lvServers.SelectedItems[0].SubItems[0].Text
    $job  = $script:Jobs | Where-Object { $_.Server -eq $name } | Select-Object -First 1
    if (-not $job) {
        Write-Log "[$name] No running task for this server" ([System.Drawing.Color]::Gray)
        return
    }
    Stop-Job-ForServer $job
})

$btnStopAll.Add_Click({
    $running = @($script:Jobs | Where-Object { -not $_.Handle.IsCompleted -and -not $_.Stopped })
    if ($running.Count -eq 0) { return }
    $ans = [System.Windows.Forms.MessageBox]::Show("לעצור את כל התהליכים ($($running.Count) פעילים)?", 'Deploy Tool',
        [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
    if ($ans -ne 'Yes') { return }
    Write-Log '=== STOP ALL requested by user ===' ([System.Drawing.Color]::Red)
    foreach ($j in $running) { Stop-Job-ForServer $j }
})

# ---------------------------------------------------------------------
$Form.Add_FormClosing({
    $Timer.Stop()
    foreach ($j in $script:Jobs) { try { $j.PS.Stop(); $j.PS.Dispose() } catch {} }
    if ($script:RunspacePool) { try { $script:RunspacePool.Close(); $script:RunspacePool.Dispose() } catch {} }
})

Write-Log "Config loaded from $ConfigPath" ([System.Drawing.Color]::Cyan)
[void]$Form.ShowDialog()
