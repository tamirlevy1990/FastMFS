# Search-MedconLogs-GUI.ps1
# Windows PowerShell 5.1+ / PowerShell 7+ (WinForms)
# ──────────────────────────────────────────────────────
# v6 — Date range + Smart zip filtering + Activity log + Download Logs
# Run:  powershell -ExecutionPolicy Bypass -File ".\Search-MedconLogs-GUI.ps1"
# ──────────────────────────────────────────────────────

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.IO.Compression

[System.Windows.Forms.Application]::EnableVisualStyles()

# ══════════════════════════════════════════════
#  COLORS & FONTS
# ══════════════════════════════════════════════
$accentColor    = [System.Drawing.Color]::FromArgb(0, 103, 192)
$successColor   = [System.Drawing.Color]::FromArgb(16, 150, 72)
$dangerColor    = [System.Drawing.Color]::FromArgb(200, 50, 50)
$archiveColor   = [System.Drawing.Color]::FromArgb(180, 120, 0)
$bgColor        = [System.Drawing.Color]::FromArgb(243, 243, 246)
$panelBg        = [System.Drawing.Color]::White
$gridAltRow     = [System.Drawing.Color]::FromArgb(245, 248, 252)
$gridHeaderBg   = [System.Drawing.Color]::FromArgb(50, 55, 65)
$gridHeaderFg   = [System.Drawing.Color]::White
$inputBg        = [System.Drawing.Color]::FromArgb(252, 252, 255)
$fontNormal     = New-Object System.Drawing.Font('Segoe UI', 9.5)
$fontBold       = New-Object System.Drawing.Font('Segoe UI', 9.5, [System.Drawing.FontStyle]::Bold)
$fontSmall      = New-Object System.Drawing.Font('Segoe UI', 8.5)
$fontTitle      = New-Object System.Drawing.Font('Segoe UI Semibold', 11)
$fontMono       = New-Object System.Drawing.Font('Consolas', 9)
$labelColor     = [System.Drawing.Color]::FromArgb(60, 60, 60)
$grayBtn        = [System.Drawing.Color]::FromArgb(140, 140, 140)

# ══════════════════════════════════════════════
#  FORM
# ══════════════════════════════════════════════
$form = New-Object System.Windows.Forms.Form
$form.Text          = 'Medcon Log Search'
$form.Size          = New-Object System.Drawing.Size(1100, 950)
$form.StartPosition = 'CenterScreen'
$form.MinimumSize   = New-Object System.Drawing.Size(900, 800)
$form.Font          = $fontNormal
$form.BackColor     = $bgColor
try { $form.Icon = [System.Drawing.SystemIcons]::Application } catch {}

# ══════════════════════════════════════════════
#  TITLE BAR
# ══════════════════════════════════════════════
$panelTitle = New-Object System.Windows.Forms.Panel
$panelTitle.Dock      = 'Top'
$panelTitle.Height    = 48
$panelTitle.BackColor = $gridHeaderBg
$form.Controls.Add($panelTitle)

$lblTitle = New-Object System.Windows.Forms.Label
$lblTitle.Text      = 'Medcon Framework Shell  —  Log Search'
$lblTitle.Font      = $fontTitle
$lblTitle.ForeColor = [System.Drawing.Color]::White
$lblTitle.Location  = New-Object System.Drawing.Point(18, 12)
$lblTitle.AutoSize  = $true
$panelTitle.Controls.Add($lblTitle)

# ══════════════════════════════════════════════
#  INPUT PANEL
# ══════════════════════════════════════════════
$panelInput = New-Object System.Windows.Forms.Panel
$panelInput.Dock      = 'Top'
$panelInput.Height    = 255
$panelInput.BackColor = $panelBg
$panelInput.Padding   = New-Object System.Windows.Forms.Padding(0, 0, 0, 5)
$form.Controls.Add($panelInput)

# ── Row 1: Date Range ──
$lblDateFrom = New-Object System.Windows.Forms.Label
$lblDateFrom.Text      = 'Date From'
$lblDateFrom.Font      = $fontNormal
$lblDateFrom.ForeColor = $labelColor
$lblDateFrom.Location  = New-Object System.Drawing.Point(18, 15)
$lblDateFrom.AutoSize  = $true
$panelInput.Controls.Add($lblDateFrom)

$dtpDateFrom = New-Object System.Windows.Forms.DateTimePicker
$dtpDateFrom.Font     = $fontNormal
$dtpDateFrom.Location = New-Object System.Drawing.Point(130, 12)
$dtpDateFrom.Size     = New-Object System.Drawing.Size(160, 26)
$dtpDateFrom.Format   = 'Short'
$dtpDateFrom.Value    = (Get-Date).Date
$panelInput.Controls.Add($dtpDateFrom)

$lblDateTo = New-Object System.Windows.Forms.Label
$lblDateTo.Text      = 'To'
$lblDateTo.Font      = $fontNormal
$lblDateTo.ForeColor = $labelColor
$lblDateTo.Location  = New-Object System.Drawing.Point(300, 15)
$lblDateTo.AutoSize  = $true
$panelInput.Controls.Add($lblDateTo)

$dtpDateTo = New-Object System.Windows.Forms.DateTimePicker
$dtpDateTo.Font     = $fontNormal
$dtpDateTo.Location = New-Object System.Drawing.Point(325, 12)
$dtpDateTo.Size     = New-Object System.Drawing.Size(160, 26)
$dtpDateTo.Format   = 'Short'
$dtpDateTo.Value    = (Get-Date).Date
$panelInput.Controls.Add($dtpDateTo)

$lblDayCount = New-Object System.Windows.Forms.Label
$lblDayCount.Text      = '(1 day)'
$lblDayCount.Font      = $fontSmall
$lblDayCount.ForeColor = [System.Drawing.Color]::Gray
$lblDayCount.Location  = New-Object System.Drawing.Point(492, 16)
$lblDayCount.AutoSize  = $true
$panelInput.Controls.Add($lblDayCount)

# ── Row 2: Search String ──
$lblSearch = New-Object System.Windows.Forms.Label
$lblSearch.Text      = 'Search String'
$lblSearch.Font      = $fontNormal
$lblSearch.ForeColor = $labelColor
$lblSearch.Location  = New-Object System.Drawing.Point(18, 50)
$lblSearch.AutoSize  = $true
$panelInput.Controls.Add($lblSearch)

$txtSearch = New-Object System.Windows.Forms.TextBox
$txtSearch.Font      = $fontNormal
$txtSearch.Location  = New-Object System.Drawing.Point(130, 47)
$txtSearch.Size      = New-Object System.Drawing.Size(370, 26)
$txtSearch.BackColor = $inputBg
$panelInput.Controls.Add($txtSearch)

# ── Row 3: Output Folder ──
$lblOutput = New-Object System.Windows.Forms.Label
$lblOutput.Text      = 'Output Folder'
$lblOutput.Font      = $fontNormal
$lblOutput.ForeColor = $labelColor
$lblOutput.Location  = New-Object System.Drawing.Point(18, 85)
$lblOutput.AutoSize  = $true
$panelInput.Controls.Add($lblOutput)

$txtOutput = New-Object System.Windows.Forms.TextBox
$txtOutput.Font      = $fontNormal
$txtOutput.Location  = New-Object System.Drawing.Point(130, 82)
$txtOutput.Size      = New-Object System.Drawing.Size(310, 26)
$txtOutput.BackColor = $inputBg
$txtOutput.Text      = [System.Environment]::GetFolderPath('Desktop')
$panelInput.Controls.Add($txtOutput)

$btnBrowseOutput = New-Object System.Windows.Forms.Button
$btnBrowseOutput.Text      = 'Browse'
$btnBrowseOutput.Font      = $fontSmall
$btnBrowseOutput.ForeColor = [System.Drawing.Color]::White
$btnBrowseOutput.BackColor = $grayBtn
$btnBrowseOutput.Location  = New-Object System.Drawing.Point(448, 81)
$btnBrowseOutput.Size      = New-Object System.Drawing.Size(55, 26)
$btnBrowseOutput.FlatStyle = 'Flat'
$btnBrowseOutput.FlatAppearance.BorderSize = 0
$btnBrowseOutput.Cursor    = [System.Windows.Forms.Cursors]::Hand
$panelInput.Controls.Add($btnBrowseOutput)

# ── Buttons ──
$btnSearch = New-Object System.Windows.Forms.Button
$btnSearch.Text      = 'Search'
$btnSearch.Font      = $fontBold
$btnSearch.ForeColor = [System.Drawing.Color]::White
$btnSearch.BackColor = $successColor
$btnSearch.Location  = New-Object System.Drawing.Point(130, 120)
$btnSearch.Size      = New-Object System.Drawing.Size(100, 32)
$btnSearch.FlatStyle = 'Flat'
$btnSearch.FlatAppearance.BorderSize = 0
$btnSearch.Cursor    = [System.Windows.Forms.Cursors]::Hand
$panelInput.Controls.Add($btnSearch)

$btnExport = New-Object System.Windows.Forms.Button
$btnExport.Text      = 'Export CSV'
$btnExport.Font      = $fontNormal
$btnExport.ForeColor = [System.Drawing.Color]::White
$btnExport.BackColor = $accentColor
$btnExport.Location  = New-Object System.Drawing.Point(240, 120)
$btnExport.Size      = New-Object System.Drawing.Size(100, 32)
$btnExport.FlatStyle = 'Flat'
$btnExport.FlatAppearance.BorderSize = 0
$btnExport.Cursor    = [System.Windows.Forms.Cursors]::Hand
$btnExport.Enabled   = $false
$panelInput.Controls.Add($btnExport)

$btnClear = New-Object System.Windows.Forms.Button
$btnClear.Text      = 'Clear'
$btnClear.Font      = $fontNormal
$btnClear.ForeColor = [System.Drawing.Color]::White
$btnClear.BackColor = $grayBtn
$btnClear.Location  = New-Object System.Drawing.Point(350, 120)
$btnClear.Size      = New-Object System.Drawing.Size(80, 32)
$btnClear.FlatStyle = 'Flat'
$btnClear.FlatAppearance.BorderSize = 0
$btnClear.Cursor    = [System.Windows.Forms.Cursors]::Hand
$panelInput.Controls.Add($btnClear)

$btnDownload = New-Object System.Windows.Forms.Button
$btnDownload.Text      = 'Download Logs'
$btnDownload.Font      = $fontBold
$btnDownload.ForeColor = [System.Drawing.Color]::White
$btnDownload.BackColor = $archiveColor
$btnDownload.Location  = New-Object System.Drawing.Point(440, 120)
$btnDownload.Size      = New-Object System.Drawing.Size(120, 32)
$btnDownload.FlatStyle = 'Flat'
$btnDownload.FlatAppearance.BorderSize = 0
$btnDownload.Cursor    = [System.Windows.Forms.Cursors]::Hand
$panelInput.Controls.Add($btnDownload)

$lblArchiveNote = New-Object System.Windows.Forms.Label
$lblArchiveNote.Text      = 'Search scans LogEx + LogArchive. Download copies the full log files locally.'
$lblArchiveNote.Font      = $fontSmall
$lblArchiveNote.ForeColor = [System.Drawing.Color]::Gray
$lblArchiveNote.Location  = New-Object System.Drawing.Point(130, 158)
$lblArchiveNote.AutoSize  = $true
$panelInput.Controls.Add($lblArchiveNote)

# ── Right side: Station Names ──
$separator = New-Object System.Windows.Forms.Panel
$separator.Location  = New-Object System.Drawing.Point(530, 10)
$separator.Size      = New-Object System.Drawing.Size(1, 230)
$separator.BackColor = [System.Drawing.Color]::FromArgb(220, 220, 225)
$panelInput.Controls.Add($separator)

$lblStations = New-Object System.Windows.Forms.Label
$lblStations.Text      = 'Station Names'
$lblStations.Font      = $fontBold
$lblStations.ForeColor = $accentColor
$lblStations.Location  = New-Object System.Drawing.Point(550, 12)
$lblStations.AutoSize  = $true
$panelInput.Controls.Add($lblStations)

$lblStationHint = New-Object System.Windows.Forms.Label
$lblStationHint.Text      = '(one per line)'
$lblStationHint.Font      = $fontSmall
$lblStationHint.ForeColor = [System.Drawing.Color]::Gray
$lblStationHint.Location  = New-Object System.Drawing.Point(660, 14)
$lblStationHint.AutoSize  = $true
$panelInput.Controls.Add($lblStationHint)

$txtStations = New-Object System.Windows.Forms.TextBox
$txtStations.Font           = $fontMono
$txtStations.Location       = New-Object System.Drawing.Point(550, 35)
$txtStations.Size           = New-Object System.Drawing.Size(310, 170)
$txtStations.Anchor         = 'Top,Left,Right'
$txtStations.Multiline      = $true
$txtStations.ScrollBars     = 'Vertical'
$txtStations.BackColor      = $inputBg
$txtStations.WordWrap       = $false
$txtStations.AcceptsReturn  = $true
$panelInput.Controls.Add($txtStations)

$lblStationCount = New-Object System.Windows.Forms.Label
$lblStationCount.Text      = '0 stations'
$lblStationCount.Font      = $fontSmall
$lblStationCount.ForeColor = [System.Drawing.Color]::Gray
$lblStationCount.Location  = New-Object System.Drawing.Point(550, 210)
$lblStationCount.AutoSize  = $true
$panelInput.Controls.Add($lblStationCount)

$btnLoadFile = New-Object System.Windows.Forms.Button
$btnLoadFile.Text      = 'Load from File'
$btnLoadFile.Font      = $fontSmall
$btnLoadFile.ForeColor = [System.Drawing.Color]::White
$btnLoadFile.BackColor = $accentColor
$btnLoadFile.Location  = New-Object System.Drawing.Point(660, 207)
$btnLoadFile.Size      = New-Object System.Drawing.Size(95, 25)
$btnLoadFile.FlatStyle = 'Flat'
$btnLoadFile.FlatAppearance.BorderSize = 0
$btnLoadFile.Cursor    = [System.Windows.Forms.Cursors]::Hand
$panelInput.Controls.Add($btnLoadFile)

$btnClearStations = New-Object System.Windows.Forms.Button
$btnClearStations.Text      = 'Clear'
$btnClearStations.Font      = $fontSmall
$btnClearStations.ForeColor = [System.Drawing.Color]::White
$btnClearStations.BackColor = $grayBtn
$btnClearStations.Location  = New-Object System.Drawing.Point(762, 207)
$btnClearStations.Size      = New-Object System.Drawing.Size(55, 25)
$btnClearStations.FlatStyle = 'Flat'
$btnClearStations.FlatAppearance.BorderSize = 0
$btnClearStations.Cursor    = [System.Windows.Forms.Cursors]::Hand
$panelInput.Controls.Add($btnClearStations)

$btnSortStations = New-Object System.Windows.Forms.Button
$btnSortStations.Text      = 'Sort A-Z'
$btnSortStations.Font      = $fontSmall
$btnSortStations.ForeColor = [System.Drawing.Color]::White
$btnSortStations.BackColor = $grayBtn
$btnSortStations.Location  = New-Object System.Drawing.Point(824, 207)
$btnSortStations.Size      = New-Object System.Drawing.Size(55, 25)
$btnSortStations.FlatStyle = 'Flat'
$btnSortStations.FlatAppearance.BorderSize = 0
$btnSortStations.Cursor    = [System.Windows.Forms.Cursors]::Hand
$panelInput.Controls.Add($btnSortStations)

# ══════════════════════════════════════════════
#  PROGRESS & FILTER BAR
# ══════════════════════════════════════════════
$panelMid = New-Object System.Windows.Forms.Panel
$panelMid.Dock      = 'Top'
$panelMid.Height    = 42
$panelMid.BackColor = $bgColor
$form.Controls.Add($panelMid)

$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(15, 10)
$progressBar.Size     = New-Object System.Drawing.Size(360, 22)
$progressBar.Style    = 'Continuous'
$panelMid.Controls.Add($progressBar)

$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Text      = 'Ready'
$lblStatus.Font      = $fontSmall
$lblStatus.Location  = New-Object System.Drawing.Point(385, 13)
$lblStatus.AutoSize  = $true
$lblStatus.ForeColor = [System.Drawing.Color]::Gray
$panelMid.Controls.Add($lblStatus)

$lblFilter = New-Object System.Windows.Forms.Label
$lblFilter.Text      = 'Filter results:'
$lblFilter.Font      = $fontSmall
$lblFilter.ForeColor = [System.Drawing.Color]::FromArgb(80, 80, 80)
$lblFilter.Location  = New-Object System.Drawing.Point(700, 13)
$lblFilter.AutoSize  = $true
$lblFilter.Anchor    = 'Top,Right'
$panelMid.Controls.Add($lblFilter)

$txtFilter = New-Object System.Windows.Forms.TextBox
$txtFilter.Font      = $fontSmall
$txtFilter.Location  = New-Object System.Drawing.Point(780, 10)
$txtFilter.Size      = New-Object System.Drawing.Size(200, 22)
$txtFilter.Anchor    = 'Top,Right'
$txtFilter.BackColor = $inputBg
$panelMid.Controls.Add($txtFilter)

# ══════════════════════════════════════════════
#  ACTIVITY LOG PANEL
# ══════════════════════════════════════════════
$panelActivity = New-Object System.Windows.Forms.Panel
$panelActivity.Dock      = 'Top'
$panelActivity.Height    = 150
$panelActivity.BackColor = $bgColor
$form.Controls.Add($panelActivity)

$lblActivityHeader = New-Object System.Windows.Forms.Label
$lblActivityHeader.Text      = 'Activity Log'
$lblActivityHeader.Font      = $fontBold
$lblActivityHeader.ForeColor = $labelColor
$lblActivityHeader.Location  = New-Object System.Drawing.Point(15, 5)
$lblActivityHeader.AutoSize  = $true
$panelActivity.Controls.Add($lblActivityHeader)

$btnClearLog = New-Object System.Windows.Forms.Button
$btnClearLog.Text      = 'Clear Log'
$btnClearLog.Font      = $fontSmall
$btnClearLog.ForeColor = [System.Drawing.Color]::White
$btnClearLog.BackColor = $grayBtn
$btnClearLog.Location  = New-Object System.Drawing.Point(95, 3)
$btnClearLog.Size      = New-Object System.Drawing.Size(70, 22)
$btnClearLog.FlatStyle = 'Flat'
$btnClearLog.FlatAppearance.BorderSize = 0
$btnClearLog.Cursor    = [System.Windows.Forms.Cursors]::Hand
$panelActivity.Controls.Add($btnClearLog)

$btnToggleLog = New-Object System.Windows.Forms.Button
$btnToggleLog.Text      = 'Hide Log'
$btnToggleLog.Font      = $fontSmall
$btnToggleLog.ForeColor = [System.Drawing.Color]::White
$btnToggleLog.BackColor = $grayBtn
$btnToggleLog.Location  = New-Object System.Drawing.Point(170, 3)
$btnToggleLog.Size      = New-Object System.Drawing.Size(70, 22)
$btnToggleLog.FlatStyle = 'Flat'
$btnToggleLog.FlatAppearance.BorderSize = 0
$btnToggleLog.Cursor    = [System.Windows.Forms.Cursors]::Hand
$panelActivity.Controls.Add($btnToggleLog)

$btnSaveLog = New-Object System.Windows.Forms.Button
$btnSaveLog.Text      = 'Save Log'
$btnSaveLog.Font      = $fontSmall
$btnSaveLog.ForeColor = [System.Drawing.Color]::White
$btnSaveLog.BackColor = $accentColor
$btnSaveLog.Location  = New-Object System.Drawing.Point(245, 3)
$btnSaveLog.Size      = New-Object System.Drawing.Size(70, 22)
$btnSaveLog.FlatStyle = 'Flat'
$btnSaveLog.FlatAppearance.BorderSize = 0
$btnSaveLog.Cursor    = [System.Windows.Forms.Cursors]::Hand
$panelActivity.Controls.Add($btnSaveLog)

$txtActivity = New-Object System.Windows.Forms.RichTextBox
$txtActivity.Location    = New-Object System.Drawing.Point(15, 28)
$txtActivity.Size        = New-Object System.Drawing.Size(1050, 115)
$txtActivity.Anchor      = 'Top,Left,Right,Bottom'
$txtActivity.Font        = $fontMono
$txtActivity.BackColor   = [System.Drawing.Color]::FromArgb(30, 30, 35)
$txtActivity.ForeColor   = [System.Drawing.Color]::FromArgb(220, 220, 220)
$txtActivity.ReadOnly    = $true
$txtActivity.ScrollBars  = 'Vertical'
$txtActivity.BorderStyle = 'None'
$panelActivity.Controls.Add($txtActivity)

$btnClearLog.Add_Click({ $txtActivity.Clear() })

$btnToggleLog.Add_Click({
    if ($panelActivity.Height -gt 30) {
        $panelActivity.Height = 30
        $btnToggleLog.Text = 'Show Log'
    } else {
        $panelActivity.Height = 150
        $btnToggleLog.Text = 'Hide Log'
    }
})

$btnSaveLog.Add_Click({
    $sfd = New-Object System.Windows.Forms.SaveFileDialog
    $sfd.Title    = 'Save Activity Log'
    $sfd.Filter   = 'Text files (*.txt)|*.txt'
    $sfd.FileName = "ActivityLog_$((Get-Date).ToString('yyyyMMdd_HHmmss')).txt"
    if ($sfd.ShowDialog() -eq 'OK') {
        [System.IO.File]::WriteAllText($sfd.FileName, $txtActivity.Text, [System.Text.Encoding]::UTF8)
    }
})

# ══════════════════════════════════════════════
#  RESULTS GRID
# ══════════════════════════════════════════════
$dgv = New-Object System.Windows.Forms.DataGridView
$dgv.Dock                         = 'Fill'
$dgv.ReadOnly                     = $true
$dgv.AllowUserToAddRows           = $false
$dgv.AllowUserToDeleteRows        = $false
$dgv.AllowUserToResizeRows        = $false
$dgv.AutoSizeColumnsMode          = 'None'
$dgv.ScrollBars                   = 'Both'
$dgv.SelectionMode                = 'FullRowSelect'
$dgv.BackgroundColor              = [System.Drawing.Color]::White
$dgv.GridColor                    = [System.Drawing.Color]::FromArgb(225, 228, 232)
$dgv.BorderStyle                  = 'None'
$dgv.CellBorderStyle              = 'SingleHorizontal'
$dgv.DefaultCellStyle.Font        = $fontMono
$dgv.DefaultCellStyle.Padding     = New-Object System.Windows.Forms.Padding(4, 3, 4, 3)
$dgv.DefaultCellStyle.SelectionBackColor = [System.Drawing.Color]::FromArgb(210, 225, 245)
$dgv.DefaultCellStyle.SelectionForeColor = [System.Drawing.Color]::Black
$dgv.ColumnHeadersDefaultCellStyle.BackColor = $gridHeaderBg
$dgv.ColumnHeadersDefaultCellStyle.ForeColor = $gridHeaderFg
$dgv.ColumnHeadersDefaultCellStyle.Font      = $fontBold
$dgv.ColumnHeadersDefaultCellStyle.Padding   = New-Object System.Windows.Forms.Padding(4, 5, 4, 5)
$dgv.ColumnHeadersHeight          = 34
$dgv.ColumnHeadersHeightSizeMode  = 'DisableResizing'
$dgv.EnableHeadersVisualStyles    = $false
$dgv.RowHeadersVisible            = $false
$dgv.RowTemplate.Height           = 26
$dgv.AlternatingRowsDefaultCellStyle.BackColor = $gridAltRow
$form.Controls.Add($dgv)

@(
    @{ Name = 'Station';     Header = 'Station';       Width = 140  },
    @{ Name = 'Source';      Header = 'Source';        Width = 80   },
    @{ Name = 'FileName';    Header = 'File Name';     Width = 250  },
    @{ Name = 'FileDate';    Header = 'File Date';     Width = 130  },
    @{ Name = 'LineNumber';  Header = 'Line #';        Width = 65   },
    @{ Name = 'MatchedLine'; Header = 'Matched Line';  Width = 2000 }
) | ForEach-Object {
    $col = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $col.Name         = $_.Name
    $col.HeaderText   = $_.Header
    $col.Width        = $_.Width
    $col.MinimumWidth = 50
    $col.SortMode     = 'Automatic'
    $col.Resizable    = 'True'
    [void]$dgv.Columns.Add($col)
}

# ══════════════════════════════════════════════
#  STATUS BAR
# ══════════════════════════════════════════════
$statusStrip = New-Object System.Windows.Forms.StatusStrip
$statusStrip.BackColor = $gridHeaderBg

$statusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
$statusLabel.Text      = '  No results'
$statusLabel.ForeColor = [System.Drawing.Color]::FromArgb(200, 200, 200)
$statusLabel.Font      = $fontSmall
$statusLabel.Spring    = $true
$statusLabel.TextAlign = 'MiddleLeft'
[void]$statusStrip.Items.Add($statusLabel)

$statusRowCount = New-Object System.Windows.Forms.ToolStripStatusLabel
$statusRowCount.Text      = 'Rows: 0'
$statusRowCount.ForeColor = [System.Drawing.Color]::FromArgb(200, 200, 200)
$statusRowCount.Font      = $fontSmall
$statusRowCount.Alignment = 'Right'
[void]$statusStrip.Items.Add($statusRowCount)

$form.Controls.Add($statusStrip)

# ══════════════════════════════════════════════
#  RIGHT-CLICK MENU
# ══════════════════════════════════════════════
$ctxMenu = New-Object System.Windows.Forms.ContextMenuStrip
$ctxMenu.Font = $fontNormal
$menuCopyLine = $ctxMenu.Items.Add('Copy Matched Line')
$menuCopyRow  = $ctxMenu.Items.Add('Copy Entire Row')
$ctxMenu.Items.Add('-')
$menuCopyAll  = $ctxMenu.Items.Add('Copy All Visible Rows')
$dgv.ContextMenuStrip = $ctxMenu

$menuCopyLine.Add_Click({
    if ($dgv.CurrentRow) {
        $text = $dgv.CurrentRow.Cells['MatchedLine'].Value
        if ($text) { [System.Windows.Forms.Clipboard]::SetText([string]$text) }
    }
})
$menuCopyRow.Add_Click({
    if ($dgv.CurrentRow) {
        $parts = @(); foreach ($cell in $dgv.CurrentRow.Cells) { $parts += $cell.Value }
        [System.Windows.Forms.Clipboard]::SetText(($parts -join "`t"))
    }
})
$menuCopyAll.Add_Click({
    $sb = New-Object System.Text.StringBuilder
    $headers = @(); foreach ($col in $dgv.Columns) { $headers += $col.HeaderText }
    [void]$sb.AppendLine($headers -join "`t")
    foreach ($row in $dgv.Rows) {
        if ($row.Visible) {
            $vals = @(); foreach ($cell in $row.Cells) { $vals += $cell.Value }
            [void]$sb.AppendLine($vals -join "`t")
        }
    }
    [System.Windows.Forms.Clipboard]::SetText($sb.ToString())
    $statusLabel.Text = "  Copied to clipboard"
})

# ══════════════════════════════════════════════
#  TOOLTIPS
# ══════════════════════════════════════════════
$toolTip = New-Object System.Windows.Forms.ToolTip
$toolTip.SetToolTip($dtpDateFrom,  'Start date - files modified on or after this date')
$toolTip.SetToolTip($dtpDateTo,    'End date - files modified on or before this date (same as From for single day)')
$toolTip.SetToolTip($txtSearch,    'Case-insensitive text search inside log lines')
$toolTip.SetToolTip($txtStations,  'Type or paste station hostnames, one per line')
$toolTip.SetToolTip($btnLoadFile,  'Load station names from a .txt file')
$toolTip.SetToolTip($txtFilter,    'Type here to filter results across all columns')
$toolTip.SetToolTip($txtOutput,    'Folder where the CSV / downloaded logs will be saved')
$toolTip.SetToolTip($btnDownload,  'Copy the full Medcon.Framework.Shell* log files for the selected date range to a local folder')

# ══════════════════════════════════════════════
#  SHARED STATE
# ══════════════════════════════════════════════
$script:searchResults = $null
$script:csvPath       = $null

# ══════════════════════════════════════════════
#  HELPERS
# ══════════════════════════════════════════════
function Write-Activity {
    param(
        [string]$Message,
        [ValidateSet('Info','Success','Warning','Error','Step','Detail')]
        [string]$Level = 'Info'
    )

    $timestamp = (Get-Date).ToString('HH:mm:ss.fff')
    $prefix = switch ($Level) {
        'Success' { '[OK]    ' }
        'Warning' { '[WARN]  ' }
        'Error'   { '[ERR]   ' }
        'Step'    { '[STEP]  ' }
        'Detail'  { '   ->   ' }
        default   { '[INFO]  ' }
    }

    $color = switch ($Level) {
        'Success' { [System.Drawing.Color]::FromArgb(100, 220, 100) }
        'Warning' { [System.Drawing.Color]::FromArgb(255, 200, 80)  }
        'Error'   { [System.Drawing.Color]::FromArgb(255, 110, 110) }
        'Step'    { [System.Drawing.Color]::FromArgb(100, 180, 255) }
        'Detail'  { [System.Drawing.Color]::FromArgb(170, 170, 170) }
        default   { [System.Drawing.Color]::FromArgb(220, 220, 220) }
    }

    $txtActivity.SelectionStart  = $txtActivity.TextLength
    $txtActivity.SelectionLength = 0
    $txtActivity.SelectionColor  = [System.Drawing.Color]::FromArgb(130, 130, 130)
    $txtActivity.AppendText("$timestamp ")
    $txtActivity.SelectionColor  = $color
    $txtActivity.AppendText("$prefix$Message`r`n")
    $txtActivity.ScrollToCaret()
    [System.Windows.Forms.Application]::DoEvents()
}

function Get-StationList {
    $rawText = $txtStations.Text
    $tokens  = $rawText -split '[\s,;]+'

    $list = [System.Collections.Generic.List[string]]::new()
    $seen = @{}
    foreach ($token in $tokens) {
        $name = $token.Trim()
        if ($name -and -not $name.StartsWith('#')) {
            $key = $name.ToLowerInvariant()
            if (-not $seen.ContainsKey($key)) {
                $seen[$key] = $true
                $list.Add($name)
            }
        }
    }
    return $list
}

function Update-StationCount {
    $list = @(Get-StationList)
    $lblStationCount.Text      = "$($list.Count) station(s)"
    $lblStationCount.ForeColor = if ($list.Count -gt 0) { $successColor } else { [System.Drawing.Color]::Gray }
}

function Update-DayCount {
    $from = $dtpDateFrom.Value.Date
    $to   = $dtpDateTo.Value.Date
    if ($to -lt $from) {
        $lblDayCount.Text      = '(invalid range)'
        $lblDayCount.ForeColor = $dangerColor
    } else {
        $days = (($to - $from).Days + 1)
        $word = if ($days -eq 1) { 'day' } else { 'days' }
        $lblDayCount.Text      = "($days $word)"
        $lblDayCount.ForeColor = [System.Drawing.Color]::Gray
    }
}

function Update-RowCount {
    $visible = 0
    foreach ($row in $dgv.Rows) { if ($row.Visible) { $visible++ } }
    $total = $dgv.Rows.Count
    if ($visible -eq $total) { $statusRowCount.Text = "Rows: $total" }
    else { $statusRowCount.Text = "Showing: $visible / $total" }
}

function Get-ZipDateRange {
    param([string]$ZipName)

    $pattern = '(\d{1,2}\.\d{1,2}\.\d{4})\s+\d{1,2}-\d{1,2}-\d{1,2}\s+to\s+(\d{1,2}\.\d{1,2}\.\d{4})\s+\d{1,2}-\d{1,2}-\d{1,2}'

    if ($ZipName -match $pattern) {
        $startStr = $matches[1]
        $endStr   = $matches[2]

        [datetime]$zipStart = [datetime]::MinValue
        [datetime]$zipEnd   = [datetime]::MinValue

        $formats = [string[]]@('d.M.yyyy', 'dd.MM.yyyy', 'M.d.yyyy', 'dd.M.yyyy', 'd.MM.yyyy')
        $culture = [System.Globalization.CultureInfo]::InvariantCulture
        $style   = [System.Globalization.DateTimeStyles]::None

        $okStart = [datetime]::TryParseExact($startStr, $formats, $culture, $style, [ref]$zipStart)
        $okEnd   = [datetime]::TryParseExact($endStr,   $formats, $culture, $style, [ref]$zipEnd)

        if ($okStart -and $okEnd) {
            return @{ Start = $zipStart.Date; End = $zipEnd.Date }
        }
    }
    return $null
}

function Test-DateRangesOverlap {
    param(
        [datetime]$SearchFrom,
        [datetime]$SearchTo,
        [datetime]$ZipFrom,
        [datetime]$ZipTo
    )
    return ($ZipFrom -le $SearchTo) -and ($ZipTo -ge $SearchFrom)
}

function Search-PlainFiles {
    param(
        [string]$UncPath,
        [string]$Station,
        [datetime]$DateFrom,
        [datetime]$DateTo,
        [string]$SearchStr,
        [System.Collections.Generic.List[object]]$Results,
        [System.Collections.Generic.List[string]]$Problems
    )

    Write-Activity "Checking main folder: $UncPath" -Level Detail

    if (-not (Test-Path -LiteralPath $UncPath)) {
        Write-Activity "Main folder not accessible" -Level Warning
        return $false
    }

    $allFiles = @(Get-ChildItem -LiteralPath $UncPath -File -Filter 'Medcon.Framework.Shell*' -ErrorAction SilentlyContinue)
    Write-Activity "Found $($allFiles.Count) Medcon.Framework.Shell* file(s) in main folder" -Level Detail

    $files = @($allFiles | Where-Object { $_.LastWriteTime.Date -ge $DateFrom -and $_.LastWriteTime.Date -le $DateTo })

    if ($files.Count -eq 0) {
        Write-Activity "No files matching the date range in main folder" -Level Warning
        return $false
    }

    Write-Activity "$($files.Count) file(s) match the date range - scanning..." -Level Info

    $matchCount = 0
    foreach ($file in $files) {
        $sizeMB = [math]::Round($file.Length / 1MB, 2)
        Write-Activity "Reading: $($file.Name) ($sizeMB MB, $($file.LastWriteTime.ToString('yyyy-MM-dd HH:mm')))" -Level Detail
        [System.Windows.Forms.Application]::DoEvents()

        $stream = $null
        $reader = $null
        try {
            $stream = [System.IO.FileStream]::new($file.FullName,
                [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
            $reader = [System.IO.StreamReader]::new($stream)

            $lineNumber = 0
            $fileMatches = 0
            while (($line = $reader.ReadLine()) -ne $null) {
                $lineNumber++
                if ($line.IndexOf($SearchStr, [System.StringComparison]::OrdinalIgnoreCase) -ge 0) {
                    $fileDate = $file.LastWriteTime.ToString('yyyy-MM-dd HH:mm')
                    $trimmed  = $line.Trim()
                    $Results.Add([PSCustomObject]@{
                        Station=$Station; Source='LogEx'; FileName=$file.Name
                        FileDate=$fileDate; LineNumber=$lineNumber; MatchedLine=$trimmed
                    })
                    [void]$dgv.Rows.Add($Station,'LogEx',$file.Name,$fileDate,$lineNumber,$trimmed)
                    $fileMatches++
                    $matchCount++
                }
            }
            Write-Activity "Scanned $lineNumber lines, found $fileMatches match(es)" -Level Detail
        }
        catch {
            Write-Activity "Error reading $($file.Name): $($_.Exception.Message)" -Level Error
            $Problems.Add("$Station - $($file.Name): $($_.Exception.Message)")
        }
        finally {
            if ($reader) { $reader.Dispose() }
            if ($stream) { $stream.Dispose() }
        }
    }

    Write-Activity "Main folder scan complete: $matchCount total match(es)" -Level Success
    return $true
}

function Search-ArchiveZips {
    param(
        [string]$ArchivePath,
        [string]$Station,
        [datetime]$DateFrom,
        [datetime]$DateTo,
        [string]$SearchStr,
        [System.Collections.Generic.List[object]]$Results,
        [System.Collections.Generic.List[string]]$Problems
    )

    Write-Activity "Checking archive folder: $ArchivePath" -Level Detail

    if (-not (Test-Path -LiteralPath $ArchivePath)) {
        Write-Activity "Archive folder not accessible" -Level Warning
        return $false
    }

    $zipFiles = @(Get-ChildItem -LiteralPath $ArchivePath -File -Filter '*.zip' -ErrorAction SilentlyContinue)
    Write-Activity "Found $($zipFiles.Count) zip file(s) in archive" -Level Detail

    if ($zipFiles.Count -eq 0) { return $false }

    $foundAny       = $false
    $zipsChecked    = 0
    $zipsSkipped    = 0
    $zipsUnparsable = 0

    foreach ($zipFile in $zipFiles) {
        $zipRange = Get-ZipDateRange -ZipName $zipFile.Name

        if ($null -ne $zipRange) {
            if (-not (Test-DateRangesOverlap -SearchFrom $DateFrom -SearchTo $DateTo -ZipFrom $zipRange.Start -ZipTo $zipRange.End)) {
                $zipsSkipped++
                continue
            }
        } else {
            $zipsUnparsable++
            Write-Activity "Could not parse date from: $($zipFile.Name) (will search anyway)" -Level Warning
        }

        $zipsChecked++
        $sizeMB = [math]::Round($zipFile.Length / 1MB, 2)
        Write-Activity "Opening zip: $($zipFile.Name) ($sizeMB MB)" -Level Detail
        [System.Windows.Forms.Application]::DoEvents()

        $zipStream = $null
        $archive   = $null
        try {
            $zipStream = [System.IO.FileStream]::new($zipFile.FullName,
                [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
            $archive = [System.IO.Compression.ZipArchive]::new($zipStream, [System.IO.Compression.ZipArchiveMode]::Read)

            $matchingEntries = @($archive.Entries | Where-Object {
                $_.Name -like 'Medcon.Framework.Shell*' -and
                $_.LastWriteTime.Date -ge $DateFrom -and
                $_.LastWriteTime.Date -le $DateTo
            })

            Write-Activity "Zip contains $($archive.Entries.Count) entries, $($matchingEntries.Count) match the date range" -Level Detail

            foreach ($entry in $matchingEntries) {
                $foundAny = $true
                Write-Activity "Reading entry: $($entry.Name)" -Level Detail
                [System.Windows.Forms.Application]::DoEvents()

                $entryStream = $null
                $entryReader = $null
                try {
                    $entryStream = $entry.Open()
                    $entryReader = [System.IO.StreamReader]::new($entryStream)
                    $lineNumber = 0
                    $entryMatches = 0
                    while (($line = $entryReader.ReadLine()) -ne $null) {
                        $lineNumber++
                        if ($line.IndexOf($SearchStr, [System.StringComparison]::OrdinalIgnoreCase) -ge 0) {
                            $fileDate    = $entry.LastWriteTime.ToString('yyyy-MM-dd HH:mm')
                            $trimmed     = $line.Trim()
                            $displayName = "{0} > {1}" -f $zipFile.Name, $entry.Name
                            $Results.Add([PSCustomObject]@{
                                Station=$Station; Source='Archive'; FileName=$displayName
                                FileDate=$fileDate; LineNumber=$lineNumber; MatchedLine=$trimmed
                            })
                            [void]$dgv.Rows.Add($Station,'Archive',$displayName,$fileDate,$lineNumber,$trimmed)
                            $entryMatches++
                        }
                    }
                    Write-Activity "Scanned $lineNumber lines, found $entryMatches match(es)" -Level Detail
                }
                finally {
                    if ($entryReader) { $entryReader.Dispose() }
                    if ($entryStream) { $entryStream.Dispose() }
                }
            }
        }
        catch {
            Write-Activity "Error reading $($zipFile.Name): $($_.Exception.Message)" -Level Error
            $Problems.Add("$Station - $($zipFile.Name): $($_.Exception.Message)")
        }
        finally {
            if ($archive)   { $archive.Dispose()   }
            if ($zipStream) { $zipStream.Dispose() }
        }
    }

    Write-Activity "Archive summary: $zipsChecked opened, $zipsSkipped skipped by date filter, $zipsUnparsable unparsable names" -Level Info
    return $foundAny
}

function Copy-PlainLogs {
    param(
        [string]$UncPath,
        [string]$Station,
        [datetime]$DateFrom,
        [datetime]$DateTo,
        [string]$DestFolder,
        [System.Collections.Generic.List[string]]$Problems
    )

    if (-not (Test-Path -LiteralPath $UncPath)) {
        Write-Activity "Main folder not accessible: $UncPath" -Level Warning
        return 0
    }

    $allFiles = @(Get-ChildItem -LiteralPath $UncPath -File -Filter 'Medcon.Framework.Shell*' -ErrorAction SilentlyContinue)
    $files    = @($allFiles | Where-Object { $_.LastWriteTime.Date -ge $DateFrom -and $_.LastWriteTime.Date -le $DateTo })

    Write-Activity "Main folder: $($files.Count) file(s) match date range" -Level Detail
    if ($files.Count -eq 0) { return 0 }

    if (-not (Test-Path -LiteralPath $DestFolder)) {
        [void](New-Item -ItemType Directory -Path $DestFolder -Force)
    }

    $copied = 0
    foreach ($file in $files) {
        try {
            $target = Join-Path $DestFolder $file.Name
            Copy-Item -LiteralPath $file.FullName -Destination $target -Force -ErrorAction Stop
            $sizeMB = [math]::Round($file.Length / 1MB, 2)
            Write-Activity "Copied: $($file.Name) ($sizeMB MB)" -Level Detail
            $copied++
        }
        catch {
            Write-Activity "Failed copying $($file.Name): $($_.Exception.Message)" -Level Error
            $Problems.Add("$Station - $($file.Name): $($_.Exception.Message)")
        }
        [System.Windows.Forms.Application]::DoEvents()
    }
    return $copied
}

function Extract-ArchiveLogs {
    param(
        [string]$ArchivePath,
        [string]$Station,
        [datetime]$DateFrom,
        [datetime]$DateTo,
        [string]$DestFolder,
        [System.Collections.Generic.List[string]]$Problems
    )

    if (-not (Test-Path -LiteralPath $ArchivePath)) {
        Write-Activity "Archive folder not accessible: $ArchivePath" -Level Warning
        return 0
    }

    $zipFiles = @(Get-ChildItem -LiteralPath $ArchivePath -File -Filter '*.zip' -ErrorAction SilentlyContinue)
    if ($zipFiles.Count -eq 0) { return 0 }

    $extracted = 0
    foreach ($zipFile in $zipFiles) {
        $zipRange = Get-ZipDateRange -ZipName $zipFile.Name
        if ($null -ne $zipRange) {
            if (-not (Test-DateRangesOverlap -SearchFrom $DateFrom -SearchTo $DateTo -ZipFrom $zipRange.Start -ZipTo $zipRange.End)) {
                continue
            }
        }

        $zipStream = $null
        $archive   = $null
        try {
            $zipStream = [System.IO.FileStream]::new($zipFile.FullName,
                [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
            $archive = [System.IO.Compression.ZipArchive]::new($zipStream, [System.IO.Compression.ZipArchiveMode]::Read)

            $matchingEntries = @($archive.Entries | Where-Object {
                $_.Name -like 'Medcon.Framework.Shell*' -and
                $_.LastWriteTime.Date -ge $DateFrom -and
                $_.LastWriteTime.Date -le $DateTo
            })

            if ($matchingEntries.Count -eq 0) { continue }

            $zipBaseName = [System.IO.Path]::GetFileNameWithoutExtension($zipFile.Name)
            $zipDestRoot = Join-Path $DestFolder ("Archive_" + $zipBaseName)
            if (-not (Test-Path -LiteralPath $zipDestRoot)) {
                [void](New-Item -ItemType Directory -Path $zipDestRoot -Force)
            }

            Write-Activity "Extracting from $($zipFile.Name): $($matchingEntries.Count) entry(ies)" -Level Detail

            foreach ($entry in $matchingEntries) {
                try {
                    $target = Join-Path $zipDestRoot $entry.Name
                    [System.IO.Compression.ZipFileExtensions]::ExtractToFile($entry, $target, $true)
                    Write-Activity "Extracted: $($entry.Name)" -Level Detail
                    $extracted++
                }
                catch {
                    Write-Activity "Failed extracting $($entry.Name): $($_.Exception.Message)" -Level Error
                    $Problems.Add("$Station - $($zipFile.Name) > $($entry.Name): $($_.Exception.Message)")
                }
                [System.Windows.Forms.Application]::DoEvents()
            }
        }
        catch {
            Write-Activity "Error opening $($zipFile.Name): $($_.Exception.Message)" -Level Error
            $Problems.Add("$Station - $($zipFile.Name): $($_.Exception.Message)")
        }
        finally {
            if ($archive)   { $archive.Dispose()   }
            if ($zipStream) { $zipStream.Dispose() }
        }
    }
    return $extracted
}

# ══════════════════════════════════════════════
#  EVENT HANDLERS
# ══════════════════════════════════════════════

$txtStations.Add_TextChanged({ Update-StationCount })
$dtpDateFrom.Add_ValueChanged({ Update-DayCount })
$dtpDateTo.Add_ValueChanged({ Update-DayCount })

$btnLoadFile.Add_Click({
    $ofd = New-Object System.Windows.Forms.OpenFileDialog
    $ofd.Title  = 'Select Station List File'
    $ofd.Filter = 'Text files (*.txt)|*.txt|All files (*.*)|*.*'
    if ($ofd.ShowDialog() -eq 'OK') {
        $content = Get-Content -LiteralPath $ofd.FileName -Raw
        if ($txtStations.Text.Trim()) {
            $answer = [System.Windows.Forms.MessageBox]::Show(
                "Replace existing stations or append?`n`nYes = Replace`nNo = Append",
                'Load Stations', 'YesNoCancel', 'Question')
            if ($answer -eq 'Yes')    { $txtStations.Text = $content }
            elseif ($answer -eq 'No') { $txtStations.Text = $txtStations.Text.TrimEnd() + "`r`n" + $content }
        } else {
            $txtStations.Text = $content
        }
        $txtOutput.Text = Split-Path -Parent (Resolve-Path -LiteralPath $ofd.FileName)
    }
})

$btnClearStations.Add_Click({ $txtStations.Text = '' })

$btnSortStations.Add_Click({
    $list = @(Get-StationList)
    if ($list.Count -gt 0) {
        $txtStations.Text = (($list | Sort-Object) -join "`r`n")
    }
})

$btnBrowseOutput.Add_Click({
    $fbd = New-Object System.Windows.Forms.FolderBrowserDialog
    $fbd.Description  = 'Select output folder'
    $fbd.SelectedPath = $txtOutput.Text
    if ($fbd.ShowDialog() -eq 'OK') {
        $txtOutput.Text = $fbd.SelectedPath
    }
})

$btnClear.Add_Click({
    $dgv.Rows.Clear()
    $lblStatus.Text       = 'Ready'
    $lblStatus.ForeColor  = [System.Drawing.Color]::Gray
    $progressBar.Value    = 0
    $btnExport.Enabled    = $false
    $txtFilter.Text       = ''
    $statusLabel.Text     = '  No results'
    $statusRowCount.Text  = 'Rows: 0'
    $script:searchResults = $null
})

$txtFilter.Add_TextChanged({
    $filterText = $txtFilter.Text.Trim()
    $dgv.SuspendLayout()
    foreach ($row in $dgv.Rows) {
        if ([string]::IsNullOrEmpty($filterText)) {
            $row.Visible = $true
        } else {
            $match = $false
            foreach ($cell in $row.Cells) {
                if ($cell.Value -and $cell.Value.ToString().IndexOf($filterText, [System.StringComparison]::OrdinalIgnoreCase) -ge 0) {
                    $match = $true; break
                }
            }
            $row.Visible = $match
        }
    }
    $dgv.ResumeLayout()
    Update-RowCount
})

# ── Download Logs ──
$btnDownload.Add_Click({

    $stations = @(Get-StationList)
    if ($stations.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show('Please enter at least one station name.', 'Missing Input', 'OK', 'Warning')
        return
    }
    if ([string]::IsNullOrWhiteSpace($txtOutput.Text) -or -not (Test-Path -LiteralPath $txtOutput.Text)) {
        [System.Windows.Forms.MessageBox]::Show('Please select a valid output folder.', 'Missing Input', 'OK', 'Warning')
        return
    }

    $dateFrom = $dtpDateFrom.Value.Date
    $dateTo   = $dtpDateTo.Value.Date
    if ($dateTo -lt $dateFrom) {
        [System.Windows.Forms.MessageBox]::Show('"Date To" cannot be earlier than "Date From".', 'Invalid Range', 'OK', 'Warning')
        return
    }

    $daySpan = ($dateTo - $dateFrom).Days + 1
    if ($daySpan -gt 31) {
        $confirm = [System.Windows.Forms.MessageBox]::Show(
            "You selected $daySpan days. Downloading may be slow and use a lot of disk space.`nContinue?",
            'Large Range', 'YesNo', 'Question')
        if ($confirm -ne 'Yes') { return }
    }

    $dateTag    = if ($dateFrom -eq $dateTo) { $dateFrom.ToString('yyyyMMdd') }
                  else { "$($dateFrom.ToString('yyyyMMdd'))-$($dateTo.ToString('yyyyMMdd'))" }
    $timestamp  = (Get-Date).ToString('yyyyMMdd_HHmmss')
    $rootFolder = Join-Path $txtOutput.Text ("MedconLogs_{0}_{1}" -f $dateTag, $timestamp)
    [void](New-Item -ItemType Directory -Path $rootFolder -Force)

    $btnSearch.Enabled    = $false
    $btnDownload.Enabled  = $false
    $progressBar.Value    = 0
    $progressBar.Maximum  = $stations.Count
    $lblStatus.ForeColor  = $accentColor

    $problems   = [System.Collections.Generic.List[string]]::new()
    $totalFiles = 0
    $sw = [System.Diagnostics.Stopwatch]::StartNew()

    Write-Activity "===========================================" -Level Step
    Write-Activity "DOWNLOAD STARTED" -Level Step
    Write-Activity "Date range: $($dateFrom.ToString('yyyy-MM-dd')) to $($dateTo.ToString('yyyy-MM-dd')) ($daySpan day(s))" -Level Info
    Write-Activity "Destination: $rootFolder" -Level Info
    Write-Activity "Stations: $($stations.Count)" -Level Info
    Write-Activity "===========================================" -Level Step

    for ($i = 0; $i -lt $stations.Count; $i++) {
        $station = $stations[$i]
        $progressBar.Value = $i + 1
        $lblStatus.Text    = "Downloading $station ($($i+1) of $($stations.Count))..."
        [System.Windows.Forms.Application]::DoEvents()

        Write-Activity "" -Level Info
        Write-Activity "[$($i+1)/$($stations.Count)] Station: $station" -Level Step

        $stationFolder = Join-Path $rootFolder $station
        $uncPath       = "\\$station\d`$\TCS_Logs\LogEx"
        $archivePath   = "\\$station\d`$\TCS_Logs\LogEx\LogArchive\Medcon.Framework.Shell"

        $copied    = Copy-PlainLogs    -UncPath $uncPath -Station $station `
                        -DateFrom $dateFrom -DateTo $dateTo `
                        -DestFolder $stationFolder -Problems $problems

        $extracted = Extract-ArchiveLogs -ArchivePath $archivePath -Station $station `
                        -DateFrom $dateFrom -DateTo $dateTo `
                        -DestFolder $stationFolder -Problems $problems

        $stationTotal = $copied + $extracted
        $totalFiles  += $stationTotal

        if ($stationTotal -eq 0) {
            Write-Activity "No files for $station in the date range" -Level Warning
            if (Test-Path -LiteralPath $stationFolder) {
                $hasContent = @(Get-ChildItem -LiteralPath $stationFolder -Recurse -ErrorAction SilentlyContinue).Count -gt 0
                if (-not $hasContent) { Remove-Item -LiteralPath $stationFolder -Recurse -Force -ErrorAction SilentlyContinue }
            }
        } else {
            Write-Activity "Station done: $copied from LogEx, $extracted from Archive" -Level Success
        }
    }

    $sw.Stop()
    $elapsed = $sw.Elapsed.ToString('mm\:ss\.f')

    $lblStatus.Text      = "Download done in $elapsed  -  $totalFiles file(s), $($problems.Count) problem(s)"
    $lblStatus.ForeColor = if ($totalFiles -gt 0) { $successColor } else { [System.Drawing.Color]::OrangeRed }
    $statusLabel.Text    = "  Downloaded $totalFiles file(s) to $rootFolder"

    Write-Activity "" -Level Info
    Write-Activity "===========================================" -Level Step
    Write-Activity "DOWNLOAD COMPLETE - $totalFiles file(s) in $elapsed" -Level Success
    Write-Activity "Location: $rootFolder" -Level Info
    Write-Activity "===========================================" -Level Step

    $btnSearch.Enabled   = $true
    $btnDownload.Enabled = $true

    if ($totalFiles -gt 0) {
        $open = [System.Windows.Forms.MessageBox]::Show(
            "Downloaded $totalFiles file(s).`nOpen the folder now?",
            'Download Complete', 'YesNo', 'Question')
        if ($open -eq 'Yes') { Invoke-Item $rootFolder }
    } elseif ($problems.Count -gt 0) {
        $msg = "No files were downloaded.`n`n" + ($problems -join "`n")
        [System.Windows.Forms.MessageBox]::Show($msg, 'Problems', 'OK', 'Warning')
    } else {
        [System.Windows.Forms.MessageBox]::Show('No matching files found in the selected date range.', 'No Files', 'OK', 'Information')
    }
})

# ── Main Search ──
$btnSearch.Add_Click({

    $stations = @(Get-StationList)

    if ($stations.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show('Please enter at least one station name.', 'Missing Input', 'OK', 'Warning')
        return
    }
    if ([string]::IsNullOrWhiteSpace($txtSearch.Text)) {
        [System.Windows.Forms.MessageBox]::Show('Please enter a search string.', 'Missing Input', 'OK', 'Warning')
        return
    }
    if ([string]::IsNullOrWhiteSpace($txtOutput.Text) -or -not (Test-Path -LiteralPath $txtOutput.Text)) {
        [System.Windows.Forms.MessageBox]::Show('Please select a valid output folder.', 'Missing Input', 'OK', 'Warning')
        return
    }

    $dateFrom = $dtpDateFrom.Value.Date
    $dateTo   = $dtpDateTo.Value.Date

    if ($dateTo -lt $dateFrom) {
        [System.Windows.Forms.MessageBox]::Show('"Date To" cannot be earlier than "Date From".', 'Invalid Range', 'OK', 'Warning')
        return
    }

    $daySpan = ($dateTo - $dateFrom).Days + 1
    if ($daySpan -gt 31) {
        $confirm = [System.Windows.Forms.MessageBox]::Show(
            "You selected a range of $daySpan days. This may take a while.`nContinue?",
            'Large Range', 'YesNo', 'Question')
        if ($confirm -ne 'Yes') { return }
    }

    $dgv.Rows.Clear()
    $txtFilter.Text      = ''
    $btnSearch.Enabled   = $false
    $btnDownload.Enabled = $false
    $btnExport.Enabled   = $false
    $progressBar.Value   = 0
    $lblStatus.ForeColor = $accentColor

    $searchString = $txtSearch.Text
    $outputFolder = $txtOutput.Text

    $results  = [System.Collections.Generic.List[object]]::new()
    $problems = [System.Collections.Generic.List[string]]::new()

    $progressBar.Maximum = $stations.Count
    $archiveHits = 0
    $sw = [System.Diagnostics.Stopwatch]::StartNew()

    Write-Activity "===========================================" -Level Step
    Write-Activity "SEARCH STARTED" -Level Step
    Write-Activity "Date range: $($dateFrom.ToString('yyyy-MM-dd')) to $($dateTo.ToString('yyyy-MM-dd')) ($daySpan day(s))" -Level Info
    Write-Activity "Pattern: '$searchString'" -Level Info
    Write-Activity "Stations to scan: $($stations.Count)" -Level Info
    Write-Activity "===========================================" -Level Step

    for ($i = 0; $i -lt $stations.Count; $i++) {
        $station = $stations[$i]
        $progressBar.Value = $i + 1
        $lblStatus.Text = "Scanning $station ($($i+1) of $($stations.Count))..."
        $statusLabel.Text = "  Searching... $($results.Count) match(es) so far"
        [System.Windows.Forms.Application]::DoEvents()

        Write-Activity "" -Level Info
        Write-Activity "[$($i+1)/$($stations.Count)] Station: $station" -Level Step

        $stationSw = [System.Diagnostics.Stopwatch]::StartNew()
        $matchesBefore = $results.Count

        $uncPath     = "\\$station\d`$\TCS_Logs\LogEx"
        $archivePath = "\\$station\d`$\TCS_Logs\LogEx\LogArchive\Medcon.Framework.Shell"

        $foundInMain      = $false
        $foundInArchive   = $false
        $stationReachable = $false

        # 1) Search main LogEx folder
        if (Test-Path -LiteralPath $uncPath) {
            $stationReachable = $true
            $foundInMain = Search-PlainFiles -UncPath $uncPath -Station $station `
                -DateFrom $dateFrom -DateTo $dateTo -SearchStr $searchString `
                -Results $results -Problems $problems
        } else {
            Write-Activity "Main folder not reachable: $uncPath" -Level Error
        }

        # 2) ALWAYS also search archive zips (date ranges may span both)
        $lblStatus.Text = "Scanning $station archive ($($i+1) of $($stations.Count))..."
        [System.Windows.Forms.Application]::DoEvents()

        Write-Activity "Now searching archive (in case range spans archived dates)..." -Level Info

        $foundInArchive = Search-ArchiveZips -ArchivePath $archivePath -Station $station `
            -DateFrom $dateFrom -DateTo $dateTo -SearchStr $searchString `
            -Results $results -Problems $problems

        if ($foundInArchive) {
            $archiveHits++
            $stationReachable = $true
        }

        # 3) Report problems
        if (-not $stationReachable) {
            $problems.Add("$station  -  path not accessible (neither LogEx nor LogArchive)")
        }
        elseif (-not $foundInMain -and -not $foundInArchive) {
            $rangeStr = if ($dateFrom -eq $dateTo) { $dateFrom.ToString('yyyy-MM-dd') }
                        else { "$($dateFrom.ToString('yyyy-MM-dd')) to $($dateTo.ToString('yyyy-MM-dd'))" }
            $problems.Add("$station  -  no matching files for $rangeStr (checked LogEx + Archive)")
            Write-Activity "No matches found for $station in date range" -Level Warning
        }

        $stationSw.Stop()
        $stationMatches = $results.Count - $matchesBefore
        Write-Activity "Station done in $($stationSw.Elapsed.ToString('mm\:ss\.f')) - $stationMatches new match(es)" -Level Success

        [System.Windows.Forms.Application]::DoEvents()
    }

    $sw.Stop()
    $elapsed = $sw.Elapsed.ToString('mm\:ss\.f')

    $script:searchResults = $results

    $dateTag = if ($dateFrom -eq $dateTo) { $dateFrom.ToString('yyyyMMdd') }
               else { "$($dateFrom.ToString('yyyyMMdd'))-$($dateTo.ToString('yyyyMMdd'))" }
    $safeString = ($searchString -replace '[^a-zA-Z0-9_\-]+', '_').Trim('_')
    if (-not $safeString) { $safeString = 'search' }
    $script:csvPath = Join-Path $outputFolder ("LogSearch_{0}_{1}.csv" -f $dateTag, $safeString)

    $archiveNote = if ($archiveHits -gt 0) { "  ($archiveHits from archive)" } else { '' }
    $lblStatus.Text      = "Done in $elapsed  -  $($results.Count) match(es), $($problems.Count) problem(s)$archiveNote"
    $lblStatus.ForeColor = if ($results.Count -gt 0) { $successColor } else { [System.Drawing.Color]::OrangeRed }
    $statusLabel.Text    = "  Search complete: $($results.Count) match(es) across $($stations.Count) station(s)$archiveNote"

    Write-Activity "" -Level Info
    Write-Activity "===========================================" -Level Step
    Write-Activity "SEARCH COMPLETE - $($results.Count) total match(es) in $elapsed" -Level Success
    if ($problems.Count -gt 0) {
        Write-Activity "$($problems.Count) problem(s) encountered" -Level Warning
    }
    Write-Activity "===========================================" -Level Step

    $btnSearch.Enabled   = $true
    $btnDownload.Enabled = $true
    $btnExport.Enabled   = ($results.Count -gt 0)
    Update-RowCount

    foreach ($row in $dgv.Rows) {
        if ($row.Cells['Source'].Value -eq 'Archive') {
            $row.DefaultCellStyle.ForeColor = $archiveColor
        }
    }

    if ($problems.Count -gt 0) {
        $msg = "Search complete with $($problems.Count) problem(s):`n`n" + ($problems -join "`n")
        [System.Windows.Forms.MessageBox]::Show($msg, 'Problems Encountered', 'OK', 'Warning')
    }
})

# ── Export CSV ──
$btnExport.Add_Click({
    if (-not $script:searchResults -or $script:searchResults.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show('No results to export.', 'Export', 'OK', 'Information')
        return
    }

    $sfd = New-Object System.Windows.Forms.SaveFileDialog
    $sfd.Title            = 'Save Results As'
    $sfd.Filter           = 'CSV files (*.csv)|*.csv'
    $sfd.FileName         = [System.IO.Path]::GetFileName($script:csvPath)
    $sfd.InitialDirectory = [System.IO.Path]::GetDirectoryName($script:csvPath)

    if ($sfd.ShowDialog() -eq 'OK') {
        $script:searchResults | Export-Csv -LiteralPath $sfd.FileName -NoTypeInformation -Encoding UTF8
        $statusLabel.Text = "  CSV saved: $($sfd.FileName)"

        $openIt = [System.Windows.Forms.MessageBox]::Show("CSV saved.`nOpen it now?", 'Exported', 'YesNo', 'Question')
        if ($openIt -eq 'Yes') { Invoke-Item $sfd.FileName }
    }
})

$txtSearch.Add_KeyDown({
    if ($_.KeyCode -eq 'Return') { $btnSearch.PerformClick() }
})

# ══════════════════════════════════════════════
#  LAUNCH
# ══════════════════════════════════════════════
[void]$form.ShowDialog()
