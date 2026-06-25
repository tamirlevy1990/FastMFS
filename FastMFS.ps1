# Search-MedconLogs-GUI.ps1
# v2 — Redesigned UI: groupboxes, date presets, toolbar, splitter, responsive layout
# Run:  powershell -ExecutionPolicy Bypass -File ".\Search-MedconLogs-GUI.ps1"

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.IO.Compression

[System.Windows.Forms.Application]::EnableVisualStyles()

# ══════════════════════════════════════════════
#  THEME
# ══════════════════════════════════════════════
$clrAccent     = [System.Drawing.Color]::FromArgb(0, 120, 215)
$clrAccentDark = [System.Drawing.Color]::FromArgb(0, 90, 170)
$clrSuccess    = [System.Drawing.Color]::FromArgb(16, 150, 72)
$clrDanger     = [System.Drawing.Color]::FromArgb(200, 50, 50)
$clrArchive    = [System.Drawing.Color]::FromArgb(180, 120, 0)
$clrWarning    = [System.Drawing.Color]::FromArgb(220, 140, 0)
$clrBg         = [System.Drawing.Color]::FromArgb(245, 246, 250)
$clrPanel      = [System.Drawing.Color]::White
$clrPanelAlt   = [System.Drawing.Color]::FromArgb(250, 251, 253)
$clrGridAlt    = [System.Drawing.Color]::FromArgb(245, 248, 252)
$clrHeaderBg   = [System.Drawing.Color]::FromArgb(40, 44, 52)
$clrHeaderFg   = [System.Drawing.Color]::White
$clrInput      = [System.Drawing.Color]::FromArgb(252, 252, 255)
$clrLabel      = [System.Drawing.Color]::FromArgb(60, 60, 60)
$clrMuted      = [System.Drawing.Color]::FromArgb(120, 120, 120)
$clrBorder     = [System.Drawing.Color]::FromArgb(220, 223, 230)
$clrGray       = [System.Drawing.Color]::FromArgb(140, 140, 140)

$fontN         = New-Object System.Drawing.Font('Segoe UI', 9.5)
$fontB         = New-Object System.Drawing.Font('Segoe UI', 9.5, [System.Drawing.FontStyle]::Bold)
$fontS         = New-Object System.Drawing.Font('Segoe UI', 8.5)
$fontT         = New-Object System.Drawing.Font('Segoe UI Semibold', 12)
$fontH         = New-Object System.Drawing.Font('Segoe UI Semibold', 10)
$fontM         = New-Object System.Drawing.Font('Consolas', 9)
$fontBig       = New-Object System.Drawing.Font('Segoe UI Semibold', 10)

# Button factory ---------------------------------------------------
function New-FlatButton {
    param(
        [string]$Text,
        [System.Drawing.Color]$Back,
        [System.Drawing.Color]$Fore = [System.Drawing.Color]::White,
        [System.Drawing.Font]$Font = $fontN,
        [int]$Width = 100,
        [int]$Height = 30
    )
    $b = New-Object System.Windows.Forms.Button
    $b.Text      = $Text
    $b.Font      = $Font
    $b.ForeColor = $Fore
    $b.BackColor = $Back
    $b.FlatStyle = 'Flat'
    $b.FlatAppearance.BorderSize = 0
    $b.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(
        [Math]::Max(0, $Back.R - 20),
        [Math]::Max(0, $Back.G - 20),
        [Math]::Max(0, $Back.B - 20))
    $b.Cursor    = [System.Windows.Forms.Cursors]::Hand
    $b.Size      = New-Object System.Drawing.Size($Width, $Height)
    $b.TextAlign = 'MiddleCenter'
    return $b
}

# ══════════════════════════════════════════════
#  FORM
# ══════════════════════════════════════════════
$form = New-Object System.Windows.Forms.Form
$form.Text          = 'Medcon Log Search'
$form.Size          = New-Object System.Drawing.Size(1180, 980)
$form.StartPosition = 'CenterScreen'
$form.MinimumSize   = New-Object System.Drawing.Size(1000, 800)
$form.Font          = $fontN
$form.BackColor     = $clrBg
try { $form.Icon = [System.Drawing.SystemIcons]::Application } catch {}

# ══════════════════════════════════════════════
#  HEADER + TOOLBAR
# ══════════════════════════════════════════════
$panelHeader = New-Object System.Windows.Forms.Panel
$panelHeader.Dock      = 'Top'
$panelHeader.Height    = 60
$panelHeader.BackColor = $clrHeaderBg
$form.Controls.Add($panelHeader)

$lblTitle = New-Object System.Windows.Forms.Label
$lblTitle.Text      = '  Medcon Framework Shell  •  Log Search'
$lblTitle.Font      = $fontT
$lblTitle.ForeColor = [System.Drawing.Color]::White
$lblTitle.Location  = New-Object System.Drawing.Point(15, 18)
$lblTitle.AutoSize  = $true
$panelHeader.Controls.Add($lblTitle)

# Status pill (right side of header)
$pillStatus = New-Object System.Windows.Forms.Label
$pillStatus.Text       = '  ● Ready  '
$pillStatus.Font       = $fontB
$pillStatus.ForeColor  = [System.Drawing.Color]::White
$pillStatus.BackColor  = [System.Drawing.Color]::FromArgb(70, 75, 85)
$pillStatus.AutoSize   = $false
$pillStatus.Size       = New-Object System.Drawing.Size(220, 28)
$pillStatus.TextAlign  = 'MiddleLeft'
$pillStatus.Anchor     = 'Top,Right'
$pillStatus.Location   = New-Object System.Drawing.Point(($form.ClientSize.Width - 240), 16)
$panelHeader.Controls.Add($pillStatus)

function Set-StatusPill {
    param([string]$Text, [string]$State = 'idle')
    $map = @{
        'idle'    = @{ Bg = [System.Drawing.Color]::FromArgb(70, 75, 85);  Dot = '●' }
        'working' = @{ Bg = $clrAccent;  Dot = '●' }
        'ok'      = @{ Bg = $clrSuccess; Dot = '●' }
        'warn'    = @{ Bg = $clrWarning; Dot = '●' }
        'err'     = @{ Bg = $clrDanger;  Dot = '●' }
    }
    $pillStatus.BackColor = $map[$State].Bg
    $pillStatus.Text      = "  $($map[$State].Dot) $Text  "
}

# Toolbar ---------------------------------------------------------
$panelToolbar = New-Object System.Windows.Forms.Panel
$panelToolbar.Dock      = 'Top'
$panelToolbar.Height    = 56
$panelToolbar.BackColor = $clrPanel
$form.Controls.Add($panelToolbar)

$tbBorder = New-Object System.Windows.Forms.Panel
$tbBorder.Dock      = 'Bottom'
$tbBorder.Height    = 1
$tbBorder.BackColor = $clrBorder
$panelToolbar.Controls.Add($tbBorder)

$btnSearch   = New-FlatButton '🔍  Search'         $clrSuccess  ([System.Drawing.Color]::White) $fontBig 130 38
$btnDownload = New-FlatButton '⬇  Download Logs'   $clrArchive  ([System.Drawing.Color]::White) $fontBig 160 38
$btnExport   = New-FlatButton '📄  Export CSV'     $clrAccent   ([System.Drawing.Color]::White) $fontBig 130 38
$btnClear    = New-FlatButton '🗙  Clear'          $clrGray     ([System.Drawing.Color]::White) $fontBig 100 38
$btnExport.Enabled = $false

$btnSearch.Location   = New-Object System.Drawing.Point(15,  9)
$btnDownload.Location = New-Object System.Drawing.Point(150, 9)
$btnExport.Location   = New-Object System.Drawing.Point(315, 9)
$btnClear.Location    = New-Object System.Drawing.Point(450, 9)

$panelToolbar.Controls.AddRange(@($btnSearch, $btnDownload, $btnExport, $btnClear))

# ══════════════════════════════════════════════
#  INPUT AREA  (3-column TableLayoutPanel)
# ══════════════════════════════════════════════
$panelInput = New-Object System.Windows.Forms.TableLayoutPanel
$panelInput.Dock        = 'Top'
$panelInput.Height      = 235
$panelInput.BackColor   = $clrBg
$panelInput.Padding     = New-Object System.Windows.Forms.Padding(10, 8, 10, 8)
$panelInput.ColumnCount = 3
$panelInput.RowCount    = 1
[void]$panelInput.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('Percent', 36)))
[void]$panelInput.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('Percent', 32)))
[void]$panelInput.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('Percent', 32)))
$form.Controls.Add($panelInput)

# --- GROUPBOX 1: Date Range -------------------------------------
$gbDate = New-Object System.Windows.Forms.GroupBox
$gbDate.Text      = '  Date Range  '
$gbDate.Font      = $fontH
$gbDate.ForeColor = $clrLabel
$gbDate.Dock      = 'Fill'
$gbDate.BackColor = $clrPanel
$gbDate.Margin    = New-Object System.Windows.Forms.Padding(0, 0, 6, 0)

$lblFrom = New-Object System.Windows.Forms.Label
$lblFrom.Text      = 'From'
$lblFrom.Font      = $fontN
$lblFrom.ForeColor = $clrLabel
$lblFrom.Location  = New-Object System.Drawing.Point(15, 30)
$lblFrom.AutoSize  = $true
$gbDate.Controls.Add($lblFrom)

$dtpDateFrom = New-Object System.Windows.Forms.DateTimePicker
$dtpDateFrom.Font     = $fontN
$dtpDateFrom.Location = New-Object System.Drawing.Point(60, 27)
$dtpDateFrom.Size     = New-Object System.Drawing.Size(140, 26)
$dtpDateFrom.Format   = 'Short'
$dtpDateFrom.Value    = (Get-Date).Date
$gbDate.Controls.Add($dtpDateFrom)

$lblTo = New-Object System.Windows.Forms.Label
$lblTo.Text      = 'To'
$lblTo.Font      = $fontN
$lblTo.ForeColor = $clrLabel
$lblTo.Location  = New-Object System.Drawing.Point(15, 65)
$lblTo.AutoSize  = $true
$gbDate.Controls.Add($lblTo)

$dtpDateTo = New-Object System.Windows.Forms.DateTimePicker
$dtpDateTo.Font     = $fontN
$dtpDateTo.Location = New-Object System.Drawing.Point(60, 62)
$dtpDateTo.Size     = New-Object System.Drawing.Size(140, 26)
$dtpDateTo.Format   = 'Short'
$dtpDateTo.Value    = (Get-Date).Date
$gbDate.Controls.Add($dtpDateTo)

$lblDayCount = New-Object System.Windows.Forms.Label
$lblDayCount.Text      = '(1 day)'
$lblDayCount.Font      = $fontS
$lblDayCount.ForeColor = $clrMuted
$lblDayCount.Location  = New-Object System.Drawing.Point(210, 66)
$lblDayCount.AutoSize  = $true
$gbDate.Controls.Add($lblDayCount)

# Quick presets
$lblPresets = New-Object System.Windows.Forms.Label
$lblPresets.Text      = 'Quick presets:'
$lblPresets.Font      = $fontS
$lblPresets.ForeColor = $clrMuted
$lblPresets.Location  = New-Object System.Drawing.Point(15, 100)
$lblPresets.AutoSize  = $true
$gbDate.Controls.Add($lblPresets)

$btnPresetToday = New-FlatButton 'Today'     $clrPanelAlt $clrLabel $fontS 70 26
$btnPresetYday  = New-FlatButton 'Yesterday' $clrPanelAlt $clrLabel $fontS 80 26
$btnPreset7d    = New-FlatButton '7 days'    $clrPanelAlt $clrLabel $fontS 70 26
$btnPreset30d   = New-FlatButton '30 days'   $clrPanelAlt $clrLabel $fontS 70 26
$btnPresetToday.FlatAppearance.BorderSize = 1
$btnPresetYday.FlatAppearance.BorderSize  = 1
$btnPreset7d.FlatAppearance.BorderSize    = 1
$btnPreset30d.FlatAppearance.BorderSize   = 1
$btnPresetToday.FlatAppearance.BorderColor = $clrBorder
$btnPresetYday.FlatAppearance.BorderColor  = $clrBorder
$btnPreset7d.FlatAppearance.BorderColor    = $clrBorder
$btnPreset30d.FlatAppearance.BorderColor   = $clrBorder
$btnPresetToday.Location = New-Object System.Drawing.Point(15,  125)
$btnPresetYday.Location  = New-Object System.Drawing.Point(90,  125)
$btnPreset7d.Location    = New-Object System.Drawing.Point(175, 125)
$btnPreset30d.Location   = New-Object System.Drawing.Point(250, 125)
$gbDate.Controls.AddRange(@($btnPresetToday, $btnPresetYday, $btnPreset7d, $btnPreset30d))

$panelInput.Controls.Add($gbDate, 0, 0)

# --- GROUPBOX 2: Search & Output -------------------------------
$gbSearch = New-Object System.Windows.Forms.GroupBox
$gbSearch.Text      = '  Search & Output  '
$gbSearch.Font      = $fontH
$gbSearch.ForeColor = $clrLabel
$gbSearch.Dock      = 'Fill'
$gbSearch.BackColor = $clrPanel
$gbSearch.Margin    = New-Object System.Windows.Forms.Padding(3, 0, 3, 0)

$lblSearch = New-Object System.Windows.Forms.Label
$lblSearch.Text      = 'Search string'
$lblSearch.Font      = $fontN
$lblSearch.ForeColor = $clrLabel
$lblSearch.Location  = New-Object System.Drawing.Point(15, 30)
$lblSearch.AutoSize  = $true
$gbSearch.Controls.Add($lblSearch)

$txtSearch = New-Object System.Windows.Forms.TextBox
$txtSearch.Font      = $fontN
$txtSearch.Location  = New-Object System.Drawing.Point(15, 50)
$txtSearch.Size      = New-Object System.Drawing.Size(310, 26)
$txtSearch.BackColor = $clrInput
$txtSearch.Anchor    = 'Top,Left,Right'
$gbSearch.Controls.Add($txtSearch)

$lblOutput = New-Object System.Windows.Forms.Label
$lblOutput.Text      = 'Output folder'
$lblOutput.Font      = $fontN
$lblOutput.ForeColor = $clrLabel
$lblOutput.Location  = New-Object System.Drawing.Point(15, 90)
$lblOutput.AutoSize  = $true
$gbSearch.Controls.Add($lblOutput)

$txtOutput = New-Object System.Windows.Forms.TextBox
$txtOutput.Font      = $fontN
$txtOutput.Location  = New-Object System.Drawing.Point(15, 110)
$txtOutput.Size      = New-Object System.Drawing.Size(240, 26)
$txtOutput.BackColor = $clrInput
$txtOutput.Text      = [System.Environment]::GetFolderPath('Desktop')
$txtOutput.Anchor    = 'Top,Left,Right'
$gbSearch.Controls.Add($txtOutput)

$btnBrowseOutput = New-FlatButton 'Browse' $clrGray ([System.Drawing.Color]::White) $fontS 60 26
$btnBrowseOutput.Location = New-Object System.Drawing.Point(265, 110)
$btnBrowseOutput.Anchor   = 'Top,Right'
$gbSearch.Controls.Add($btnBrowseOutput)

$lblHint = New-Object System.Windows.Forms.Label
$lblHint.Text      = "Search scans LogEx + LogArchive.`r`nDownload copies the full log files locally."
$lblHint.Font      = $fontS
$lblHint.ForeColor = $clrMuted
$lblHint.Location  = New-Object System.Drawing.Point(15, 148)
$lblHint.Size      = New-Object System.Drawing.Size(310, 40)
$lblHint.Anchor    = 'Top,Left,Right'
$gbSearch.Controls.Add($lblHint)

$panelInput.Controls.Add($gbSearch, 1, 0)

# --- GROUPBOX 3: Stations --------------------------------------
$gbStations = New-Object System.Windows.Forms.GroupBox
$gbStations.Text      = '  Stations  '
$gbStations.Font      = $fontH
$gbStations.ForeColor = $clrLabel
$gbStations.Dock      = 'Fill'
$gbStations.BackColor = $clrPanel
$gbStations.Margin    = New-Object System.Windows.Forms.Padding(6, 0, 0, 0)

$lblStationHint = New-Object System.Windows.Forms.Label
$lblStationHint.Text      = 'One hostname per line (commas, spaces, and # comments also accepted)'
$lblStationHint.Font      = $fontS
$lblStationHint.ForeColor = $clrMuted
$lblStationHint.Location  = New-Object System.Drawing.Point(12, 28)
$lblStationHint.AutoSize  = $true
$gbStations.Controls.Add($lblStationHint)

$txtStations = New-Object System.Windows.Forms.TextBox
$txtStations.Font           = $fontM
$txtStations.Location       = New-Object System.Drawing.Point(12, 50)
$txtStations.Size           = New-Object System.Drawing.Size(330, 110)
$txtStations.Anchor         = 'Top,Left,Right,Bottom'
$txtStations.Multiline      = $true
$txtStations.ScrollBars     = 'Vertical'
$txtStations.BackColor      = $clrInput
$txtStations.WordWrap       = $false
$txtStations.AcceptsReturn  = $true
$gbStations.Controls.Add($txtStations)

$lblStationCount = New-Object System.Windows.Forms.Label
$lblStationCount.Text      = '0 stations'
$lblStationCount.Font      = $fontB
$lblStationCount.ForeColor = $clrMuted
$lblStationCount.Location  = New-Object System.Drawing.Point(12, 168)
$lblStationCount.AutoSize  = $true
$lblStationCount.Anchor    = 'Bottom,Left'
$gbStations.Controls.Add($lblStationCount)

$btnLoadFile      = New-FlatButton 'Load File' $clrAccent ([System.Drawing.Color]::White) $fontS 80 24
$btnClearStations = New-FlatButton 'Clear'     $clrGray   ([System.Drawing.Color]::White) $fontS 55 24
$btnSortStations  = New-FlatButton 'Sort A-Z'  $clrGray   ([System.Drawing.Color]::White) $fontS 65 24
$btnLoadFile.Location      = New-Object System.Drawing.Point(140, 167)
$btnClearStations.Location = New-Object System.Drawing.Point(225, 167)
$btnSortStations.Location  = New-Object System.Drawing.Point(285, 167)
$btnLoadFile.Anchor        = 'Bottom,Right'
$btnClearStations.Anchor   = 'Bottom,Right'
$btnSortStations.Anchor    = 'Bottom,Right'
$gbStations.Controls.AddRange(@($btnLoadFile, $btnClearStations, $btnSortStations))

$panelInput.Controls.Add($gbStations, 2, 0)

# ══════════════════════════════════════════════
#  PROGRESS + FILTER STRIP
# ══════════════════════════════════════════════
$panelMid = New-Object System.Windows.Forms.Panel
$panelMid.Dock      = 'Top'
$panelMid.Height    = 40
$panelMid.BackColor = $clrPanel
$form.Controls.Add($panelMid)

$midBorder = New-Object System.Windows.Forms.Panel
$midBorder.Dock      = 'Bottom'
$midBorder.Height    = 1
$midBorder.BackColor = $clrBorder
$panelMid.Controls.Add($midBorder)

$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(15, 10)
$progressBar.Size     = New-Object System.Drawing.Size(380, 20)
$progressBar.Style    = 'Continuous'
$panelMid.Controls.Add($progressBar)

$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Text      = 'Ready'
$lblStatus.Font      = $fontS
$lblStatus.Location  = New-Object System.Drawing.Point(405, 13)
$lblStatus.AutoSize  = $true
$lblStatus.ForeColor = $clrMuted
$panelMid.Controls.Add($lblStatus)

$lblFilter = New-Object System.Windows.Forms.Label
$lblFilter.Text      = 'Filter results:'
$lblFilter.Font      = $fontS
$lblFilter.ForeColor = $clrLabel
$lblFilter.Location  = New-Object System.Drawing.Point(($form.ClientSize.Width - 320), 13)
$lblFilter.AutoSize  = $true
$lblFilter.Anchor    = 'Top,Right'
$panelMid.Controls.Add($lblFilter)

$txtFilter = New-Object System.Windows.Forms.TextBox
$txtFilter.Font      = $fontS
$txtFilter.Location  = New-Object System.Drawing.Point(($form.ClientSize.Width - 230), 10)
$txtFilter.Size      = New-Object System.Drawing.Size(215, 22)
$txtFilter.Anchor    = 'Top,Right'
$txtFilter.BackColor = $clrInput
$panelMid.Controls.Add($txtFilter)

# ══════════════════════════════════════════════
#  STATUS STRIP (bottom)
# ══════════════════════════════════════════════
$statusStrip = New-Object System.Windows.Forms.StatusStrip
$statusStrip.BackColor = $clrHeaderBg
$statusStrip.SizingGrip = $false

$statusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
$statusLabel.Text      = '  No results'
$statusLabel.ForeColor = [System.Drawing.Color]::FromArgb(200, 200, 200)
$statusLabel.Font      = $fontS
$statusLabel.Spring    = $true
$statusLabel.TextAlign = 'MiddleLeft'
[void]$statusStrip.Items.Add($statusLabel)

$statusRowCount = New-Object System.Windows.Forms.ToolStripStatusLabel
$statusRowCount.Text      = 'Rows: 0'
$statusRowCount.ForeColor = [System.Drawing.Color]::FromArgb(200, 200, 200)
$statusRowCount.Font      = $fontS
$statusRowCount.Alignment = 'Right'
[void]$statusStrip.Items.Add($statusRowCount)

$form.Controls.Add($statusStrip)

# ══════════════════════════════════════════════
#  SPLIT: Results (top) | Activity Log (bottom)
# ══════════════════════════════════════════════
$splitMain = New-Object System.Windows.Forms.SplitContainer
$splitMain.Dock        = 'Fill'
$splitMain.Orientation = 'Horizontal'
$splitMain.SplitterWidth = 6
$splitMain.BackColor   = $clrBorder
$splitMain.Panel1.BackColor = $clrPanel
$splitMain.Panel2.BackColor = $clrPanel
$form.Controls.Add($splitMain)
# After form sizes settle, set splitter (deferred)
$form.Add_Shown({
    try {
        $available = $splitMain.Height - $splitMain.SplitterWidth
        $splitMain.SplitterDistance = [int]($available * 0.62)
    } catch {}
})

# --- Results grid in Panel1 ---
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
$dgv.DefaultCellStyle.Font        = $fontM
$dgv.DefaultCellStyle.Padding     = New-Object System.Windows.Forms.Padding(4, 3, 4, 3)
$dgv.DefaultCellStyle.SelectionBackColor = [System.Drawing.Color]::FromArgb(210, 225, 245)
$dgv.DefaultCellStyle.SelectionForeColor = [System.Drawing.Color]::Black
$dgv.ColumnHeadersDefaultCellStyle.BackColor = $clrHeaderBg
$dgv.ColumnHeadersDefaultCellStyle.ForeColor = $clrHeaderFg
$dgv.ColumnHeadersDefaultCellStyle.Font      = $fontB
$dgv.ColumnHeadersDefaultCellStyle.Padding   = New-Object System.Windows.Forms.Padding(4, 5, 4, 5)
$dgv.ColumnHeadersHeight          = 34
$dgv.ColumnHeadersHeightSizeMode  = 'DisableResizing'
$dgv.EnableHeadersVisualStyles    = $false
$dgv.RowHeadersVisible            = $false
$dgv.RowTemplate.Height           = 26
$dgv.AlternatingRowsDefaultCellStyle.BackColor = $clrGridAlt
$splitMain.Panel1.Controls.Add($dgv)

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

# --- Activity Log in Panel2 ---
$panelLogHeader = New-Object System.Windows.Forms.Panel
$panelLogHeader.Dock      = 'Top'
$panelLogHeader.Height    = 32
$panelLogHeader.BackColor = $clrPanel
$splitMain.Panel2.Controls.Add($panelLogHeader)

$lblActivityHeader = New-Object System.Windows.Forms.Label
$lblActivityHeader.Text      = '  Activity Log'
$lblActivityHeader.Font      = $fontH
$lblActivityHeader.ForeColor = $clrLabel
$lblActivityHeader.Location  = New-Object System.Drawing.Point(8, 7)
$lblActivityHeader.AutoSize  = $true
$panelLogHeader.Controls.Add($lblActivityHeader)

$btnClearLog = New-FlatButton 'Clear'    $clrGray  ([System.Drawing.Color]::White) $fontS 60 22
$btnSaveLog  = New-FlatButton 'Save'     $clrAccent ([System.Drawing.Color]::White) $fontS 60 22
$btnClearLog.Location = New-Object System.Drawing.Point(110, 5)
$btnSaveLog.Location  = New-Object System.Drawing.Point(175, 5)
$panelLogHeader.Controls.AddRange(@($btnClearLog, $btnSaveLog))

$txtActivity = New-Object System.Windows.Forms.RichTextBox
$txtActivity.Dock        = 'Fill'
$txtActivity.Font        = $fontM
$txtActivity.BackColor   = [System.Drawing.Color]::FromArgb(28, 30, 36)
$txtActivity.ForeColor   = [System.Drawing.Color]::FromArgb(220, 220, 220)
$txtActivity.ReadOnly    = $true
$txtActivity.ScrollBars  = 'Vertical'
$txtActivity.BorderStyle = 'None'
$splitMain.Panel2.Controls.Add($txtActivity)
$txtActivity.BringToFront()
$panelLogHeader.BringToFront()

# ══════════════════════════════════════════════
#  RIGHT-CLICK MENU on grid
# ══════════════════════════════════════════════
$ctxMenu = New-Object System.Windows.Forms.ContextMenuStrip
$ctxMenu.Font = $fontN
$menuCopyLine = $ctxMenu.Items.Add('Copy Matched Line')
$menuCopyRow  = $ctxMenu.Items.Add('Copy Entire Row')
[void]$ctxMenu.Items.Add('-')
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
$toolTip.SetToolTip($dtpDateTo,    'End date - files modified on or before this date')
$toolTip.SetToolTip($txtSearch,    'Case-insensitive text search inside log lines (press Enter to search)')
$toolTip.SetToolTip($txtStations,  'Type or paste station hostnames, one per line')
$toolTip.SetToolTip($btnLoadFile,  'Load station names from a .txt file')
$toolTip.SetToolTip($txtFilter,    'Filter results across all columns')
$toolTip.SetToolTip($txtOutput,    'Folder for CSV / downloaded logs')
$toolTip.SetToolTip($btnDownload,  'Copy the full log files for the selected date range to a local folder')
$toolTip.SetToolTip($btnSearch,    'Scan logs for the search string across all selected stations')
$toolTip.SetToolTip($btnExport,    'Export current results to CSV')
$toolTip.SetToolTip($btnClear,     'Clear results and reset')

# ══════════════════════════════════════════════
#  SHARED STATE
# ══════════════════════════════════════════════
$script:searchResults = $null
$script:csvPath       = $null

# ══════════════════════════════════════════════
#  HELPERS (unchanged logic)
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
    $tokens = $txtStations.Text -split '[\s,;]+'
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
    $lblStationCount.ForeColor = if ($list.Count -gt 0) { $clrSuccess } else { $clrMuted }
}

function Update-DayCount {
    $from = $dtpDateFrom.Value.Date
    $to   = $dtpDateTo.Value.Date
    if ($to -lt $from) {
        $lblDayCount.Text      = '(invalid range)'
        $lblDayCount.ForeColor = $clrDanger
    } else {
        $days = (($to - $from).Days + 1)
        $word = if ($days -eq 1) { 'day' } else { 'days' }
        $lblDayCount.Text      = "($days $word)"
        $lblDayCount.ForeColor = $clrMuted
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
        $startStr = $matches[1]; $endStr = $matches[2]
        [datetime]$zipStart = [datetime]::MinValue
        [datetime]$zipEnd   = [datetime]::MinValue
        $formats = [string[]]@('d.M.yyyy', 'dd.MM.yyyy', 'M.d.yyyy', 'dd.M.yyyy', 'd.MM.yyyy')
        $culture = [System.Globalization.CultureInfo]::InvariantCulture
        $style   = [System.Globalization.DateTimeStyles]::None
        $okStart = [datetime]::TryParseExact($startStr, $formats, $culture, $style, [ref]$zipStart)
        $okEnd   = [datetime]::TryParseExact($endStr,   $formats, $culture, $style, [ref]$zipEnd)
        if ($okStart -and $okEnd) { return @{ Start = $zipStart.Date; End = $zipEnd.Date } }
    }
    return $null
}

function Test-DateRangesOverlap {
    param([datetime]$SearchFrom, [datetime]$SearchTo, [datetime]$ZipFrom, [datetime]$ZipTo)
    return ($ZipFrom -le $SearchTo) -and ($ZipTo -ge $SearchFrom)
}

function Search-PlainFiles {
    param([string]$UncPath,[string]$Station,[datetime]$DateFrom,[datetime]$DateTo,[string]$SearchStr,
          [System.Collections.Generic.List[object]]$Results,[System.Collections.Generic.List[string]]$Problems)
    Write-Activity "Checking main folder: $UncPath" -Level Detail
    if (-not (Test-Path -LiteralPath $UncPath)) { Write-Activity "Main folder not accessible" -Level Warning; return $false }
    $allFiles = @(Get-ChildItem -LiteralPath $UncPath -File -Filter 'Medcon.Framework.Shell*' -ErrorAction SilentlyContinue)
    Write-Activity "Found $($allFiles.Count) file(s) in main folder" -Level Detail
    $files = @($allFiles | Where-Object { $_.LastWriteTime.Date -ge $DateFrom -and $_.LastWriteTime.Date -le $DateTo })
    if ($files.Count -eq 0) { Write-Activity "No files matching date range" -Level Warning; return $false }
    Write-Activity "$($files.Count) file(s) match - scanning..." -Level Info
    $matchCount = 0
    foreach ($file in $files) {
        $sizeMB = [math]::Round($file.Length / 1MB, 2)
        Write-Activity "Reading: $($file.Name) ($sizeMB MB)" -Level Detail
        [System.Windows.Forms.Application]::DoEvents()
        $stream = $null; $reader = $null
        try {
            $stream = [System.IO.FileStream]::new($file.FullName, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
            $reader = [System.IO.StreamReader]::new($stream)
            $lineNumber = 0; $fileMatches = 0
            while (($line = $reader.ReadLine()) -ne $null) {
                $lineNumber++
                if ($line.IndexOf($SearchStr, [System.StringComparison]::OrdinalIgnoreCase) -ge 0) {
                    $fileDate = $file.LastWriteTime.ToString('yyyy-MM-dd HH:mm')
                    $trimmed  = $line.Trim()
                    $Results.Add([PSCustomObject]@{ Station=$Station; Source='LogEx'; FileName=$file.Name; FileDate=$fileDate; LineNumber=$lineNumber; MatchedLine=$trimmed })
                    [void]$dgv.Rows.Add($Station,'LogEx',$file.Name,$fileDate,$lineNumber,$trimmed)
                    $fileMatches++; $matchCount++
                }
            }
            Write-Activity "Scanned $lineNumber lines, $fileMatches match(es)" -Level Detail
        } catch {
            Write-Activity "Error reading $($file.Name): $($_.Exception.Message)" -Level Error
            $Problems.Add("$Station - $($file.Name): $($_.Exception.Message)")
        } finally {
            if ($reader) { $reader.Dispose() }; if ($stream) { $stream.Dispose() }
        }
    }
    Write-Activity "Main folder scan complete: $matchCount match(es)" -Level Success
    return $true
}

function Search-ArchiveZips {
    param([string]$ArchivePath,[string]$Station,[datetime]$DateFrom,[datetime]$DateTo,[string]$SearchStr,
          [System.Collections.Generic.List[object]]$Results,[System.Collections.Generic.List[string]]$Problems)
    Write-Activity "Checking archive folder: $ArchivePath" -Level Detail
    if (-not (Test-Path -LiteralPath $ArchivePath)) { Write-Activity "Archive folder not accessible" -Level Warning; return $false }
    $zipFiles = @(Get-ChildItem -LiteralPath $ArchivePath -File -Filter '*.zip' -ErrorAction SilentlyContinue)
    Write-Activity "Found $($zipFiles.Count) zip file(s)" -Level Detail
    if ($zipFiles.Count -eq 0) { return $false }
    $foundAny = $false; $zipsChecked = 0; $zipsSkipped = 0; $zipsUnparsable = 0
    foreach ($zipFile in $zipFiles) {
        $zipRange = Get-ZipDateRange -ZipName $zipFile.Name
        if ($null -ne $zipRange) {
            if (-not (Test-DateRangesOverlap -SearchFrom $DateFrom -SearchTo $DateTo -ZipFrom $zipRange.Start -ZipTo $zipRange.End)) { $zipsSkipped++; continue }
        } else {
            $zipsUnparsable++
            Write-Activity "Could not parse date from: $($zipFile.Name) (will search anyway)" -Level Warning
        }
        $zipsChecked++
        $sizeMB = [math]::Round($zipFile.Length / 1MB, 2)
        Write-Activity "Opening zip: $($zipFile.Name) ($sizeMB MB)" -Level Detail
        [System.Windows.Forms.Application]::DoEvents()
        $zipStream = $null; $archive = $null
        try {
            $zipStream = [System.IO.FileStream]::new($zipFile.FullName, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
            $archive = [System.IO.Compression.ZipArchive]::new($zipStream, [System.IO.Compression.ZipArchiveMode]::Read)
            $matchingEntries = @($archive.Entries | Where-Object {
                $_.Name -like 'Medcon.Framework.Shell*' -and
                $_.LastWriteTime.Date -ge $DateFrom -and $_.LastWriteTime.Date -le $DateTo
            })
            Write-Activity "Zip has $($archive.Entries.Count) entries, $($matchingEntries.Count) match" -Level Detail
            foreach ($entry in $matchingEntries) {
                $foundAny = $true
                Write-Activity "Reading entry: $($entry.Name)" -Level Detail
                [System.Windows.Forms.Application]::DoEvents()
                $entryStream = $null; $entryReader = $null
                try {
                    $entryStream = $entry.Open()
                    $entryReader = [System.IO.StreamReader]::new($entryStream)
                    $lineNumber = 0; $entryMatches = 0
                    while (($line = $entryReader.ReadLine()) -ne $null) {
                        $lineNumber++
                        if ($line.IndexOf($SearchStr, [System.StringComparison]::OrdinalIgnoreCase) -ge 0) {
                            $fileDate = $entry.LastWriteTime.ToString('yyyy-MM-dd HH:mm')
                            $trimmed  = $line.Trim()
                            $displayName = "{0} > {1}" -f $zipFile.Name, $entry.Name
                            $Results.Add([PSCustomObject]@{ Station=$Station; Source='Archive'; FileName=$displayName; FileDate=$fileDate; LineNumber=$lineNumber; MatchedLine=$trimmed })
                            [void]$dgv.Rows.Add($Station,'Archive',$displayName,$fileDate,$lineNumber,$trimmed)
                            $entryMatches++
                        }
                    }
                    Write-Activity "Scanned $lineNumber lines, $entryMatches match(es)" -Level Detail
                } finally {
                    if ($entryReader) { $entryReader.Dispose() }; if ($entryStream) { $entryStream.Dispose() }
                }
            }
        } catch {
            Write-Activity "Error reading $($zipFile.Name): $($_.Exception.Message)" -Level Error
            $Problems.Add("$Station - $($zipFile.Name): $($_.Exception.Message)")
        } finally {
            if ($archive) { $archive.Dispose() }; if ($zipStream) { $zipStream.Dispose() }
        }
    }
    Write-Activity "Archive summary: $zipsChecked opened, $zipsSkipped skipped, $zipsUnparsable unparsable" -Level Info
    return $foundAny
}

function Copy-PlainLogs {
    param([string]$UncPath,[string]$Station,[datetime]$DateFrom,[datetime]$DateTo,[string]$DestFolder,
          [System.Collections.Generic.List[string]]$Problems)
    if (-not (Test-Path -LiteralPath $UncPath)) { Write-Activity "Main folder not accessible: $UncPath" -Level Warning; return 0 }
    $allFiles = @(Get-ChildItem -LiteralPath $UncPath -File -Filter 'Medcon.Framework.Shell*' -ErrorAction SilentlyContinue)
    $files = @($allFiles | Where-Object { $_.LastWriteTime.Date -ge $DateFrom -and $_.LastWriteTime.Date -le $DateTo })
    Write-Activity "Main folder: $($files.Count) file(s) match" -Level Detail
    if ($files.Count -eq 0) { return 0 }
    if (-not (Test-Path -LiteralPath $DestFolder)) { [void](New-Item -ItemType Directory -Path $DestFolder -Force) }
    $copied = 0
    foreach ($file in $files) {
        try {
            $target = Join-Path $DestFolder $file.Name
            Copy-Item -LiteralPath $file.FullName -Destination $target -Force -ErrorAction Stop
            $sizeMB = [math]::Round($file.Length / 1MB, 2)
            Write-Activity "Copied: $($file.Name) ($sizeMB MB)" -Level Detail
            $copied++
        } catch {
            Write-Activity "Failed copying $($file.Name): $($_.Exception.Message)" -Level Error
            $Problems.Add("$Station - $($file.Name): $($_.Exception.Message)")
        }
        [System.Windows.Forms.Application]::DoEvents()
    }
    return $copied
}

function Extract-ArchiveLogs {
    param([string]$ArchivePath,[string]$Station,[datetime]$DateFrom,[datetime]$DateTo,[string]$DestFolder,
          [System.Collections.Generic.List[string]]$Problems)
    if (-not (Test-Path -LiteralPath $ArchivePath)) { Write-Activity "Archive folder not accessible: $ArchivePath" -Level Warning; return 0 }
    $zipFiles = @(Get-ChildItem -LiteralPath $ArchivePath -File -Filter '*.zip' -ErrorAction SilentlyContinue)
    if ($zipFiles.Count -eq 0) { return 0 }
    $extracted = 0
    foreach ($zipFile in $zipFiles) {
        $zipRange = Get-ZipDateRange -ZipName $zipFile.Name
        if ($null -ne $zipRange) {
            if (-not (Test-DateRangesOverlap -SearchFrom $DateFrom -SearchTo $DateTo -ZipFrom $zipRange.Start -ZipTo $zipRange.End)) { continue }
        }
        $zipStream = $null; $archive = $null
        try {
            $zipStream = [System.IO.FileStream]::new($zipFile.FullName, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
            $archive = [System.IO.Compression.ZipArchive]::new($zipStream, [System.IO.Compression.ZipArchiveMode]::Read)
            $matchingEntries = @($archive.Entries | Where-Object {
                $_.Name -like 'Medcon.Framework.Shell*' -and
                $_.LastWriteTime.Date -ge $DateFrom -and $_.LastWriteTime.Date -le $DateTo
            })
            if ($matchingEntries.Count -eq 0) { continue }
            $zipBaseName = [System.IO.Path]::GetFileNameWithoutExtension($zipFile.Name)
            $zipDestRoot = Join-Path $DestFolder ("Archive_" + $zipBaseName)
            if (-not (Test-Path -LiteralPath $zipDestRoot)) { [void](New-Item -ItemType Directory -Path $zipDestRoot -Force) }
            Write-Activity "Extracting from $($zipFile.Name): $($matchingEntries.Count) entry(ies)" -Level Detail
            foreach ($entry in $matchingEntries) {
                try {
                    $target = Join-Path $zipDestRoot $entry.Name
                    [System.IO.Compression.ZipFileExtensions]::ExtractToFile($entry, $target, $true)
                    Write-Activity "Extracted: $($entry.Name)" -Level Detail
                    $extracted++
                } catch {
                    Write-Activity "Failed extracting $($entry.Name): $($_.Exception.Message)" -Level Error
                    $Problems.Add("$Station - $($zipFile.Name) > $($entry.Name): $($_.Exception.Message)")
                }
                [System.Windows.Forms.Application]::DoEvents()
            }
        } catch {
            Write-Activity "Error opening $($zipFile.Name): $($_.Exception.Message)" -Level Error
            $Problems.Add("$Station - $($zipFile.Name): $($_.Exception.Message)")
        } finally {
            if ($archive) { $archive.Dispose() }; if ($zipStream) { $zipStream.Dispose() }
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

# Date presets
$btnPresetToday.Add_Click({
    $today = (Get-Date).Date
    $dtpDateFrom.Value = $today; $dtpDateTo.Value = $today
})
$btnPresetYday.Add_Click({
    $y = (Get-Date).Date.AddDays(-1)
    $dtpDateFrom.Value = $y; $dtpDateTo.Value = $y
})
$btnPreset7d.Add_Click({
    $today = (Get-Date).Date
    $dtpDateFrom.Value = $today.AddDays(-6); $dtpDateTo.Value = $today
})
$btnPreset30d.Add_Click({
    $today = (Get-Date).Date
    $dtpDateFrom.Value = $today.AddDays(-29); $dtpDateTo.Value = $today
})

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
    if ($list.Count -gt 0) { $txtStations.Text = (($list | Sort-Object) -join "`r`n") }
})
$btnBrowseOutput.Add_Click({
    $fbd = New-Object System.Windows.Forms.FolderBrowserDialog
    $fbd.Description  = 'Select output folder'
    $fbd.SelectedPath = $txtOutput.Text
    if ($fbd.ShowDialog() -eq 'OK') { $txtOutput.Text = $fbd.SelectedPath }
})
$btnClear.Add_Click({
    $dgv.Rows.Clear()
    $lblStatus.Text       = 'Ready'
    $lblStatus.ForeColor  = $clrMuted
    $progressBar.Value    = 0
    $btnExport.Enabled    = $false
    $txtFilter.Text       = ''
    $statusLabel.Text     = '  No results'
    $statusRowCount.Text  = 'Rows: 0'
    $script:searchResults = $null
    Set-StatusPill 'Ready' 'idle'
})

$btnClearLog.Add_Click({ $txtActivity.Clear() })
$btnSaveLog.Add_Click({
    $sfd = New-Object System.Windows.Forms.SaveFileDialog
    $sfd.Title    = 'Save Activity Log'
    $sfd.Filter   = 'Text files (*.txt)|*.txt'
    $sfd.FileName = "ActivityLog_$((Get-Date).ToString('yyyyMMdd_HHmmss')).txt"
    if ($sfd.ShowDialog() -eq 'OK') {
        [System.IO.File]::WriteAllText($sfd.FileName, $txtActivity.Text, [System.Text.Encoding]::UTF8)
    }
})

$txtFilter.Add_TextChanged({
    $filterText = $txtFilter.Text.Trim()
    $dgv.SuspendLayout()
    foreach ($row in $dgv.Rows) {
        if ([string]::IsNullOrEmpty($filterText)) { $row.Visible = $true }
        else {
            $match = $false
            foreach ($cell in $row.Cells) {
                if ($cell.Value -and $cell.Value.ToString().IndexOf($filterText, [System.StringComparison]::OrdinalIgnoreCase) -ge 0) { $match = $true; break }
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
    if ($stations.Count -eq 0) { [System.Windows.Forms.MessageBox]::Show('Please enter at least one station name.', 'Missing Input', 'OK', 'Warning'); return }
    if ([string]::IsNullOrWhiteSpace($txtOutput.Text) -or -not (Test-Path -LiteralPath $txtOutput.Text)) {
        [System.Windows.Forms.MessageBox]::Show('Please select a valid output folder.', 'Missing Input', 'OK', 'Warning'); return
    }
    $dateFrom = $dtpDateFrom.Value.Date; $dateTo = $dtpDateTo.Value.Date
    if ($dateTo -lt $dateFrom) { [System.Windows.Forms.MessageBox]::Show('"Date To" cannot be earlier than "Date From".', 'Invalid Range', 'OK', 'Warning'); return }
    $daySpan = ($dateTo - $dateFrom).Days + 1
    if ($daySpan -gt 31) {
        $confirm = [System.Windows.Forms.MessageBox]::Show("You selected $daySpan days. Downloading may be slow and use a lot of disk space.`nContinue?", 'Large Range', 'YesNo', 'Question')
        if ($confirm -ne 'Yes') { return }
    }

    $dateTag    = if ($dateFrom -eq $dateTo) { $dateFrom.ToString('yyyyMMdd') } else { "$($dateFrom.ToString('yyyyMMdd'))-$($dateTo.ToString('yyyyMMdd'))" }
    $timestamp  = (Get-Date).ToString('yyyyMMdd_HHmmss')
    $rootFolder = Join-Path $txtOutput.Text ("MedconLogs_{0}_{1}" -f $dateTag, $timestamp)
    [void](New-Item -ItemType Directory -Path $rootFolder -Force)

    $btnSearch.Enabled = $false; $btnDownload.Enabled = $false
    $progressBar.Value = 0; $progressBar.Maximum = $stations.Count
    Set-StatusPill 'Downloading...' 'working'

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

        $copied    = Copy-PlainLogs    -UncPath $uncPath -Station $station -DateFrom $dateFrom -DateTo $dateTo -DestFolder $stationFolder -Problems $problems
        $extracted = Extract-ArchiveLogs -ArchivePath $archivePath -Station $station -DateFrom $dateFrom -DateTo $dateTo -DestFolder $stationFolder -Problems $problems

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
    $lblStatus.ForeColor = if ($totalFiles -gt 0) { $clrSuccess } else { [System.Drawing.Color]::OrangeRed }
    $statusLabel.Text    = "  Downloaded $totalFiles file(s) to $rootFolder"
    if ($totalFiles -gt 0) { Set-StatusPill "Downloaded $totalFiles" 'ok' } else { Set-StatusPill 'No files' 'warn' }

    Write-Activity "" -Level Info
    Write-Activity "===========================================" -Level Step
    Write-Activity "DOWNLOAD COMPLETE - $totalFiles file(s) in $elapsed" -Level Success
    Write-Activity "Location: $rootFolder" -Level Info
    Write-Activity "===========================================" -Level Step

    $btnSearch.Enabled = $true; $btnDownload.Enabled = $true

    if ($totalFiles -gt 0) {
        $open = [System.Windows.Forms.MessageBox]::Show("Downloaded $totalFiles file(s).`nOpen the folder now?", 'Download Complete', 'YesNo', 'Question')
        if ($open -eq 'Yes') { Invoke-Item $rootFolder }
    } elseif ($problems.Count -gt 0) {
        $msg = "No files were downloaded.`n`n" + ($problems -join "`n")
        [System.Windows.Forms.MessageBox]::Show($msg, 'Problems', 'OK', 'Warning')
    } else {
        [System.Windows.Forms.MessageBox]::Show('No matching files found in the selected date range.', 'No Files', 'OK', 'Information')
    }
})

# ── Search ──
$btnSearch.Add_Click({
    $stations = @(Get-StationList)
    if ($stations.Count -eq 0) { [System.Windows.Forms.MessageBox]::Show('Please enter at least one station name.', 'Missing Input', 'OK', 'Warning'); return }
    if ([string]::IsNullOrWhiteSpace($txtSearch.Text)) { [System.Windows.Forms.MessageBox]::Show('Please enter a search string.', 'Missing Input', 'OK', 'Warning'); return }
    if ([string]::IsNullOrWhiteSpace($txtOutput.Text) -or -not (Test-Path -LiteralPath $txtOutput.Text)) {
        [System.Windows.Forms.MessageBox]::Show('Please select a valid output folder.', 'Missing Input', 'OK', 'Warning'); return
    }
    $dateFrom = $dtpDateFrom.Value.Date; $dateTo = $dtpDateTo.Value.Date
    if ($dateTo -lt $dateFrom) { [System.Windows.Forms.MessageBox]::Show('"Date To" cannot be earlier than "Date From".', 'Invalid Range', 'OK', 'Warning'); return }
    $daySpan = ($dateTo - $dateFrom).Days + 1
    if ($daySpan -gt 31) {
        $confirm = [System.Windows.Forms.MessageBox]::Show("You selected a range of $daySpan days. This may take a while.`nContinue?", 'Large Range', 'YesNo', 'Question')
        if ($confirm -ne 'Yes') { return }
    }

    $dgv.Rows.Clear()
    $txtFilter.Text = ''
    $btnSearch.Enabled = $false; $btnDownload.Enabled = $false; $btnExport.Enabled = $false
    $progressBar.Value = 0
    Set-StatusPill 'Searching...' 'working'

    $searchString = $txtSearch.Text; $outputFolder = $txtOutput.Text
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
        $foundInMain = $false; $foundInArchive = $false; $stationReachable = $false

        if (Test-Path -LiteralPath $uncPath) {
            $stationReachable = $true
            $foundInMain = Search-PlainFiles -UncPath $uncPath -Station $station -DateFrom $dateFrom -DateTo $dateTo -SearchStr $searchString -Results $results -Problems $problems
        } else {
            Write-Activity "Main folder not reachable: $uncPath" -Level Error
        }

        $lblStatus.Text = "Scanning $station archive ($($i+1) of $($stations.Count))..."
        [System.Windows.Forms.Application]::DoEvents()
        Write-Activity "Now searching archive..." -Level Info

        $foundInArchive = Search-ArchiveZips -ArchivePath $archivePath -Station $station -DateFrom $dateFrom -DateTo $dateTo -SearchStr $searchString -Results $results -Problems $problems
        if ($foundInArchive) { $archiveHits++; $stationReachable = $true }

        if (-not $stationReachable) {
            $problems.Add("$station  -  path not accessible (neither LogEx nor LogArchive)")
        } elseif (-not $foundInMain -and -not $foundInArchive) {
            $rangeStr = if ($dateFrom -eq $dateTo) { $dateFrom.ToString('yyyy-MM-dd') } else { "$($dateFrom.ToString('yyyy-MM-dd')) to $($dateTo.ToString('yyyy-MM-dd'))" }
            $problems.Add("$station  -  no matching files for $rangeStr")
            Write-Activity "No matches found for $station in date range" -Level Warning
        }

        $stationSw.Stop()
        $stationMatches = $results.Count - $matchesBefore
        Write-Activity "Station done in $($stationSw.Elapsed.ToString('mm\:ss\.f')) - $stationMatches match(es)" -Level Success
        [System.Windows.Forms.Application]::DoEvents()
    }

    $sw.Stop()
    $elapsed = $sw.Elapsed.ToString('mm\:ss\.f')
    $script:searchResults = $results

    $dateTag = if ($dateFrom -eq $dateTo) { $dateFrom.ToString('yyyyMMdd') } else { "$($dateFrom.ToString('yyyyMMdd'))-$($dateTo.ToString('yyyyMMdd'))" }
    $safeString = ($searchString -replace '[^a-zA-Z0-9_\-]+', '_').Trim('_')
    if (-not $safeString) { $safeString = 'search' }
    $script:csvPath = Join-Path $outputFolder ("LogSearch_{0}_{1}.csv" -f $dateTag, $safeString)

    $archiveNote = if ($archiveHits -gt 0) { "  ($archiveHits from archive)" } else { '' }
    $lblStatus.Text      = "Done in $elapsed  -  $($results.Count) match(es), $($problems.Count) problem(s)$archiveNote"
    $lblStatus.ForeColor = if ($results.Count -gt 0) { $clrSuccess } else { [System.Drawing.Color]::OrangeRed }
    $statusLabel.Text    = "  Search complete: $($results.Count) match(es) across $($stations.Count) station(s)$archiveNote"
    if ($results.Count -gt 0) { Set-StatusPill "$($results.Count) match(es)" 'ok' } else { Set-StatusPill 'No matches' 'warn' }

    Write-Activity "" -Level Info
    Write-Activity "===========================================" -Level Step
    Write-Activity "SEARCH COMPLETE - $($results.Count) match(es) in $elapsed" -Level Success
    if ($problems.Count -gt 0) { Write-Activity "$($problems.Count) problem(s) encountered" -Level Warning }
    Write-Activity "===========================================" -Level Step

    $btnSearch.Enabled = $true; $btnDownload.Enabled = $true
    $btnExport.Enabled = ($results.Count -gt 0)
    Update-RowCount

    foreach ($row in $dgv.Rows) {
        if ($row.Cells['Source'].Value -eq 'Archive') { $row.DefaultCellStyle.ForeColor = $clrArchive }
    }
    if ($problems.Count -gt 0) {
        $msg = "Search complete with $($problems.Count) problem(s):`n`n" + ($problems -join "`n")
        [System.Windows.Forms.MessageBox]::Show($msg, 'Problems Encountered', 'OK', 'Warning')
    }
})

# ── Export CSV ──
$btnExport.Add_Click({
    if (-not $script:searchResults -or $script:searchResults.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show('No results to export.', 'Export', 'OK', 'Information'); return
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

$txtSearch.Add_KeyDown({ if ($_.KeyCode -eq 'Return') { $btnSearch.PerformClick() } })

# ══════════════════════════════════════════════
#  LAUNCH
# ══════════════════════════════════════════════
[void]$form.ShowDialog()
