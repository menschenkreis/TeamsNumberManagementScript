# =============================================================================
# UI COMPONENTS MODULE - GUI Construction and Event Handlers
# =============================================================================
# This module contains all GUI construction code and general UI event handlers
# for the Teams Phone Manager application.
#
# Components included:
# - Main Form
# - Configuration GroupBox (Orange API settings)
# - Execution Log GroupBox
# - Tag Statistics GroupBox and DataGridView
# - Grid Statistics GroupBox
# - Progress Bar
# - Top Actions GroupBox (Data Operations)
# - Main DataGridView
# - Context Menus
# - Bottom Actions GroupBox (Tags, Teams Actions, Number Actions, Orange Actions)
# - General UI Event Handlers
#
# NOTE: This module must be loaded AFTER Core, HelperFunctions, SettingsManagement,
#       and TagsAndStats modules as it depends on their functions.
# =============================================================================

#region 8. GUI Construction
$global:form = New-Object System.Windows.Forms.Form
$global:form.Text = "Teams Phone Manager v56.2"
$global:form.Size = New-Object System.Drawing.Size(1600, 920)
$global:form.WindowState = "Maximized" # START MAXIMIZED
$global:form.StartPosition = "CenterScreen"
$global:form.BackColor = "#F0F0F0"

# --- CONFIGURATION GROUP ---
$grpConfig = New-Object System.Windows.Forms.GroupBox
$grpConfig.Location = New-Object System.Drawing.Point(20, 10)
# UPDATED: Width to 1240 to match Grid
$grpConfig.Size = New-Object System.Drawing.Size(1240, 70)
$grpConfig.Text = "Orange API Configuration"
$grpConfig.Anchor = "Top, Left, Right"

# Layout
$lblAuth = New-Object System.Windows.Forms.Label; $lblAuth.Location = New-Object System.Drawing.Point(10, 25); $lblAuth.Size = New-Object System.Drawing.Size(80, 20); $lblAuth.Text = "Auth Header:"
$global:txtOrangeAuth = New-Object System.Windows.Forms.TextBox; $global:txtOrangeAuth.Location = New-Object System.Drawing.Point(90, 22); $global:txtOrangeAuth.Size = New-Object System.Drawing.Size(200, 20); $global:txtOrangeAuth.PasswordChar = "*"

$lblCust = New-Object System.Windows.Forms.Label; $lblCust.Location = New-Object System.Drawing.Point(300, 25); $lblCust.Size = New-Object System.Drawing.Size(90, 20); $lblCust.Text = "Customer Id:"
$global:txtOrangeCust = New-Object System.Windows.Forms.TextBox; $global:txtOrangeCust.Location = New-Object System.Drawing.Point(390, 22); $global:txtOrangeCust.Size = New-Object System.Drawing.Size(60, 20); $global:txtOrangeCust.Text = ""

$lblKey = New-Object System.Windows.Forms.Label; $lblKey.Location = New-Object System.Drawing.Point(460, 25); $lblKey.Size = New-Object System.Drawing.Size(50, 20); $lblKey.Text = "API Key:"
$global:txtOrangeKey = New-Object System.Windows.Forms.TextBox; $global:txtOrangeKey.Location = New-Object System.Drawing.Point(510, 22); $global:txtOrangeKey.Size = New-Object System.Drawing.Size(120, 20); $global:txtOrangeKey.PasswordChar = "*"

$lblProxy = New-Object System.Windows.Forms.Label; $lblProxy.Location = New-Object System.Drawing.Point(640, 25); $lblProxy.Size = New-Object System.Drawing.Size(50, 20); $lblProxy.Text = "Proxy:"
$global:txtProxy = New-Object System.Windows.Forms.TextBox; $global:txtProxy.Location = New-Object System.Drawing.Point(690, 22); $global:txtProxy.Size = New-Object System.Drawing.Size(150, 20); $global:txtProxy.Text = ""

# Buttons
$btnLoadXml = New-Object System.Windows.Forms.Button; $btnLoadXml.Location = New-Object System.Drawing.Point(850, 19); $btnLoadXml.Size = New-Object System.Drawing.Size(80, 25); $btnLoadXml.Text = "Load XML"
$btnSaveXml = New-Object System.Windows.Forms.Button; $btnSaveXml.Location = New-Object System.Drawing.Point(935, 19); $btnSaveXml.Size = New-Object System.Drawing.Size(80, 25); $btnSaveXml.Text = "Save XML"

# UPDATED: Moved Help Button to the far right (1120) to utilize new width
$btnOrangeHelp = New-Object System.Windows.Forms.Button
$btnOrangeHelp.Location = New-Object System.Drawing.Point(1120, 19)
$btnOrangeHelp.Size = New-Object System.Drawing.Size(100, 25)
$btnOrangeHelp.Text = "Get API Help"
$btnOrangeHelp.BackColor = "#17a2b8"
$btnOrangeHelp.ForeColor = "White"
$btnOrangeHelp.Add_Click({ Start-Process "https://developer.orange.com/apis/businesstalk/getting-started" })

# Events
$btnLoadXml.Add_Click({
    $ofd = New-Object System.Windows.Forms.OpenFileDialog; $ofd.Filter = "XML Files (*.xml)|*.xml"; $ofd.Title = "Select Settings XML"
    if ($ofd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { Import-SettingsXml -Path $ofd.FileName }
})
$btnSaveXml.Add_Click({ Export-SettingsXml })

$grpConfig.Controls.AddRange(@($lblAuth, $global:txtOrangeAuth, $lblCust, $global:txtOrangeCust, $lblKey, $global:txtOrangeKey, $lblProxy, $global:txtProxy, $btnLoadXml, $btnSaveXml, $btnOrangeHelp))

# -- Log Box --
$grpLog = New-Object System.Windows.Forms.GroupBox; $grpLog.Location = New-Object System.Drawing.Point(1280, 10); $grpLog.Size = New-Object System.Drawing.Size(290, 380); $grpLog.Text = "Execution Log"; $grpLog.Anchor = "Top, Right"
$global:txtLog = New-Object System.Windows.Forms.TextBox; $global:txtLog.Location = New-Object System.Drawing.Point(10, 20); $global:txtLog.Size = New-Object System.Drawing.Size(270, 350); $global:txtLog.Multiline = $true; $global:txtLog.ScrollBars = "Vertical"; $global:txtLog.ReadOnly = $true; $global:txtLog.BackColor = "White"; $global:txtLog.Anchor = "Top, Bottom, Left, Right"
$grpLog.Controls.Add($global:txtLog)

# -- Tag Stats Box --
$grpTagStats = New-Object System.Windows.Forms.GroupBox
$grpTagStats.Location = New-Object System.Drawing.Point(1280, 400)
$grpTagStats.Size = New-Object System.Drawing.Size(290, 260)
$grpTagStats.Text = "Tag Statistics (From Data)"
$grpTagStats.Anchor = "Top, Right"

$global:dgvTagStats = New-Object System.Windows.Forms.DataGridView
$global:dgvTagStats.Location = New-Object System.Drawing.Point(10, 20)
$global:dgvTagStats.Size = New-Object System.Drawing.Size(270, 230)
$global:dgvTagStats.AllowUserToAddRows = $false
$global:dgvTagStats.ReadOnly = $true
$global:dgvTagStats.RowHeadersVisible = $false
$global:dgvTagStats.SelectionMode = "FullRowSelect"
$global:dgvTagStats.AutoSizeColumnsMode = "Fill"
$global:dgvTagStats.Anchor = "Top, Bottom, Left, Right"
# Disable default OS styling to allow custom colors
$global:dgvTagStats.EnableHeadersVisualStyles = $false

# Set standard header colors
$global:dgvTagStats.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(240, 240, 240)
$global:dgvTagStats.ColumnHeadersDefaultCellStyle.ForeColor = [System.Drawing.Color]::Black

# Force the "Selection" color to match the "Normal" color so it doesn't look highlighted
$global:dgvTagStats.ColumnHeadersDefaultCellStyle.SelectionBackColor = [System.Drawing.Color]::FromArgb(240, 240, 240)
$global:dgvTagStats.ColumnHeadersDefaultCellStyle.SelectionForeColor = [System.Drawing.Color]::Black
[void]$global:dgvTagStats.Columns.Add("Location", "Tag")
[void]$global:dgvTagStats.Columns.Add("Total", "Total")
[void]$global:dgvTagStats.Columns.Add("Free", "Free")
$global:dgvTagStats.Columns[0].FillWeight = 50; $global:dgvTagStats.Columns[1].FillWeight = 25; $global:dgvTagStats.Columns[2].FillWeight = 25

# Initialize Stats Grid Rows (with defaults)
Update-TagUI

# NEW: Context Menu for Tag Stats Export
$ctxTagStats = New-Object System.Windows.Forms.ContextMenuStrip
$miExportTags = $ctxTagStats.Items.Add("Export to CSV")
$miExportTags.Add_Click({
    $sfd = New-Object System.Windows.Forms.SaveFileDialog
    $sfd.Filter = "CSV Files (*.csv)|*.csv"
    $sfd.Title = "Save Tag Stats"
    $sfd.FileName = "TagStats_$(Get-Date -Format 'yyyyMMdd_HHmm').csv"

    if ($sfd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $path = $sfd.FileName
        Write-Log "Exporting Tag Stats to $path..."
        try {
            $exportList = New-Object System.Collections.ArrayList
            foreach ($row in $global:dgvTagStats.Rows) {
                $obj = [Ordered]@{
                    Tag   = $row.Cells[0].Value
                    Total = $row.Cells[1].Value
                    Free  = $row.Cells[2].Value
                }
                [void]$exportList.Add([PSCustomObject]$obj)
            }
            $exportList | Export-Csv -Path $path -NoTypeInformation -Delimiter "," -Encoding UTF8
            Write-Log "Tag Stats exported successfully."
            [System.Windows.Forms.MessageBox]::Show("Tag Statistics exported successfully!", "Done")
        } catch {
            Write-Log "Export Error: $($_.Exception.Message)"
            [System.Windows.Forms.MessageBox]::Show("Export Failed: $($_.Exception.Message)", "Error")
        }
    }
})
$global:dgvTagStats.ContextMenuStrip = $ctxTagStats

$grpTagStats.Controls.Add($global:dgvTagStats)

# -- General Stats Box --
$grpStats = New-Object System.Windows.Forms.GroupBox
$grpStats.Location = New-Object System.Drawing.Point(1280, 670)
$grpStats.Size = New-Object System.Drawing.Size(290, 100)
$grpStats.Text = "Grid Statistics"
$grpStats.Anchor = "Top, Right"

$lblTitleTotal = New-Object System.Windows.Forms.Label; $lblTitleTotal.Location = New-Object System.Drawing.Point(15, 25); $lblTitleTotal.Size = New-Object System.Drawing.Size(140, 20); $lblTitleTotal.Text = "Total Rows:"; $lblTitleTotal.Font = New-Object System.Drawing.Font("Consolas", 10)
$lblTitleDisp = New-Object System.Windows.Forms.Label; $lblTitleDisp.Location = New-Object System.Drawing.Point(15, 48); $lblTitleDisp.Size = New-Object System.Drawing.Size(140, 20); $lblTitleDisp.Text = "Displayed Rows:"; $lblTitleDisp.Font = New-Object System.Drawing.Font("Consolas", 10)
$lblTitleSel = New-Object System.Windows.Forms.Label; $lblTitleSel.Location = New-Object System.Drawing.Point(15, 71); $lblTitleSel.Size = New-Object System.Drawing.Size(140, 20); $lblTitleSel.Text = "Selected Rows:"; $lblTitleSel.Font = New-Object System.Drawing.Font("Consolas", 10)
$global:lblValTotal = New-Object System.Windows.Forms.Label; $global:lblValTotal.Location = New-Object System.Drawing.Point(160, 25); $global:lblValTotal.Size = New-Object System.Drawing.Size(100, 20); $global:lblValTotal.Text = "0"; $global:lblValTotal.Font = New-Object System.Drawing.Font("Consolas", 10)
$global:lblValDisp = New-Object System.Windows.Forms.Label; $global:lblValDisp.Location = New-Object System.Drawing.Point(160, 48); $global:lblValDisp.Size = New-Object System.Drawing.Size(100, 20); $global:lblValDisp.Text = "0"; $global:lblValDisp.Font = New-Object System.Drawing.Font("Consolas", 10)
$global:lblValSel = New-Object System.Windows.Forms.Label; $global:lblValSel.Location = New-Object System.Drawing.Point(160, 71); $global:lblValSel.Size = New-Object System.Drawing.Size(100, 20); $global:lblValSel.Text = "0"; $global:lblValSel.Font = New-Object System.Drawing.Font("Consolas", 10)
$grpStats.Controls.AddRange(@($lblTitleTotal, $lblTitleDisp, $lblTitleSel, $global:lblValTotal, $global:lblValDisp, $global:lblValSel))

# -- Progress Bar --
$progressBar = New-Object System.Windows.Forms.ProgressBar; $progressBar.Location = New-Object System.Drawing.Point(20, 780); $progressBar.Size = New-Object System.Drawing.Size(1240, 10); $progressBar.Style = "Continuous"; $progressBar.Anchor = "Bottom, Left, Right"
$global:form.Controls.Add($progressBar)

# -- Top Controls (Wrapped in GroupBox) --
$grpTopActions = New-Object System.Windows.Forms.GroupBox
$grpTopActions.Location = New-Object System.Drawing.Point(20, 90)
# UPDATED: Width to 1240 to match Grid and Config Box
$grpTopActions.Size = New-Object System.Drawing.Size(1240, 60)
$grpTopActions.Text = "Data Operations"
$grpTopActions.Anchor = "Top, Left, Right"

$innerY = 20 # Y-Position for controls inside the GroupBox

# 1. Connection Buttons (Left)
$btnConnect = New-Object System.Windows.Forms.Button; $btnConnect.Location = New-Object System.Drawing.Point(10, $innerY); $btnConnect.Size = New-Object System.Drawing.Size(80, 30); $btnConnect.Text = "1. Connect"; $btnConnect.BackColor = "#0078D7"; $btnConnect.ForeColor = "White"
$btnFetchData = New-Object System.Windows.Forms.Button; $btnFetchData.Location = New-Object System.Drawing.Point(95, $innerY); $btnFetchData.Size = New-Object System.Drawing.Size(120, 30); $btnFetchData.Text = "2. Get Data"; $btnFetchData.Enabled = $false
$btnSyncOrange = New-Object System.Windows.Forms.Button; $btnSyncOrange.Location = New-Object System.Drawing.Point(220, $innerY); $btnSyncOrange.Size = New-Object System.Drawing.Size(100, 30); $btnSyncOrange.Text = "3. Re-Sync"; $btnSyncOrange.BackColor = "#FF8C00"; $btnSyncOrange.ForeColor = "White"; $btnSyncOrange.Enabled = $false

# Separator 1
$sepTop1 = New-Object System.Windows.Forms.Label; $sepTop1.Location = New-Object System.Drawing.Point(325, $innerY); $sepTop1.Size = New-Object System.Drawing.Size(2, 30); $sepTop1.BorderStyle = "Fixed3D"

# 2. Export & Free Number
$btnExport = New-Object System.Windows.Forms.Button; $btnExport.Location = New-Object System.Drawing.Point(335, $innerY); $btnExport.Size = New-Object System.Drawing.Size(90, 30); $btnExport.Text = "Export"; $btnExport.BackColor = "#28a745"; $btnExport.ForeColor = "White"; $btnExport.Add_Click({ Export-SelectedToCSV })
$btnGetFree = New-Object System.Windows.Forms.Button; $btnGetFree.Location = New-Object System.Drawing.Point(430, $innerY); $btnGetFree.Size = New-Object System.Drawing.Size(90, 30); $btnGetFree.Text = "Get Free #"; $btnGetFree.BackColor = "#006400"; $btnGetFree.ForeColor = "White"

# Separator 2
$sepFilter = New-Object System.Windows.Forms.Label; $sepFilter.Location = New-Object System.Drawing.Point(525, $innerY); $sepFilter.Size = New-Object System.Drawing.Size(2, 30); $sepFilter.BorderStyle = "Fixed3D"

# 3. Filter Area
$lblFilter = New-Object System.Windows.Forms.Label; $lblFilterY = $innerY + 5; $lblFilter.Location = New-Object System.Drawing.Point(535, $lblFilterY); $lblFilter.Size = New-Object System.Drawing.Size(35, 20); $lblFilter.Text = "Text:"
$txtFilter = New-Object System.Windows.Forms.TextBox; $txtFilterY = $innerY + 2; $txtFilter.Location = New-Object System.Drawing.Point(570, $txtFilterY); $txtFilter.Size = New-Object System.Drawing.Size(100, 20)

# Tag Filter Dropdown
$lblFilterTag = New-Object System.Windows.Forms.Label; $lblFilterTag.Location = New-Object System.Drawing.Point(675, $lblFilterY); $lblFilterTag.Size = New-Object System.Drawing.Size(30, 20); $lblFilterTag.Text = "Tag:"
$global:cmbFilterTag = New-Object System.Windows.Forms.ComboBox; $global:cmbFilterTag.Location = New-Object System.Drawing.Point(705, $txtFilterY); $global:cmbFilterTag.Size = New-Object System.Drawing.Size(90, 20); $global:cmbFilterTag.DropDownStyle = "DropDownList"
$global:cmbFilterTag.Items.Add("All")
$global:cmbFilterTag.SelectedIndex = 0
$btnApplyFilter = New-Object System.Windows.Forms.Button; $btnApplyFilter.Location = New-Object System.Drawing.Point(800, $innerY); $btnApplyFilter.Size = New-Object System.Drawing.Size(80, 30); $btnApplyFilter.Text = "Apply Filter"; $btnApplyFilter.BackColor = "#A9A9A9"; $btnApplyFilter.ForeColor = "White"

# 4. Toggles & Columns (Shifted Right to use the 1240 width)
$btnToggleUnassigned = New-Object System.Windows.Forms.Button
# UPDATED X: 900
$btnToggleUnassigned.Location = New-Object System.Drawing.Point(900, $innerY)
$btnToggleUnassigned.Size = New-Object System.Drawing.Size(120, 30)
$btnToggleUnassigned.Text = "Hide Unassigned"
$btnToggleUnassigned.BackColor = "#e0e0e0"

$btnSelectCols = New-Object System.Windows.Forms.Button
# UPDATED X: 1040
$btnSelectCols.Location = New-Object System.Drawing.Point(1040, $innerY)
$btnSelectCols.Size = New-Object System.Drawing.Size(70, 30); $btnSelectCols.Text = "Columns"; $btnSelectCols.BackColor = "#808080"; $btnSelectCols.ForeColor = "White"

# Help Button
$btnHelp = New-Object System.Windows.Forms.Button
# UPDATED X: 1130
$btnHelp.Location = New-Object System.Drawing.Point(1130, $innerY)
$btnHelp.Size = New-Object System.Drawing.Size(60, 30)
$btnHelp.Text = "Help"
$btnHelp.BackColor = "#17a2b8"
$btnHelp.ForeColor = "White"

# --- Re-Add Logic Blocks ---
$actionApplyFilter = {
    if ($dataGridView.DataSource -is [System.Data.DataTable]) {
        $view = $dataGridView.DataSource.DefaultView; $filterParts = New-Object System.Collections.ArrayList
        $rawVal = $txtFilter.Text.Trim();
        if (-not [string]::IsNullOrWhiteSpace($rawVal)) {
            $safeVal = $rawVal.Replace("'", "''").Replace("[", "[[]").Replace("*", "[*]").Replace("%", "[%]"); $cols = $dataGridView.DataSource.Columns; $textFilterParts = New-Object System.Collections.ArrayList
            foreach ($col in $cols) { [void]$textFilterParts.Add("[$($col.ColumnName)] LIKE '%$safeVal%'") }; [void]$filterParts.Add("(" + ($textFilterParts -join " OR ") + ")")
        }
        if ($global:hideUnassigned) { [void]$filterParts.Add("(UserPrincipalName IS NOT NULL AND UserPrincipalName <> '')") }
        $selectedTag = $global:cmbFilterTag.SelectedItem; if ($selectedTag -and $selectedTag -ne "All") { [void]$filterParts.Add("([Tag] LIKE '%$selectedTag%')") }
        if ($filterParts.Count -gt 0) { try { $view.RowFilter = $filterParts -join " AND " } catch {} } else { $view.RowFilter = "" }; Update-Stats
    }
}
$btnApplyFilter.Add_Click($actionApplyFilter); $txtFilter.Add_KeyDown({ param($s, $e) if ($e.KeyCode -eq [System.Windows.Forms.Keys]::Enter) { & $actionApplyFilter; $e.SuppressKeyPress = $true } }); $global:cmbFilterTag.Add_SelectionChangeCommitted($actionApplyFilter)
$btnToggleUnassigned.Add_Click({ $global:hideUnassigned = -not $global:hideUnassigned; if ($global:hideUnassigned) { $btnToggleUnassigned.Text = "Hide Unassigned"; $btnToggleUnassigned.BackColor = "#b3d9ff" } else { $btnToggleUnassigned.Text = "Hide Unassigned"; $btnToggleUnassigned.BackColor = "#e0e0e0" }; & $actionApplyFilter })

$btnGetFree.Add_Click({
    if ($dataGridView.Rows.Count -eq 0) { return }
    $startIndex = 0; if ($dataGridView.SelectedRows.Count -gt 0) { $startIndex = $dataGridView.SelectedRows[$dataGridView.SelectedRows.Count - 1].Index + 1 }
    $foundIndex = -1; for ($i = $startIndex; $i -lt $dataGridView.Rows.Count; $i++) { $upn = [string]$dataGridView.Rows[$i].Cells["UserPrincipalName"].Value; if ([string]::IsNullOrWhiteSpace($upn)) { $foundIndex = $i; break } }
    if ($foundIndex -eq -1) { for ($i = 0; $i -lt $startIndex; $i++) { $upn = [string]$dataGridView.Rows[$i].Cells["UserPrincipalName"].Value; if ([string]::IsNullOrWhiteSpace($upn)) { $foundIndex = $i; break } } }
    if ($foundIndex -ne -1) { $dataGridView.ClearSelection(); $dataGridView.Rows[$foundIndex].Selected = $true; $dataGridView.FirstDisplayedScrollingRowIndex = $foundIndex; Update-Stats } else { [System.Windows.Forms.MessageBox]::Show("No unassigned numbers found.", "Info") }
})

$btnHelp.Add_Click({
    $fHelp = New-Object System.Windows.Forms.Form; $fHelp.Text = "Help"; $fHelp.Size = New-Object System.Drawing.Size(800, 600); $fHelp.StartPosition = "CenterParent"
    $txtHelp = New-Object System.Windows.Forms.TextBox; $txtHelp.Multiline = $true; $txtHelp.ReadOnly = $true; $txtHelp.Location = New-Object System.Drawing.Point(10, 10); $txtHelp.Size = New-Object System.Drawing.Size(760, 490); $txtHelp.ScrollBars = "Vertical"; $txtHelp.BackColor = "White"; $txtHelp.Font = New-Object System.Drawing.Font("Consolas", 10)
    $basePath = if ($PSScriptRoot) { $PSScriptRoot } else { $PWD.Path }; $helpPath = Join-Path $basePath "help.txt"; if (Test-Path $helpPath) { try { $txtHelp.Text = Get-Content $helpPath -Raw -Encoding UTF8 } catch { $txtHelp.Text = "Error: $($_.Exception.Message)" } } else { $txtHelp.Text = "help.txt not found." }
    $btnClose = New-Object System.Windows.Forms.Button; $btnClose.Text = "Close"; $btnClose.Location = New-Object System.Drawing.Point(350, 520); $btnClose.DialogResult = "OK"; $fHelp.Controls.AddRange(@($txtHelp, $btnClose)); $fHelp.Add_Shown({ $txtHelp.Select(0, 0); $btnClose.Focus() }); $fHelp.ShowDialog()
})

$grpTopActions.Controls.AddRange(@($btnConnect, $btnFetchData, $btnSyncOrange, $sepTop1, $btnExport, $btnGetFree, $sepFilter, $lblFilter, $txtFilter, $lblFilterTag, $global:cmbFilterTag, $btnApplyFilter, $btnToggleUnassigned, $btnSelectCols, $btnHelp))

# -- Grid (Y Position shifted down to 160 to accommodate groupbox) --
$gridY = 160
$dataGridView = New-Object System.Windows.Forms.DataGridView; $dataGridView.Location = New-Object System.Drawing.Point(20, $gridY); $dataGridView.Size = New-Object System.Drawing.Size(1240, 610); $dataGridView.Anchor = "Top, Bottom, Left, Right"; $dataGridView.AllowUserToAddRows = $false; $dataGridView.SelectionMode = "FullRowSelect"; $dataGridView.MultiSelect = $true; $dataGridView.ReadOnly = $true; $dataGridView.AutoSizeColumnsMode = "AllCells"
$dataGridView.DefaultCellStyle.SelectionBackColor = [System.Drawing.Color]::FromArgb(224, 224, 224)
$dataGridView.DefaultCellStyle.SelectionForeColor = [System.Drawing.Color]::Black

$dataGridView.EnableHeadersVisualStyles = $false
$dataGridView.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(240, 240, 240)
$dataGridView.ColumnHeadersDefaultCellStyle.ForeColor = [System.Drawing.Color]::Black
$dataGridView.ColumnHeadersDefaultCellStyle.SelectionBackColor = [System.Drawing.Color]::FromArgb(240, 240, 240)
$dataGridView.ColumnHeadersDefaultCellStyle.SelectionForeColor = [System.Drawing.Color]::Black
$dataGridView.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)

$dataGridView.Add_SelectionChanged({ $btnAssign.Enabled = ($dataGridView.SelectedRows.Count -eq 1); Update-Stats })

# -- Context Menu --
$ctxMenu = New-Object System.Windows.Forms.ContextMenuStrip
$miRefreshOrange = $ctxMenu.Items.Add("Refresh Orange Info")
$miRefreshTeams = $ctxMenu.Items.Add("Refresh Teams Info")
$miChangePolicy = $ctxMenu.Items.Add("Change Voice Routing Policy")
# NEW: Grant Meeting Policy
$miGrantMeetingPolicy = $ctxMenu.Items.Add("Grant Meeting Policy")
$miEnableEV = $ctxMenu.Items.Add("Enable Enterprise Voice") # NEW
$ctxMenu.Items.Add("-")
$miForcePublish = $ctxMenu.Items.Add("Force Publish to OC")
$dataGridView.ContextMenuStrip = $ctxMenu

# -- Actions --
# UPDATED LAYOUT FOR VISUAL SEPARATION
$grpTag = New-Object System.Windows.Forms.GroupBox; $grpTag.Location = New-Object System.Drawing.Point(20, 800); $grpTag.Size = New-Object System.Drawing.Size(1550, 50); $grpTag.Text = "Actions"; $grpTag.Anchor = "Bottom, Left, Right"

# Group 1: Tags
$lblTagInput = New-Object System.Windows.Forms.Label; $lblTagInput.Location = New-Object System.Drawing.Point(10, 22); $lblTagInput.Size = New-Object System.Drawing.Size(70, 20); $lblTagInput.Text = "Select Tag:"
$global:cmbTag = New-Object System.Windows.Forms.ComboBox; $global:cmbTag.Location = New-Object System.Drawing.Point(80, 19); $global:cmbTag.Size = New-Object System.Drawing.Size(120, 20); $global:cmbTag.DropDownStyle = "DropDownList"
Update-TagUI # Initial Populate

$cbBlacklist = New-Object System.Windows.Forms.CheckBox; $cbBlacklist.Location = New-Object System.Drawing.Point(210, 19); $cbBlacklist.Size = New-Object System.Drawing.Size(70, 20); $cbBlacklist.Text = "Blacklist"
$cbReserved = New-Object System.Windows.Forms.CheckBox; $cbReserved.Location = New-Object System.Drawing.Point(290, 19); $cbReserved.Size = New-Object System.Drawing.Size(80, 20); $cbReserved.Text = "Reserved"
$cbPremium = New-Object System.Windows.Forms.CheckBox; $cbPremium.Location = New-Object System.Drawing.Point(380, 19); $cbPremium.Size = New-Object System.Drawing.Size(80, 20); $cbPremium.Text = "Premium"
$btnApplyTag = New-Object System.Windows.Forms.Button; $btnApplyTag.Location = New-Object System.Drawing.Point(470, 17); $btnApplyTag.Size = New-Object System.Drawing.Size(80, 25); $btnApplyTag.Text = "Apply Tag"

# NEW BUTTON: Remove Tags
$btnRemoveTags = New-Object System.Windows.Forms.Button; $btnRemoveTags.Location = New-Object System.Drawing.Point(555, 17); $btnRemoveTags.Size = New-Object System.Drawing.Size(90, 25); $btnRemoveTags.Text = "Remove Tags"; $btnRemoveTags.BackColor = "#CD5C5C"; $btnRemoveTags.ForeColor = "White"


# SEPARATOR 1: Tags | Teams Actions (Shifted Right)
$sepAction1 = New-Object System.Windows.Forms.Label; $sepAction1.Location = New-Object System.Drawing.Point(650, 15); $sepAction1.Size = New-Object System.Drawing.Size(2, 30); $sepAction1.BorderStyle = "Fixed3D"

# Group 2: Teams Actions (Shifted Right)
$btnAssign = New-Object System.Windows.Forms.Button; $btnAssign.Location = New-Object System.Drawing.Point(665, 17); $btnAssign.Size = New-Object System.Drawing.Size(80, 25); $btnAssign.Text = "Assign"; $btnAssign.BackColor = "#5cb85c"; $btnAssign.ForeColor = "White"; $btnAssign.Enabled = $false
$btnUnassign = New-Object System.Windows.Forms.Button; $btnUnassign.Location = New-Object System.Drawing.Point(755, 17); $btnUnassign.Size = New-Object System.Drawing.Size(90, 25); $btnUnassign.Text = "Unassign"; $btnUnassign.BackColor = "#F0AD4E"; $btnUnassign.ForeColor = "White"

# SEPARATOR 2: Teams Actions | Number Actions (Shifted Right)
$sepAction2 = New-Object System.Windows.Forms.Label; $sepAction2.Location = New-Object System.Drawing.Point(860, 15); $sepAction2.Size = New-Object System.Drawing.Size(2, 30); $sepAction2.BorderStyle = "Fixed3D"

# Group 3: Number Actions (Remove) (Shifted Right)
$btnRemove = New-Object System.Windows.Forms.Button; $btnRemove.Location = New-Object System.Drawing.Point(875, 17); $btnRemove.Size = New-Object System.Drawing.Size(90, 25); $btnRemove.Text = "Remove"; $btnRemove.BackColor = "#D9534F"; $btnRemove.ForeColor = "White"

# SEPARATOR 3: Number Actions | Orange Actions (Shifted Right)
$sepAction3 = New-Object System.Windows.Forms.Label; $sepAction3.Location = New-Object System.Drawing.Point(980, 15); $sepAction3.Size = New-Object System.Drawing.Size(2, 30); $sepAction3.BorderStyle = "Fixed3D"

# Group 4: Orange Actions (Shifted Right)
$btnReleaseOC = New-Object System.Windows.Forms.Button; $btnReleaseOC.Location = New-Object System.Drawing.Point(995, 17); $btnReleaseOC.Size = New-Object System.Drawing.Size(110, 25); $btnReleaseOC.Text = "Release from OC"; $btnReleaseOC.BackColor = "#C71585"; $btnReleaseOC.ForeColor = "White"
$btnPublishOC = New-Object System.Windows.Forms.Button; $btnPublishOC.Location = New-Object System.Drawing.Point(1115, 17); $btnPublishOC.Size = New-Object System.Drawing.Size(110, 25); $btnPublishOC.Text = "Publish to OC"; $btnPublishOC.BackColor = "#008B8B"; $btnPublishOC.ForeColor = "White"

# NEW BUTTON: Manual Publish (Shifted Right)
$btnManualPublish = New-Object System.Windows.Forms.Button; $btnManualPublish.Location = New-Object System.Drawing.Point(1235, 17); $btnManualPublish.Size = New-Object System.Drawing.Size(110, 25); $btnManualPublish.Text = "Manual Publish"; $btnManualPublish.BackColor = "#4682B4"; $btnManualPublish.ForeColor = "White"

$global:cbDebug = New-Object System.Windows.Forms.CheckBox; $global:cbDebug.Location = New-Object System.Drawing.Point(1370, 19); $global:cbDebug.Size = New-Object System.Drawing.Size(80, 20); $global:cbDebug.Text = "Debug"

$grpTag.Controls.AddRange(@($lblTagInput, $global:cmbTag, $cbBlacklist, $cbReserved, $cbPremium, $btnApplyTag, $btnRemoveTags, $sepAction1, $btnAssign, $btnUnassign, $sepAction2, $btnRemove, $sepAction3, $btnReleaseOC, $btnPublishOC, $btnManualPublish, $global:cbDebug))
#endregion

#region 9. Logic - General UI
$btnSelectCols.Add_Click({
    if ($dataGridView.DataSource -eq $null) { [System.Windows.Forms.MessageBox]::Show("Please load data first.", "Info"); return }
    $frmCols = New-Object System.Windows.Forms.Form; $frmCols.Text = "Select Columns"; $frmCols.Size = New-Object System.Drawing.Size(300, 500); $frmCols.StartPosition = "CenterParent"; $frmCols.FormBorderStyle = "FixedDialog"
    $clb = New-Object System.Windows.Forms.CheckedListBox; $clb.Location = New-Object System.Drawing.Point(10, 10); $clb.Size = New-Object System.Drawing.Size(260, 400); $clb.CheckOnClick = $true
    foreach ($col in $dataGridView.Columns) { $state = if ($col.Visible) { [System.Windows.Forms.CheckState]::Checked } else { [System.Windows.Forms.CheckState]::Unchecked }; [void]$clb.Items.Add($col.Name, $state) }
    $btnOkCols = New-Object System.Windows.Forms.Button; $btnOkCols.Text = "OK"; $btnOkCols.DialogResult = "OK"; $btnOkCols.Location = New-Object System.Drawing.Point(100, 420); $frmCols.Controls.Add($clb); $frmCols.Controls.Add($btnOkCols); $frmCols.AcceptButton = $btnOkCols
    if ($frmCols.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { $global:form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor; $dataGridView.SuspendLayout(); for ($i = 0; $i -lt $clb.Items.Count; $i++) { $colName = $clb.Items[$i]; $isVisible = $clb.GetItemChecked($i); if ($colName -eq "TelephoneNumber") { $isVisible = $true }; $dataGridView.Columns[$colName].Visible = $isVisible }; $dataGridView.ResumeLayout(); $global:form.Cursor = [System.Windows.Forms.Cursors]::Default }; $frmCols.Dispose()
})
#endregion
