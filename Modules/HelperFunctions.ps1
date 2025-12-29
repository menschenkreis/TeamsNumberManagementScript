# =============================================================================
# HELPER FUNCTIONS MODULE - UI Helpers, Input Dialogs, and Utility Functions
# =============================================================================

function Write-Log($message) {
    $timestamp = Get-Date -Format "HH:mm:ss"
    $fullMsg = "[$timestamp] $message"

    if ($global:txtLog) {
        $global:txtLog.AppendText("$fullMsg`r`n")
        $global:txtLog.ScrollToCaret()
        [System.Windows.Forms.Application]::DoEvents()
    }

    if ($global:logFilePath) {
        try {
            Add-Content -Path $global:logFilePath -Value $fullMsg -ErrorAction SilentlyContinue
        } catch {}
    }
}

function Write-Debug($message) {
    if ($global:cbDebug.Checked) {
        Write-Log "[DEBUG] $message"
    }
}

# --- HELPER: CHECK BLACKLIST ---
function Test-IsBlacklisted {
    param([System.Windows.Forms.DataGridViewRow]$Row)
    $tagVal = [string]$Row.Cells["Tag"].Value
    if ($tagVal -match "Blacklist") { return $true }
    return $false
}

function Update-ProgressUI {
    param(
        [int]$Current,
        [int]$Total,
        [string]$Activity = "Processing"
    )

    if ($progressBar -and $Total -gt 0) {
        $pct = [Math]::Min(100, [int](($Current / $Total) * 100))
        $progressBar.Value = $pct
        $global:form.Text = "Teams Phone Manager v56.2 - $Activity ($Current / $Total)"
    }
    [System.Windows.Forms.Application]::DoEvents()
}

function Reset-ProgressUI {
    if ($progressBar) { $progressBar.Value = 0 }
    if ($global:form) { $global:form.Text = "Teams Phone Manager v56.2" }
    [System.Windows.Forms.Application]::DoEvents()
}

function Update-Stats {
    $total = if ($global:masterDataTable) { $global:masterDataTable.Rows.Count } else { 0 }
    $disp = if ($dataGridView) { $dataGridView.Rows.Count } else { 0 }
    $sel = if ($dataGridView) { $dataGridView.SelectedRows.Count } else { 0 }

    if ($global:lblValTotal) { $global:lblValTotal.Text = "$total" }
    if ($global:lblValDisp) { $global:lblValDisp.Text = "$disp" }
    if ($global:lblValSel) { $global:lblValSel.Text = "$sel" }
}

function Get-SimpleInput {
    param([string]$Title = "Input", [string]$Prompt = "Please enter value:")
    $f = New-Object System.Windows.Forms.Form; $f.Width = 400; $f.Height = 180; $f.Text = $Title; $f.StartPosition = "CenterParent"; $f.FormBorderStyle = "FixedDialog"; $f.MaximizeBox = $false; $f.MinimizeBox = $false
    $lbl = New-Object System.Windows.Forms.Label; $lbl.Location = New-Object System.Drawing.Point(20, 20); $lbl.Size = New-Object System.Drawing.Size(340, 20); $lbl.Text = $Prompt; $f.Controls.Add($lbl)
    $txt = New-Object System.Windows.Forms.TextBox; $txt.Location = New-Object System.Drawing.Point(20, 50); $txt.Size = New-Object System.Drawing.Size(340, 20); $f.Controls.Add($txt)
    $btnOk = New-Object System.Windows.Forms.Button; $btnOk.Text = "OK"; $btnOk.DialogResult = "OK"; $btnOk.Location = New-Object System.Drawing.Point(200, 90); $f.Controls.Add($btnOk)
    $btnCancel = New-Object System.Windows.Forms.Button; $btnCancel.Text = "Cancel"; $btnCancel.DialogResult = "Cancel"; $btnCancel.Location = New-Object System.Drawing.Point(280, 90); $f.Controls.Add($btnCancel)
    $f.AcceptButton = $btnOk; $f.CancelButton = $btnCancel
    if ($f.ShowDialog() -eq "OK") { return $txt.Text } return $null
}

function Get-SelectionInput {
    param(
        [string]$Title = "Select Option",
        [string]$Prompt = "Please select a value:",
        [string[]]$Options
    )
    $f = New-Object System.Windows.Forms.Form
    $f.Width = 400
    $f.Height = 200
    $f.Text = $Title
    $f.StartPosition = "CenterParent"
    $f.FormBorderStyle = "FixedDialog"
    $f.MaximizeBox = $false
    $f.MinimizeBox = $false

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Location = New-Object System.Drawing.Point(20, 20)
    $lbl.Size = New-Object System.Drawing.Size(340, 20)
    $lbl.Text = $Prompt
    $f.Controls.Add($lbl)

    $cmb = New-Object System.Windows.Forms.ComboBox
    $cmb.Location = New-Object System.Drawing.Point(20, 50)
    $cmb.Size = New-Object System.Drawing.Size(340, 20)
    $cmb.DropDownStyle = "DropDownList"
    if ($Options) {
        foreach ($opt in $Options) { [void]$cmb.Items.Add($opt) }
        if ($cmb.Items.Count -gt 0) { $cmb.SelectedIndex = 0 }
    }
    $f.Controls.Add($cmb)

    $btnOk = New-Object System.Windows.Forms.Button
    $btnOk.Text = "OK"
    $btnOk.DialogResult = "OK"
    $btnOk.Location = New-Object System.Drawing.Point(200, 110)
    $f.Controls.Add($btnOk)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Cancel"
    $btnCancel.DialogResult = "Cancel"
    $btnCancel.Location = New-Object System.Drawing.Point(280, 110)
    $f.Controls.Add($btnCancel)

    $f.AcceptButton = $btnOk; $f.CancelButton = $btnCancel

    if ($f.ShowDialog() -eq "OK") { return $cmb.SelectedItem }
    return $null
}

function Get-ManualPublishInput {
    param([string[]]$Sites)

    $f = New-Object System.Windows.Forms.Form
    $f.Width = 450
    $f.Height = 250
    $f.Text = "Manual Publish to OC"
    $f.StartPosition = "CenterParent"
    $f.FormBorderStyle = "FixedDialog"
    $f.MaximizeBox = $false
    $f.MinimizeBox = $false

    # Start Number
    $lblStart = New-Object System.Windows.Forms.Label; $lblStart.Location = New-Object System.Drawing.Point(20, 20); $lblStart.Size = New-Object System.Drawing.Size(100, 20); $lblStart.Text = "Start Number:"
    $txtStart = New-Object System.Windows.Forms.TextBox; $txtStart.Location = New-Object System.Drawing.Point(130, 18); $txtStart.Size = New-Object System.Drawing.Size(280, 20)

    # End Number
    $lblEnd = New-Object System.Windows.Forms.Label; $lblEnd.Location = New-Object System.Drawing.Point(20, 50); $lblEnd.Size = New-Object System.Drawing.Size(100, 20); $lblEnd.Text = "End Number:"
    $txtEnd = New-Object System.Windows.Forms.TextBox; $txtEnd.Location = New-Object System.Drawing.Point(130, 48); $txtEnd.Size = New-Object System.Drawing.Size(280, 20)

    # Voice Site
    $lblSite = New-Object System.Windows.Forms.Label; $lblSite.Location = New-Object System.Drawing.Point(20, 80); $lblSite.Size = New-Object System.Drawing.Size(100, 20); $lblSite.Text = "Voice Site:"
    $cmbSite = New-Object System.Windows.Forms.ComboBox; $cmbSite.Location = New-Object System.Drawing.Point(130, 78); $cmbSite.Size = New-Object System.Drawing.Size(280, 20); $cmbSite.DropDownStyle = "DropDownList"
    if ($Sites) { foreach ($s in $Sites) { [void]$cmbSite.Items.Add($s) } }
    if ($cmbSite.Items.Count -gt 0) { $cmbSite.SelectedIndex = 0 }

    # Warning
    $lblWarn = New-Object System.Windows.Forms.Label; $lblWarn.Location = New-Object System.Drawing.Point(20, 110); $lblWarn.Size = New-Object System.Drawing.Size(390, 40); $lblWarn.Text = "Note: Spaces and leading '+' signs will be automatically removed."
    $lblWarn.ForeColor = "Gray"

    $f.Controls.AddRange(@($lblStart, $txtStart, $lblEnd, $txtEnd, $lblSite, $cmbSite, $lblWarn))

    # Buttons
    $btnOk = New-Object System.Windows.Forms.Button; $btnOk.Text = "Publish"; $btnOk.DialogResult = "OK"; $btnOk.Location = New-Object System.Drawing.Point(230, 160); $f.Controls.Add($btnOk)
    $btnCancel = New-Object System.Windows.Forms.Button; $btnCancel.Text = "Cancel"; $btnCancel.DialogResult = "Cancel"; $btnCancel.Location = New-Object System.Drawing.Point(310, 160); $f.Controls.Add($btnCancel)

    $f.AcceptButton = $btnOk; $f.CancelButton = $btnCancel

    if ($f.ShowDialog() -eq "OK") {
        return @{
            StartNumber = $txtStart.Text
            EndNumber   = $txtEnd.Text
            VoiceSite   = $cmbSite.SelectedItem
        }
    }
    return $null
}

function Export-SelectedToCSV {
    $sel = $dataGridView.SelectedRows
    if ($sel.Count -eq 0) { [System.Windows.Forms.MessageBox]::Show("Please select rows to export.", "Info"); return }
    $sfd = New-Object System.Windows.Forms.SaveFileDialog; $sfd.Filter = "CSV Files (*.csv)|*.csv"; $sfd.Title = "Save Export As"; $sfd.FileName = "TeamsPhoneExport_$(Get-Date -Format 'yyyyMMdd_HHmm').csv"
    if ($sfd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $path = $sfd.FileName; Write-Log "Exporting $($sel.Count) rows to CSV..."; $global:form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        try {
            $exportList = New-Object System.Collections.ArrayList
            foreach ($row in $sel) {
                $obj = [Ordered]@{}; foreach ($colName in $global:tableColumns) { $val = $row.Cells[$colName].Value; if ($null -eq $val) { $val = "" }; $obj[$colName] = [string]$val }; [void]$exportList.Add([PSCustomObject]$obj)
            }
            $exportList | Export-Csv -Path $path -NoTypeInformation -Delimiter "," -Encoding UTF8; Write-Log "Export saved to: $path"; [System.Windows.Forms.MessageBox]::Show("Export successful!", "Done")
        } catch { Write-Log "Export Error: $($_.Exception.Message)"; [System.Windows.Forms.MessageBox]::Show("Export Failed: $($_.Exception.Message)", "Error") } finally { $global:form.Cursor = [System.Windows.Forms.Cursors]::Default }
    }
}
