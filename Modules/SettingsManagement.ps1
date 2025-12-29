# =============================================================================
# SETTINGS MANAGEMENT MODULE - XML Settings Import/Export
# =============================================================================

function Update-TagUI {
    # UPDATED: Only refreshes the Dropdown in Actions from settings.
    # It DOES NOT touch the Stats Grid anymore (Stats are now data-driven).
    if ($global:cmbTag) {
        $global:cmbTag.Items.Clear()
        foreach ($tag in $global:allowedTags) { [void]$global:cmbTag.Items.Add($tag) }
    }
}

function Import-SettingsXml {
    param([string]$Path)

    if (-not (Test-Path $Path)) {
        Write-Log "Settings XML not found at: $Path"
        return
    }

    try {
        [xml]$xml = Get-Content $Path
        $settings = $xml.Settings

        # Populate Input Fields
        if ($settings.OrangeCustomerID) { $global:txtOrangeCust.Text = $settings.OrangeCustomerID }
        if ($settings.OrangeAuthHeader) { $global:txtOrangeAuth.Text = $settings.OrangeAuthHeader }
        if ($settings.OrangeApiKey) { $global:txtOrangeKey.Text = $settings.OrangeApiKey }
        if ($settings.Proxy) { $global:txtProxy.Text = $settings.Proxy }

        # Populate Threshold
        if ($settings.LowStockAlertThreshold) {
            $global:lowStockThreshold = [int]$settings.LowStockAlertThreshold
        }

        # Populate Tag List
        if ($settings.SelectTagList) {
            $rawTags = $settings.SelectTagList
            if ($rawTags -match ",") {
                $global:allowedTags = ($rawTags -split ",").Trim() | Sort-Object -Unique
            } else {
                $global:allowedTags = @($rawTags.Trim())
            }
            Update-TagUI
        }

        $global:settingsXmlPath = $Path
        Write-Log "Settings loaded from: $Path"

        # Update GroupBox Text to show loaded file
        if ($grpConfig) { $grpConfig.Text = "Orange API Configuration (Loaded: $(Split-Path $Path -Leaf))" }

    } catch {
        Write-Log "Error loading XML: $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show("Failed to load XML settings.`n$($_.Exception.Message)", "Error")
    }
}

function Export-SettingsXml {
    # Saves current UI values to the currently selected XML file
    if ([string]::IsNullOrWhiteSpace($global:settingsXmlPath)) {
        [System.Windows.Forms.MessageBox]::Show("No XML file is currently selected. Please select a file first.", "Warning")
        return
    }

    try {
        [xml]$xml = Get-Content $global:settingsXmlPath

        # Update values from UI
        $xml.Settings.OrangeCustomerID = $global:txtOrangeCust.Text
        $xml.Settings.OrangeAuthHeader = $global:txtOrangeAuth.Text
        $xml.Settings.OrangeApiKey = $global:txtOrangeKey.Text
        $xml.Settings.Proxy = $global:txtProxy.Text

        # Update non-UI values (preserve what is in memory)
        $xml.Settings.LowStockAlertThreshold = $global:lowStockThreshold.ToString()
        $xml.Settings.SelectTagList = ($global:allowedTags -join ",")

        $xml.Save($global:settingsXmlPath)
        Write-Log "Settings saved to: $global:settingsXmlPath"
        [System.Windows.Forms.MessageBox]::Show("Settings saved successfully!", "Success")

    } catch {
        Write-Log "Error saving XML: $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show("Failed to save settings.`n$($_.Exception.Message)", "Error")
    }
}
