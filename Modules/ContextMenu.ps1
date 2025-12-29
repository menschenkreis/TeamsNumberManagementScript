# =============================================================================
# CONTEXT MENU MODULE - Right-Click Context Menu Actions
# =============================================================================
# NOTE: This module contains event handlers for context menu items
# Must be loaded AFTER UIComponents.ps1

#region Context Menu Actions

# --- CONTEXT MENU ACTION: Refresh Orange ---
$miRefreshOrange.Add_Click({
    # GUARD: Key Check
    if ([string]::IsNullOrWhiteSpace($global:txtOrangeKey.Text)) {
        [System.Windows.Forms.MessageBox]::Show("Orange API Key is not populated. Feature disabled.", "Feature Disabled", "OK", "Warning"); return
    }

    $sel = $dataGridView.SelectedRows; if ($sel.Count -eq 0) { return }
    if ([string]::IsNullOrWhiteSpace($global:txtOrangeAuth.Text) -or [string]::IsNullOrWhiteSpace($global:txtOrangeCust.Text)) {
        [System.Windows.Forms.MessageBox]::Show("Orange configuration fields missing.", "Error"); return
    }

    Write-Log "Refreshing Orange info for $($sel.Count) number(s)..."
    $global:form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor

    $counter = 0
    foreach ($row in $sel) {
        $counter++
        Update-ProgressUI -Current $counter -Total $sel.Count -Activity "Refreshing Orange Info"
        $ph = $row.Cells["TelephoneNumber"].Value
        try {
            $obj = Get-SingleOrangeNumber -PhoneNumber $ph
            if ($obj) {
                $siteName = ""; $siteId = ""; $vs = $obj.voiceSite
                if ($null -ne $vs) { $siteId = $vs.voiceSiteId; $siteName = $vs.technicalSiteName; if ([string]::IsNullOrWhiteSpace($siteId) -and $vs -is [System.Collections.IDictionary]) { $siteId = $vs['voiceSiteId']; $siteName = $vs['technicalSiteName'] } }
                $row.Cells["OrangeSite"].Value = $siteName; $row.Cells["OrangeSiteId"].Value = $siteId; $row.Cells["OrangeStatus"].Value = $obj.status; $row.Cells["OrangeUsage"].Value = $obj.usage
                $global:orangeHistoryMap[$ph] = $obj.history
                Write-Log "Updated $ph : Status=$($obj.status)"
            } else { Write-Log "Warning: $ph not found in Orange or error occurred." }
        } catch { Write-Log "Error refreshing $ph : $($_.Exception.Message)" }
    }
    $global:form.Cursor = [System.Windows.Forms.Cursors]::Default
    Reset-ProgressUI
})

# --- CONTEXT MENU ACTION: Refresh Teams Info ---
$miRefreshTeams.Add_Click({
    $sel = $dataGridView.SelectedRows; if ($sel.Count -eq 0) { [System.Windows.Forms.MessageBox]::Show("Please select rows to refresh.", "Info"); return }

    $global:form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    Write-Log "Refreshing Teams info for $($sel.Count) number(s)..."

    $counter = 0
    foreach ($row in $sel) {
        $counter++
        Update-ProgressUI -Current $counter -Total $sel.Count -Activity "Refreshing Teams Info"
        $ph = $row.Cells["TelephoneNumber"].Value
        try {
            Write-Debug "Executing: Get-CsPhoneNumberAssignment -TelephoneNumber $ph"
            $numData = Get-CsPhoneNumberAssignment -TelephoneNumber $ph -ErrorAction Stop

            if ($numData) {
                $row.Cells["ActivationState"].Value = $numData.ActivationState; $row.Cells["NumberType"].Value = $numData.NumberType; $row.Cells["City"].Value = $numData.City
                $row.Cells["IsoCountryCode"].Value = $numData.IsoCountryCode; $row.Cells["IsoSubdivision"].Value = $numData.IsoSubdivision; $row.Cells["NumberSource"].Value = $numData.NumberSource
                $row.Cells["Tag"].Value = if ($numData.Tag) { $numData.Tag -join ", " } else { "" }

                $userId = $numData.AssignedPstnTargetId
                if (-not [string]::IsNullOrWhiteSpace($userId)) {
                    try {
                        Write-Debug "Executing: Get-CsOnlineUser -Identity $userId"
                        $userData = Get-CsOnlineUser -Identity $userId -ErrorAction Stop
                        $row.Cells["UserPrincipalName"].Value = $userData.UserPrincipalName; $row.Cells["DisplayName"].Value = $userData.DisplayName
                        $row.Cells["OnlineVoiceRoutingPolicy"].Value = $userData.OnlineVoiceRoutingPolicy; $row.Cells["EnterpriseVoiceEnabled"].Value = $userData.EnterpriseVoiceEnabled

                        # UPDATED: Map new optional user columns on refresh
                        if ($row.DataGridView.Columns.Contains("AccountEnabled")) { $row.Cells["AccountEnabled"].Value = $userData.AccountEnabled }
                        if ($row.DataGridView.Columns.Contains("PreferredDataLocation")) { $row.Cells["PreferredDataLocation"].Value = $userData.PreferredDataLocation }
                        if ($row.DataGridView.Columns.Contains("UsageLocation")) { $row.Cells["UsageLocation"].Value = $userData.UsageLocation }
                        if ($row.DataGridView.Columns.Contains("TeamsMeetingPolicy")) { $row.Cells["TeamsMeetingPolicy"].Value = $userData.TeamsMeetingPolicy }

                        # UPDATED: Handle array FeatureTypes
                        if ($row.DataGridView.Columns.Contains("FeatureTypes")) {
                            if ($userData.FeatureTypes -and $userData.FeatureTypes -is [Array]) {
                                $row.Cells["FeatureTypes"].Value = $userData.FeatureTypes -join ", "
                            } else {
                                $row.Cells["FeatureTypes"].Value = $userData.FeatureTypes
                            }
                        }

                        Write-Log "  > Refreshed $ph (Assigned: $($userData.DisplayName))"
                    } catch { Write-Log "  > $ph is assigned ($userId), but user fetch failed: $($_.Exception.Message)" }
                } else {
                    $row.Cells["UserPrincipalName"].Value = ""; $row.Cells["DisplayName"].Value = ""; $row.Cells["OnlineVoiceRoutingPolicy"].Value = ""; $row.Cells["EnterpriseVoiceEnabled"].Value = ""
                    # Clear new cols
                    if ($row.DataGridView.Columns.Contains("AccountEnabled")) { $row.Cells["AccountEnabled"].Value = "" }
                    if ($row.DataGridView.Columns.Contains("PreferredDataLocation")) { $row.Cells["PreferredDataLocation"].Value = "" }
                    if ($row.DataGridView.Columns.Contains("UsageLocation")) { $row.Cells["UsageLocation"].Value = "" }
                    if ($row.DataGridView.Columns.Contains("TeamsMeetingPolicy")) { $row.Cells["TeamsMeetingPolicy"].Value = "" }
                    if ($row.DataGridView.Columns.Contains("FeatureTypes")) { $row.Cells["FeatureTypes"].Value = "" }
                    Write-Log "  > Refreshed $ph (Unassigned)"
                }
            } else { Write-Log "  > Warning: Number $ph returned no data from Teams." }
        } catch { Write-Log "  > Error refreshing $ph : $($_.Exception.Message)" }
    }
    Update-TagStatistics
    $global:form.Cursor = [System.Windows.Forms.Cursors]::Default
    Reset-ProgressUI
})

# --- CONTEXT MENU ACTION: Force Publish to OC ---
$miForcePublish.Add_Click({
    # GUARD: Key Check
    if ([string]::IsNullOrWhiteSpace($global:txtOrangeKey.Text)) {
        [System.Windows.Forms.MessageBox]::Show("Orange API Key is not populated. Feature disabled.", "Feature Disabled", "OK", "Warning"); return
    }

    $sel = $dataGridView.SelectedRows; if ($sel.Count -eq 0) { [System.Windows.Forms.MessageBox]::Show("Please select rows to publish.", "Info"); return }

    # BLACKLIST CHECK
    $rowsToProcess = New-Object System.Collections.ArrayList
    foreach ($row in $sel) {
        if (Test-IsBlacklisted $row) { Write-Log "Blocked: $($row.Cells['TelephoneNumber'].Value) is Blacklisted."; continue }
        [void]$rowsToProcess.Add($row)
    }
    if ($rowsToProcess.Count -eq 0) { [System.Windows.Forms.MessageBox]::Show("Operation blocked for all selected numbers (Blacklisted).", "Blocked", "OK", "Warning"); return }

    if ($global:masterDataTable -eq $null -or $global:masterDataTable.Rows.Count -eq 0) { [System.Windows.Forms.MessageBox]::Show("No data loaded. Cannot find available Orange Sites.", "Error"); return }

    $siteMap = @{}
    foreach ($row in $global:masterDataTable.Rows) {
        $sId = $row["OrangeSiteId"]; $sName = $row["OrangeSite"]
        if (-not [string]::IsNullOrWhiteSpace($sId) -and -not [string]::IsNullOrWhiteSpace($sName)) { $display = "$sName ($sId)"; if (-not $siteMap.ContainsKey($display)) { $siteMap[$display] = $sId } }
    }

    if ($siteMap.Count -eq 0) { [System.Windows.Forms.MessageBox]::Show("No known Orange Sites found in the current dataset. Please ensure data is synced.", "Warning"); return }

    $sortedOptions = $siteMap.Keys | Sort-Object
    $selectedOption = Get-SelectionInput -Title "Force Publish to OC" -Prompt "Select the Target Orange Site:" -Options $sortedOptions

    if ($selectedOption) {
        $targetSiteId = $siteMap[$selectedOption]; $targetSiteName = ($selectedOption -split " \(")[0]
        if ([System.Windows.Forms.MessageBox]::Show("Force publish $($rowsToProcess.Count) numbers to Site: $targetSiteName ($targetSiteId)?", "Confirm", "YesNo", "Warning") -eq "Yes") {
            $global:form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
            foreach ($r in $rowsToProcess) { $r.Cells["OrangeSiteId"].Value = $targetSiteId; $r.Cells["OrangeSite"].Value = $targetSiteName }

            $successfulRows = Publish-OrangeNumbers -RowObjects $rowsToProcess -UsageType "CallingUserAssignment"
            if ($successfulRows.Count -gt 0) {
                Update-SelectedRows -Rows $successfulRows -NewOrangeStatus "published"
                [System.Windows.Forms.MessageBox]::Show("Publish command sent successfully.", "Success")
            } else { Write-Log "No rows were updated in grid because API call(s) failed." }

            $global:form.Cursor = [System.Windows.Forms.Cursors]::Default
            Reset-ProgressUI
        }
    }
})

# --- CONTEXT MENU ACTION: Change Policy ---
$miChangePolicy.Add_Click({
    $sel = $dataGridView.SelectedRows; if ($sel.Count -ne 1) { [System.Windows.Forms.MessageBox]::Show("Please select exactly one row.", "Info"); return }
    $row = $sel[0]; $upn = $row.Cells["UserPrincipalName"].Value; $currentPolicy = $row.Cells["OnlineVoiceRoutingPolicy"].Value

    # BLACKLIST CHECK
    if (Test-IsBlacklisted $row) { [System.Windows.Forms.MessageBox]::Show("Operation blocked: This number is Blacklisted.", "Blocked", "OK", "Warning"); return }

    if ([string]::IsNullOrWhiteSpace($upn)) { [System.Windows.Forms.MessageBox]::Show("This number is not assigned to a user.", "Warning"); return }
    if ($global:voiceRoutingPolicies.Count -eq 0) { [System.Windows.Forms.MessageBox]::Show("No policies loaded. Please run 'Get Data' again or check connections.", "Error"); return }

    $selectedPolicy = Get-SelectionInput -Title "Change Voice Policy" -Prompt "Select new policy for $upn (Current: $currentPolicy)" -Options $global:voiceRoutingPolicies

    if ($selectedPolicy) {
        try {
            $global:form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
            Write-Log "Granting policy '$selectedPolicy' to $upn..."
            Write-Debug "Executing: Grant-CsOnlineVoiceRoutingPolicy -Identity $upn -PolicyName $selectedPolicy"
            Grant-CsOnlineVoiceRoutingPolicy -Identity $upn -PolicyName $selectedPolicy -ErrorAction Stop
            Write-Log "Success. Updating grid..."
            $row.Cells["OnlineVoiceRoutingPolicy"].Value = $selectedPolicy
            [System.Windows.Forms.MessageBox]::Show("Policy updated successfully.", "Success")
        } catch {
            Write-Log "Policy Grant Failed: $($_.Exception.Message)"
            [System.Windows.Forms.MessageBox]::Show("Failed to grant policy: $($_.Exception.Message)", "Error")
        } finally { $global:form.Cursor = [System.Windows.Forms.Cursors]::Default }
    }
})

# --- CONTEXT MENU ACTION: Grant Meeting Policy ---
$miGrantMeetingPolicy.Add_Click({
    $sel = $dataGridView.SelectedRows; if ($sel.Count -ne 1) { [System.Windows.Forms.MessageBox]::Show("Please select exactly one row.", "Info"); return }
    $row = $sel[0]; $upn = $row.Cells["UserPrincipalName"].Value; $currentPolicy = $row.Cells["TeamsMeetingPolicy"].Value

    # BLACKLIST CHECK
    if (Test-IsBlacklisted $row) { [System.Windows.Forms.MessageBox]::Show("Operation blocked: This number is Blacklisted.", "Blocked", "OK", "Warning"); return }

    if ([string]::IsNullOrWhiteSpace($upn)) { [System.Windows.Forms.MessageBox]::Show("This number is not assigned to a user.", "Warning"); return }
    if ($global:teamsMeetingPolicies.Count -eq 0) { [System.Windows.Forms.MessageBox]::Show("No meeting policies loaded.", "Error"); return }

    $selectedPolicy = Get-SelectionInput -Title "Grant Meeting Policy" -Prompt "Select new meeting policy for $upn (Current: $currentPolicy)" -Options $global:teamsMeetingPolicies

    if ($selectedPolicy) {
        try {
            $global:form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
            Write-Log "Granting meeting policy '$selectedPolicy' to $upn..."
            Write-Debug "Executing: Grant-CsTeamsMeetingPolicy -Identity $upn -PolicyName '$selectedPolicy'"
            Grant-CsTeamsMeetingPolicy -Identity $upn -PolicyName $selectedPolicy -ErrorAction Stop
            Write-Log "Success. Updating grid..."
            $row.Cells["TeamsMeetingPolicy"].Value = $selectedPolicy
            [System.Windows.Forms.MessageBox]::Show("Meeting Policy updated successfully.", "Success")
        } catch {
            Write-Log "Meeting Policy Grant Failed: $($_.Exception.Message)"
            [System.Windows.Forms.MessageBox]::Show("Failed to grant policy: $($_.Exception.Message)", "Error")
        } finally { $global:form.Cursor = [System.Windows.Forms.Cursors]::Default }
    }
})

# --- CONTEXT MENU ACTION: Enable EV ---
$miEnableEV.Add_Click({
    $sel = $dataGridView.SelectedRows
    if ($sel.Count -eq 0) { [System.Windows.Forms.MessageBox]::Show("Please select rows.", "Info"); return }

    # BLACKLIST CHECK
    $rowsToProcess = New-Object System.Collections.ArrayList
    foreach ($row in $sel) {
        if (Test-IsBlacklisted $row) { Write-Log "Blocked: $($row.Cells['TelephoneNumber'].Value) is Blacklisted."; continue }
        [void]$rowsToProcess.Add($row)
    }
    if ($rowsToProcess.Count -eq 0) { [System.Windows.Forms.MessageBox]::Show("Operation blocked for all selected numbers (Blacklisted).", "Blocked", "OK", "Warning"); return }

    if ([System.Windows.Forms.MessageBox]::Show("Enable Enterprise Voice for $($rowsToProcess.Count) user(s)?", "Confirm", "YesNo", "Question") -ne "Yes") { return }

    $global:form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $counter = 0

    foreach ($row in $rowsToProcess) {
        $counter++
        Update-ProgressUI -Current $counter -Total $rowsToProcess.Count -Activity "Enabling EV"
        $upn = $row.Cells["UserPrincipalName"].Value

        if (-not [string]::IsNullOrWhiteSpace($upn)) {
            try {
                Write-Log "Enabling Enterprise Voice for $upn..."
                Write-Debug "Executing: Set-CsPhoneNumberAssignment -Identity $upn -EnterpriseVoiceEnabled `$true"
                Set-CsPhoneNumberAssignment -Identity $upn -EnterpriseVoiceEnabled $true -ErrorAction Stop
                Write-Log "Success."
                $row.Cells["EnterpriseVoiceEnabled"].Value = $true
            } catch {
                Write-Log "Failed to enable EV for $upn : $($_.Exception.Message)"
            }
        } else {
            Write-Log "Skipped row (No User assigned)."
        }
    }
    $global:form.Cursor = [System.Windows.Forms.Cursors]::Default
    Reset-ProgressUI
    [System.Windows.Forms.MessageBox]::Show("Process Complete.", "Done")
})

#endregion
