# =============================================================================
# ACTIONS MODULE - Main Action Button Event Handlers
# =============================================================================
# NOTE: This module contains event handlers for action buttons
# Must be loaded AFTER UIComponents.ps1

#region 12. Logic - Main Actions (Assign/Release/Publish)
$btnAssign.Add_Click({
    $sel = $dataGridView.SelectedRows; if ($sel.Count -ne 1) { return }; $row = $sel[0]; $ph = $row.Cells["TelephoneNumber"].Value; $type = $row.Cells["NumberType"].Value
    
    # BLACKLIST CHECK
    if (Test-IsBlacklisted $row) { [System.Windows.Forms.MessageBox]::Show("Operation blocked: This number is Blacklisted.", "Blocked", "OK", "Warning"); return }

    # 1. Get Input (UPN or SAM)
    $inputUser = Get-SimpleInput -Title "Assign Number" -Prompt "Enter User UPN or SamAccountName for $ph :"
    if ([string]::IsNullOrWhiteSpace($inputUser)) { return }

    $targetUpn = $null

    # 2. Check if it looks like a UPN (Email format)
    if ($inputUser -match '^\w+([-+.!'']\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*$') {
        $targetUpn = $inputUser
    } else {
        # 3. Not a UPN, assume SamAccountName and lookup
        try {
            $adUser = Get-ADUser -Filter "SamAccountName -eq '$inputUser'" -Properties UserPrincipalName -ErrorAction Stop
            if ($adUser) {
                $targetUpn = $adUser.UserPrincipalName
                Write-Log "Resolved SamAccountName '$inputUser' to UPN '$targetUpn'"
            } else {
                 [System.Windows.Forms.MessageBox]::Show("Could not find AD User with SamAccountName: $inputUser", "User Not Found", "OK", "Error")
                 return
            }
        } catch {
             [System.Windows.Forms.MessageBox]::Show("Failed to lookup SamAccountName. Ensure ActiveDirectory module is available.`nError: $($_.Exception.Message)", "Lookup Error", "OK", "Error")
             return
        }
    }

    $upn = $targetUpn

    $userObj = $null; foreach ($u in $global:teamsUsersMap.Values) { if ($u.UserPrincipalName -eq $upn) { $userObj = $u; break } }
    if ($null -eq $userObj) { [System.Windows.Forms.MessageBox]::Show("User '$upn' not found in downloaded index.", "Error", "OK", "Error"); return }
    try {
        $global:form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor; Write-Log "Assigning $ph to $upn ($type)..."; 
        Write-Debug "Executing: Set-CsPhoneNumberAssignment -Identity $upn -PhoneNumber $ph -PhoneNumberType $type"
        Set-CsPhoneNumberAssignment -Identity $upn -PhoneNumber $ph -PhoneNumberType $type -ErrorAction Stop; Write-Log "Teams Assignment Successful."
        $row.Cells["UserPrincipalName"].Value = $userObj.UserPrincipalName; $row.Cells["DisplayName"].Value = $userObj.DisplayName; $row.Cells["ActivationState"].Value = "Assigned"
        Write-Log "Syncing to On-Prem AD..."; try { $adUser = Get-ADUser -Filter "UserPrincipalName -eq '$upn'" -ErrorAction Stop; if ($adUser) { Set-ADUser -Identity $adUser -OfficePhone $ph -ErrorAction Stop; Write-Log "AD OfficePhone updated for $upn." } } catch { Write-Log "AD Sync Warning: $($_.Exception.Message)" }
        Update-TagStatistics # Update stats
    } catch { Write-Log "Assignment Failed: $($_.Exception.Message)"; [System.Windows.Forms.MessageBox]::Show("Assignment Failed: $($_.Exception.Message)", "Error") } finally { $global:form.Cursor = [System.Windows.Forms.Cursors]::Default }
})

$btnReleaseOC.Add_Click({
    # GUARD: Key Check
    if ([string]::IsNullOrWhiteSpace($global:txtOrangeKey.Text)) {
        [System.Windows.Forms.MessageBox]::Show("Orange API Key is not populated. Feature disabled.", "Feature Disabled", "OK", "Warning"); return
    }

    $selectedRows = $dataGridView.SelectedRows; if ($selectedRows.Count -eq 0) { Write-Log "Select rows."; return }
    
    # BLACKLIST CHECK
    $rowsToProcess = New-Object System.Collections.ArrayList; 
    foreach ($row in $selectedRows) {
        if (Test-IsBlacklisted $row) { Write-Log "Blocked: $($row.Cells['TelephoneNumber'].Value) is Blacklisted."; continue }
        [void]$rowsToProcess.Add($row) 
    }

    if ($rowsToProcess.Count -eq 0) { [System.Windows.Forms.MessageBox]::Show("Operation blocked for all selected numbers (Blacklisted).", "Blocked", "OK", "Warning"); return }

    if ([System.Windows.Forms.MessageBox]::Show("Release $($rowsToProcess.Count) numbers?", "Confirm", "YesNo", "Warning") -ne "Yes") { return }
    
    $global:form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    
    $counter = 0
    foreach ($row in $rowsToProcess) {
        $counter++
        Update-ProgressUI -Current $counter -Total $rowsToProcess.Count -Activity "Unassigning in Teams"
        
        $upn = $row.Cells["UserPrincipalName"].Value; $phone = $row.Cells["TelephoneNumber"].Value; $numType = $row.Cells["NumberType"].Value
        if (-not [string]::IsNullOrWhiteSpace($upn)) { 
            Write-Log "Unassigning $phone from $upn..."; 
            try { 
                Write-Debug "Executing: Remove-CsPhoneNumberAssignment -Identity $upn -PhoneNumber $phone -PhoneNumberType $numType"
                Remove-CsPhoneNumberAssignment -Identity $upn -PhoneNumber $phone -PhoneNumberType $numType -ErrorAction Stop; Write-Log "Unassigned $phone." 
                
                # NEW: Try to sync to AD (Clear OfficePhone)
                Write-Log "Syncing to On-Prem AD (Clearing OfficePhone)..."
                try {
                    $adUser = Get-ADUser -Filter "UserPrincipalName -eq '$upn'" -ErrorAction Stop
                    if ($adUser) {
                        Set-ADUser -Identity $adUser -Clear "telephoneNumber" -ErrorAction Stop
                        Write-Log "AD OfficePhone cleared for $upn."
                    }
                } catch {
                    Write-Log "AD Sync Warning: $($_.Exception.Message)"
                }

            } catch { $errMsg = $_.Exception.Message; if ($errMsg -match "on-premises Active Directory" -or $errMsg -match "synchronized to the cloud") { Write-Log "Detected On-Prem user. Clearing AD attributes..."; try { Set-ADUser -Identity $upn -Clear "msRTCSIP-Line", "telephoneNumber" -ErrorAction Stop; Write-Log "Success: AD attributes cleared." } catch { Write-Log "Failed to clear AD attributes: $($_.Exception.Message)." } } else { Write-Log "Failed unassign: ${phone}: $errMsg" }; continue } }
    }
    
    Update-ProgressUI -Current 0 -Total 100 -Activity "Releasing from OC"
    $successfulRows = Unpublish-OrangeNumbersBatch -RowObjects $rowsToProcess
    
    if ($successfulRows.Count -gt 0) { Update-SelectedRows -Rows $successfulRows -NewOrangeStatus "released" -ClearUserData } else { Write-Log "No rows were updated in grid because API call(s) failed." }
    Update-TagStatistics 
    $global:form.Cursor = [System.Windows.Forms.Cursors]::Default
    Reset-ProgressUI
})

$btnPublishOC.Add_Click({
    # GUARD: Key Check
    if ([string]::IsNullOrWhiteSpace($global:txtOrangeKey.Text)) {
        [System.Windows.Forms.MessageBox]::Show("Orange API Key is not populated. Feature disabled.", "Feature Disabled", "OK", "Warning"); return
    }

    $selectedRows = $dataGridView.SelectedRows; if ($selectedRows.Count -eq 0) { Write-Log "Select rows."; return }
    
    # BLACKLIST CHECK
    $rowsToProcess = New-Object System.Collections.ArrayList; 
    foreach ($row in $selectedRows) { 
        if (Test-IsBlacklisted $row) { Write-Log "Blocked: $($row.Cells['TelephoneNumber'].Value) is Blacklisted."; continue }
        [void]$rowsToProcess.Add($row) 
    }
    
    if ($rowsToProcess.Count -eq 0) { [System.Windows.Forms.MessageBox]::Show("Operation blocked for all selected numbers (Blacklisted).", "Blocked", "OK", "Warning"); return }

    if ([System.Windows.Forms.MessageBox]::Show("Publish $($rowsToProcess.Count) numbers?", "Confirm", "YesNo", "Information") -ne "Yes") { return }
    
    $global:form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $successfulRows = Publish-OrangeNumbers -RowObjects $rowsToProcess -UsageType "CallingUserAssignment" 
    if ($successfulRows.Count -gt 0) { Update-SelectedRows -Rows $successfulRows -NewOrangeStatus "published" } else { Write-Log "No rows were updated in grid because API call(s) failed." }
    $global:form.Cursor = [System.Windows.Forms.Cursors]::Default
    Reset-ProgressUI
})

# --- MANUAL PUBLISH LOGIC ---
$btnManualPublish.Add_Click({
    # GUARD: Key Check
    if ([string]::IsNullOrWhiteSpace($global:txtOrangeKey.Text)) {
        [System.Windows.Forms.MessageBox]::Show("Orange API Key is not populated. Feature disabled.", "Feature Disabled", "OK", "Warning"); return
    }

    if ($global:masterDataTable -eq $null -or $global:masterDataTable.Rows.Count -eq 0) { 
        [System.Windows.Forms.MessageBox]::Show("No data loaded. Please 'Get Data' first to populate Voice Sites.", "Error"); return 
    }

    # Extract Voice Sites from loaded data
    $siteMap = @{} 
    foreach ($row in $global:masterDataTable.Rows) {
        $sId = $row["OrangeSiteId"]; $sName = $row["OrangeSite"]
        if (-not [string]::IsNullOrWhiteSpace($sId) -and -not [string]::IsNullOrWhiteSpace($sName)) { 
            $display = "$sName ($sId)"
            if (-not $siteMap.ContainsKey($display)) { $siteMap[$display] = $sId } 
        }
    }
    
    if ($siteMap.Count -eq 0) { [System.Windows.Forms.MessageBox]::Show("No known Orange Sites found in the current dataset.", "Warning"); return }

    $sortedOptions = $siteMap.Keys | Sort-Object
    
    # Open Popup
    $inputData = Get-ManualPublishInput -Sites $sortedOptions
    
    if ($inputData) {
        # Normalization
        $startNum = $inputData.StartNumber.Replace(" ", "").Replace("+", "").Trim()
        $endNum   = $inputData.EndNumber.Replace(" ", "").Replace("+", "").Trim()
        $rawSite  = $inputData.VoiceSite
        $siteId   = $siteMap[$rawSite]
        
        if ([string]::IsNullOrWhiteSpace($startNum) -or [string]::IsNullOrWhiteSpace($endNum)) {
             [System.Windows.Forms.MessageBox]::Show("Start and End numbers cannot be empty.", "Error"); return
        }

        Write-Log "Manual Publish: $startNum to $endNum on Site: $rawSite"
        $global:form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        
        # Prepare API Call
        $customerCode = $global:txtOrangeCust.Text.Trim()
        $Service = "opc"
        
        $proxyParams = Get-ProxyParams
        
        try {
            $apiHeaders = Get-OrangeHeaders
            $numbersPayload = @( @{ usage = "CallingUserAssignment"; rangeStart = $startNum; rangeStop = $endNum } )
            $body = @{ voiceOffer = "btlvs"; voiceSiteId = $siteId; numbers = $numbersPayload } | ConvertTo-Json -Depth 5 -Compress
            $url = "https://api.orange.com/btalk/v1/customers/$customerCode/$Service/cloudvoicenumbers"
            
            Write-Log "Sending request to Orange API..."
            # SPLATTING FIX
            Invoke-WebRequest -Uri $url -Method POST -Headers $apiHeaders -Body $body -ContentType 'application/json' @proxyParams -UseBasicParsing | Out-Null
            
            Write-Log "Success: Manual Publish command sent."
            [System.Windows.Forms.MessageBox]::Show("Manual Publish Successful.`n`nPlease click 'Re-Sync Orange' to see updates in the grid.", "Success")
            
        } catch {
            $errBody = "Unknown"; if ($_.Exception.Response) { try { $reader = New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream()); $errBody = $reader.ReadToEnd() } catch {} }
            Write-Log "Manual Publish Failed: $($_.Exception.Message) | Resp: $errBody" 
            [System.Windows.Forms.MessageBox]::Show("Publish Failed: $($_.Exception.Message)", "Error")
        } finally {
             $global:form.Cursor = [System.Windows.Forms.Cursors]::Default
        }
    }
})

$dataGridView.Add_CellFormatting({ param($s,$e) if($e.ColumnIndex -ge 0 -and $dataGridView.Columns[$e.ColumnIndex].Name -match "^Orange") { $e.CellStyle.BackColor = [System.Drawing.Color]::Bisque } })
$dataGridView.Add_CellDoubleClick({ param($s,$e) if($e.RowIndex -ge 0){ $p=$dataGridView.Rows[$e.RowIndex].Cells["TelephoneNumber"].Value; Write-Log "-- $p --"; if($global:orangeHistoryMap[$p]){$global:orangeHistoryMap[$p] | ForEach-Object {Write-Log "$($_.date) | $($_.status)"}} } })
#endregion

#region 13. Logic - Tagging Operations
# SMART TAGGING
$btnApplyTag.Add_Click({ 
    $sel = $dataGridView.SelectedRows; if ($sel.Count -eq 0) { return }
    $desiredTags = @(); if ($global:cmbTag.SelectedItem) { $desiredTags += $global:cmbTag.SelectedItem }; if ($cbBlacklist.Checked) { $desiredTags += "Blacklist" }; if ($cbReserved.Checked)  { $desiredTags += "Reserved" }; if ($cbPremium.Checked) { $desiredTags += "Premium" }
    if ($desiredTags.Count -eq 0) { [System.Windows.Forms.MessageBox]::Show("Please select at least one tag option.", "Info"); return }
    
    Write-Log "Applying smart tagging to $($sel.Count) number(s)..."
    $counter = 0
    foreach ($r in $sel) {
        $counter++
        Update-ProgressUI -Current $counter -Total $sel.Count -Activity "Applying Tags"
        
        $ph = $r.Cells["TelephoneNumber"].Value; $currentTagStr = [string]$r.Cells["Tag"].Value
        if ([string]::IsNullOrWhiteSpace($currentTagStr)) { $currentTags = @() } else { $currentTags = $currentTagStr -split "," | ForEach-Object { $_.Trim() } }
        $tagsToRemove = @(); $tagsToAdd = @()
        $newLocationTag = $desiredTags | Where-Object { $global:allowedTags -contains $_ } | Select-Object -First 1
        if ($newLocationTag) { $oldLocationTag = $currentTags | Where-Object { $global:allowedTags -contains $_ } | Select-Object -First 1; if ($oldLocationTag -and $oldLocationTag -ne $newLocationTag) { $tagsToRemove += $oldLocationTag } }
        $specialTags = @("Blacklist", "Reserved", "Premium"); foreach ($st in $specialTags) { $isDesired = $desiredTags -contains $st; $isCurrent = $currentTags -contains $st; if ($isCurrent -and -not $isDesired) { $tagsToRemove += $st }; if ($isDesired -and -not $isCurrent) { $tagsToAdd += $st } }
        if ($newLocationTag) { if (-not ($currentTags -contains $newLocationTag)) { $tagsToAdd += $newLocationTag } }
        try { foreach ($t in $tagsToRemove) { Write-Log "Removing tag '$t' from $ph..."; Write-Debug "Executing: Remove-CsPhoneNumberTag -PhoneNumber $ph -Tag $t"; Remove-CsPhoneNumberTag -PhoneNumber $ph -Tag $t -ErrorAction Stop }; foreach ($t in $tagsToAdd) { Write-Log "Adding tag '$t' to $ph..."; Write-Debug "Executing: Set-CsPhoneNumberTag -PhoneNumber $ph -Tag $t"; Set-CsPhoneNumberTag -PhoneNumber $ph -Tag $t -ErrorAction Stop }; $finalTags = $currentTags | Where-Object { -not ($tagsToRemove -contains $_) }; $finalTags += $tagsToAdd; $r.Cells["Tag"].Value = ($finalTags | Sort-Object | Get-Unique) -join ", "; if ($tagsToAdd.Count -eq 0 -and $tagsToRemove.Count -eq 0) { Write-Log "No tag changes needed for $ph." } else { Write-Log "Tags updated for $ph." } } catch { Write-Log "Failed to update tags for ${ph}: $($_.Exception.Message)" }
    }
    Update-SelectedRows -Rows $sel
    Update-TagStatistics # Update stats
    Reset-ProgressUI
})

# --- REMOVE TAGS LOGIC ---
$btnRemoveTags.Add_Click({
    $sel = $dataGridView.SelectedRows; if ($sel.Count -eq 0) { return }
    
    if ([System.Windows.Forms.MessageBox]::Show("Clear ALL tags (including Blacklist) for $($sel.Count) numbers?", "Confirm Clear", "YesNo", "Warning") -ne "Yes") { return }
    
    Write-Log "Clearing all tags for $($sel.Count) number(s)..."
    $counter = 0
    foreach ($r in $sel) {
        $counter++
        Update-ProgressUI -Current $counter -Total $sel.Count -Activity "Clearing Tags"
        
        $ph = $r.Cells["TelephoneNumber"].Value; $currentTagStr = [string]$r.Cells["Tag"].Value
        
        if ([string]::IsNullOrWhiteSpace($currentTagStr)) { 
            Write-Log "Skipped $ph (No tags to clear)."
            continue 
        }
        
        $currentTags = $currentTagStr -split "," | ForEach-Object { $_.Trim() }
        
        try { 
            foreach ($t in $currentTags) { 
                Write-Log "Removing tag '$t' from $ph..."; 
                Write-Debug "Executing: Remove-CsPhoneNumberTag -PhoneNumber $ph -Tag $t"; 
                Remove-CsPhoneNumberTag -PhoneNumber $ph -Tag $t -ErrorAction Stop 
            }
            $r.Cells["Tag"].Value = "" # Clear in grid
            Write-Log "All tags cleared for $ph."
        } catch { 
            Write-Log "Failed to clear tags for ${ph}: $($_.Exception.Message)" 
        }
    }
    Update-SelectedRows -Rows $sel
    Update-TagStatistics # Update stats
    Reset-ProgressUI
})
#endregion

#region 14. Logic - Unassign & Remove
# --- UNASSIGN ---
$btnUnassign.Add_Click({ 
    $sel=$dataGridView.SelectedRows; 
    if ($sel.Count -eq 0) { return }

    # BLACKLIST CHECK
    $rowsToProcess = New-Object System.Collections.ArrayList
    foreach ($row in $sel) {
        if (Test-IsBlacklisted $row) { Write-Log "Blocked: $($row.Cells['TelephoneNumber'].Value) is Blacklisted."; continue }
        [void]$rowsToProcess.Add($row)
    }
    if ($rowsToProcess.Count -eq 0) { [System.Windows.Forms.MessageBox]::Show("Operation blocked for all selected numbers (Blacklisted).", "Blocked", "OK", "Warning"); return }

    # UPDATED: Increased information prompt
    $confirmMsg = "You are about to unassign $($rowsToProcess.Count) user(s) from their phone numbers.`n`nThis will:`n1. Remove the assignment in Microsoft Teams.`n2. Attempt to clear 'OfficePhone' and 'TelephoneNumber' attributes in On-Premises AD.`n`nAre you sure you want to proceed?"
    
    if ([System.Windows.Forms.MessageBox]::Show($confirmMsg,"Confirm Unassignment","YesNo","Warning") -eq "Yes"){ 
        Write-Log "Processing Unassign for $($rowsToProcess.Count) row(s)..."
        $counter = 0
        foreach($r in $rowsToProcess){ 
            $counter++
            Update-ProgressUI -Current $counter -Total $rowsToProcess.Count -Activity "Unassigning"
            
            $p=$r.Cells["TelephoneNumber"].Value; 
            $u=$r.Cells["UserPrincipalName"].Value; 
            $t=$r.Cells["NumberType"].Value; 
            
            if (-not [string]::IsNullOrWhiteSpace($u)) { 
                Write-Log "Unassigning $p ($u)..."; 
                try { 
                    Write-Debug "Executing: Remove-CsPhoneNumberAssignment -Identity $u -PhoneNumber $p -PhoneNumberType $t"
                    Remove-CsPhoneNumberAssignment -Identity $u -PhoneNumber $p -PhoneNumberType $t -ErrorAction Stop; 
                    Write-Log "Unassigned $p." 
                    
                    # NEW: Try to sync to AD (Clear OfficePhone)
                    Write-Log "Syncing to On-Prem AD (Clearing OfficePhone)..."
                    try {
                        $adUser = Get-ADUser -Filter "UserPrincipalName -eq '$u'" -ErrorAction Stop
                        if ($adUser) {
                            Set-ADUser -Identity $adUser -Clear "telephoneNumber" -ErrorAction Stop
                            Write-Log "AD OfficePhone cleared for $u."
                        }
                    } catch {
                        Write-Log "AD Sync Warning: $($_.Exception.Message)"
                    }
                    
                } catch { 
                    $errMsg = $_.Exception.Message; 
                    if ($errMsg -match "on-premises Active Directory" -or $errMsg -match "synchronized to the cloud") { 
                        Write-Log "Detected On-Prem user. Clearing AD attributes..."
                        try { 
                            Get-ADUser -Filter "UserPrincipalName -eq '$u'" -Properties msRTCSIP-Line, telephoneNumber -ErrorAction Stop | Set-ADUser -Clear "msRTCSIP-Line", "telephoneNumber" -ErrorAction Stop; 
                            Write-Log "Success: AD attributes cleared." 
                        } catch { Write-Log "Failed to clear AD attributes: $($_.Exception.Message)." } 
                    } else { Write-Log "Error unassigning ${p}: $errMsg" } 
                } 
            } else { Write-Log "Skipped $p (Not assigned)." } 
        }; 
        Update-SelectedRows -Rows $rowsToProcess -ClearUserData
        Update-TagStatistics 
        Reset-ProgressUI
    } 
})

# --- REMOVE ---
$btnRemove.Add_Click({ 
    $sel=$dataGridView.SelectedRows; 
    if ($sel.Count -eq 0) { return }

    # BLACKLIST CHECK
    $rowsToProcess = New-Object System.Collections.ArrayList
    foreach ($row in $sel) {
        if (Test-IsBlacklisted $row) { Write-Log "Blocked: $($row.Cells['TelephoneNumber'].Value) is Blacklisted."; continue }
        [void]$rowsToProcess.Add($row)
    }
    if ($rowsToProcess.Count -eq 0) { [System.Windows.Forms.MessageBox]::Show("Operation blocked for all selected numbers (Blacklisted).", "Blocked", "OK", "Warning"); return }

    # UPDATED: Increased information prompt
    $confirmMsg = "WARNING: You are about to permanently REMOVE $($rowsToProcess.Count) phone number(s) from your tenant.`n`nThis action cannot be undone. The numbers will be released back to the provider or lost.`n`nAny assigned users will be unassigned first.`n`nAre you absolutely sure?"

    if ([System.Windows.Forms.MessageBox]::Show($confirmMsg,"Permanent Removal","YesNo","Error") -eq "Yes"){ 
        $toRemove=@(); Write-Log "Removing $($rowsToProcess.Count) number(s)..."
        $counter = 0
        foreach($r in $rowsToProcess){ 
            $counter++
            Update-ProgressUI -Current $counter -Total $rowsToProcess.Count -Activity "Removing Number"
            
            $p=$r.Cells["TelephoneNumber"].Value; 
            $u=$r.Cells["UserPrincipalName"].Value; 
            $t=$r.Cells["NumberType"].Value; 
            Write-Log "Removing $p..."; 
            try{ 
                if (-not [string]::IsNullOrWhiteSpace($u)) { 
                    Write-Log "  Unassigning user first..."; 
                    try { 
                        Write-Debug "Executing: Remove-CsPhoneNumberAssignment -Identity $u -PhoneNumber $p -PhoneNumberType $t"
                        Remove-CsPhoneNumberAssignment -Identity $u -PhoneNumber $p -PhoneNumberType $t -ErrorAction Stop 
                    } catch { 
                        $errMsg = $_.Exception.Message; 
                        if ($errMsg -match "on-premises Active Directory" -or $errMsg -match "synchronized to the cloud") { 
                            Write-Log "Detected On-Prem user. Clearing AD attributes..."
                            try { 
                                Get-ADUser -Filter "UserPrincipalName -eq '$u'" -Properties msRTCSIP-Line, telephoneNumber -ErrorAction Stop | Set-ADUser -Clear "msRTCSIP-Line", "telephoneNumber" -ErrorAction Stop; 
                                Write-Log "Success: AD attributes cleared." 
                            } catch { Write-Log "Failed to clear AD attributes: $($_.Exception.Message)." } 
                        } else { throw $_ } 
                    } 
                }; 
                Write-Debug "Executing: Remove-CsOnlineTelephoneNumber -TelephoneNumber $p"
                Remove-CsOnlineTelephoneNumber -TelephoneNumber ([string[]]@($p)) -ErrorAction Stop; 
                $toRemove += $r; 
                Write-Log "Removed $p." 
            } catch { Write-Log "Error removing ${p}: $($_.Exception.Message)" } 
        }; 
        foreach ($r in $toRemove) { [void]$dataGridView.Rows.Remove($r) } 
        Update-Stats
        Update-TagStatistics 
        Reset-ProgressUI
    } 
})
#endregion
