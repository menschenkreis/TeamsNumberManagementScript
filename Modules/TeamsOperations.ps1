# =============================================================================
# TEAMS OPERATIONS MODULE - Connection, Fetching, and Orange Sync Logic
# =============================================================================
# NOTE: This module contains event handlers that are registered to UI buttons.
# It must be loaded AFTER UIComponents.ps1

#region Teams Connection Logic
# UPDATED CONNECT BUTTON LOGIC (User Request: No Minimize)
$btnConnect.Add_Click({
        try {
        # REMOVED: $global:form.WindowState = [System.Windows.Forms.FormWindowState]::Minimized
        [System.Windows.Forms.Application]::DoEvents()

        if (Get-Module -ListAvailable -Name MicrosoftTeams) {
            # Module exists, connect
            Connect-MicrosoftTeams -ErrorAction Stop
            # REMOVED: $global:form.WindowState = [System.Windows.Forms.FormWindowState]::Maximized
            Write-Log "Connected."
            $btnFetchData.Enabled = $true
            $global:form.Activate()
        } else {
            # Module missing, ask to install
            # REMOVED: $global:form.WindowState = [System.Windows.Forms.FormWindowState]::Maximized
            $result = [System.Windows.Forms.MessageBox]::Show("MicrosoftTeams PowerShell module is missing.`n`nWould you like to install it now? (Scope: CurrentUser)", "Missing Requirement", "YesNo", "Question")

            if ($result -eq "Yes") {
                try {
                    Write-Log "Installing MicrosoftTeams module..."
                    $global:form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
                    # Install
                    Install-Module -Name MicrosoftTeams -Scope CurrentUser -Force -ErrorAction Stop
                    Write-Log "Installation successful."

                    # Connect immediately after install
                    Connect-MicrosoftTeams -ErrorAction Stop
                    Write-Log "Connected."
                    $btnFetchData.Enabled = $true
                    $global:form.Activate()
                } catch {
                    Write-Log "Installation/Connection Failed: $($_.Exception.Message)"
                    [System.Windows.Forms.MessageBox]::Show("Failed to install or connect: $($_.Exception.Message)", "Error")
                } finally {
                    $global:form.Cursor = [System.Windows.Forms.Cursors]::Default
                }
            }
        }
    } catch {
        # REMOVED: $global:form.WindowState = [System.Windows.Forms.FormWindowState]::Maximized
        $global:form.Cursor = [System.Windows.Forms.Cursors]::Default
        Write-Log "Connection Error: $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show("Connection Error: $($_.Exception.Message)", "Error")
    }
    })
#endregion

#region Data Fetching Logic
# --- COMBINED FETCH LOGIC ---
$btnFetchData.Add_Click({
        # LOGIC UPDATE: We removed the blocking validation here.
        # We will check if Orange is enabled dynamically.
        $orangeEnabled = -not [string]::IsNullOrWhiteSpace($global:txtOrangeKey.Text)

    try {
        $global:form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        $dataGridView.DataSource = $null
        Reset-ProgressUI
        Write-Log "--- STARTING DATA RETRIEVAL ---"

        # 0. PRE-FETCH POLICIES
        Write-Log "Step 1/5: Caching Policies..."
        Update-ProgressUI -Current 10 -Total 100 -Activity "Fetch Policies"
        try {
            # Voice Routing Policies
            Write-Log "  > Fetching Voice Routing Policies..."
            $policies = Get-CsOnlineVoiceRoutingPolicy -ErrorAction Stop
            $global:voiceRoutingPolicies = $policies | Select-Object -ExpandProperty Identity

            # NEW: Teams Meeting Policies
            Write-Log "  > Fetching Teams Meeting Policies..."
            $tmPolicies = Get-CsTeamsMeetingPolicy -ErrorAction Stop
            $global:teamsMeetingPolicies = $tmPolicies | Select-Object -ExpandProperty Identity | Sort-Object

            Write-Log "  > Cached $($global:voiceRoutingPolicies.Count) voice policies and $($global:teamsMeetingPolicies.Count) meeting policies."
        } catch { Write-Log "  > Warning: Failed to fetch policies. ($($_.Exception.Message))"; $global:voiceRoutingPolicies = @(); $global:teamsMeetingPolicies = @() }

        # 1. GET USERS
        Write-Log "Step 2/5: Fetching Teams Users..."
        Update-ProgressUI -Current 20 -Total 100 -Activity "Fetch Users"
        Write-Debug "Executing: Get-CsOnlineUser -ResultSize 20000"
        $users = Get-CsOnlineUser -ResultSize 20000 -ErrorAction Stop

        Write-Log "  > Found $($users.Count) users. Building Index..."
        $global:teamsUsersMap = @{}
        foreach ($u in $users) { if ($u.Identity) { $global:teamsUsersMap[$u.Identity] = $u } }

        # 2. GET NUMBERS
        Write-Log "Step 3/5: Fetching Phone Numbers (Batched)..."
        Update-ProgressUI -Current 30 -Total 100 -Activity "Fetch Numbers"
        $allNumbers = New-Object System.Collections.ArrayList
        $batchSize = 1000; $skip = 0
        while ($skip -lt 10000) {
            Write-Debug "Executing: Get-CsPhoneNumberAssignment -Skip $skip -Top $batchSize"
            $batch = Get-CsPhoneNumberAssignment -Skip $skip -Top $batchSize -ErrorAction Stop
            if (!$batch) { break }
            [void]$allNumbers.AddRange($batch)
            $skip += $batchSize
            Write-Log "  > Fetched batch. Total so far: $($allNumbers.Count)"
            Update-ProgressUI -Current (30 + ($skip / 200)) -Total 100 -Activity "Fetching Numbers ($($allNumbers.Count))"
            if ($batch.Count -lt $batchSize) { break }
        }

        Write-Log "  > Total numbers fetched: $($allNumbers.Count). Building Data Table..."
        $global:masterDataTable = New-Object System.Data.DataTable
        foreach ($c in $global:tableColumns) { $col = New-Object System.Data.DataColumn $c, ([System.String]); $global:masterDataTable.Columns.Add($col) }

        foreach ($num in $allNumbers) {
            $row = $global:masterDataTable.NewRow()
            $row["TelephoneNumber"] = $num.TelephoneNumber
            $row["NumberType"] = $num.NumberType
            $row["ActivationState"] = $num.ActivationState
            $row["City"] = $num.City
            $row["IsoCountryCode"] = $num.IsoCountryCode
            $row["IsoSubdivision"] = $num.IsoSubdivision
            $row["NumberSource"] = $num.NumberSource
            $row["Tag"] = if ($num.Tag) { $num.Tag -join ", " } else { "" }

            $userId = $num.AssignedPstnTargetId
            if ($userId -and $global:teamsUsersMap.ContainsKey($userId)) {
                $u = $global:teamsUsersMap[$userId]
                $row["UserPrincipalName"] = $u.UserPrincipalName
                $row["DisplayName"] = $u.DisplayName
                $row["OnlineVoiceRoutingPolicy"] = $u.OnlineVoiceRoutingPolicy
                $row["EnterpriseVoiceEnabled"] = $u.EnterpriseVoiceEnabled

                # Use checking to ensure columns exist in table definition
                if ($global:tableColumns -contains "AccountEnabled") { $row["AccountEnabled"] = $u.AccountEnabled }
                if ($global:tableColumns -contains "PreferredDataLocation") { $row["PreferredDataLocation"] = $u.PreferredDataLocation }
                if ($global:tableColumns -contains "UsageLocation") { $row["UsageLocation"] = $u.UsageLocation }
                if ($global:tableColumns -contains "TeamsMeetingPolicy") { $row["TeamsMeetingPolicy"] = $u.TeamsMeetingPolicy }

                if ($global:tableColumns -contains "FeatureTypes") {
                    if ($u.FeatureTypes -and $u.FeatureTypes -is [Array]) {
                        $row["FeatureTypes"] = $u.FeatureTypes -join ", "
                    } else {
                        $row["FeatureTypes"] = $u.FeatureTypes
                    }
                }
            }
            [void]$global:masterDataTable.Rows.Add($row)
        }
        Write-Log "  > Teams Data Table built in memory."

        # 3. ORANGE DATA - CONDITIONAL EXECUTION
        if ($orangeEnabled) {
            Write-Log "Step 4/5: Fetching & Merging Orange Data..."

            # Check for required Auth/Customer if Key is present
            if ([string]::IsNullOrWhiteSpace($global:txtOrangeAuth.Text) -or [string]::IsNullOrWhiteSpace($global:txtOrangeCust.Text)) {
                Write-Log "  > Warning: Orange API Key provided, but Auth Header or Customer Code missing. Skipping Orange Sync."
            } else {
                Update-ProgressUI -Current 60 -Total 100 -Activity "Fetch Orange Data"
                $orangeData = $null
                try {
                    $orangeData = Get-AllOrangeCloudNumbers
                    Write-Log "  > Orange data fetched ($($orangeData.Count) records)."
                } catch { Write-Log "  > Orange Fetch Failed: $($_.Exception.Message). Showing Teams data only."; $orangeData = $null }

                # Merge Logic
                if ($orangeData) {
                    Write-Log "  > Merging datasets..."
                    $global:orangeHistoryMap = @{}; $orangeIndex = @{}
                    foreach ($oNum in $orangeData) { $key = [string]$oNum.number.Replace("+", "").Trim(); $orangeIndex[$key] = $oNum; $global:orangeHistoryMap["+" + $key] = $oNum.history }

                    $dt = $global:masterDataTable; $processedKeys = @{}

                    for ($i = 0; $i -lt $dt.Rows.Count; $i++) {
                        $row = $dt.Rows[$i]
                        $tKey = [string]$row["TelephoneNumber"].Replace("+", "").Trim()
                        if ($orangeIndex.ContainsKey($tKey)) {
                            $oObj = $orangeIndex[$tKey]; $siteName = ""; $siteId = ""; $vs = $oObj.voiceSite
                            if ($null -ne $vs) { $siteId = $vs.voiceSiteId; $siteName = $vs.technicalSiteName; if ([string]::IsNullOrWhiteSpace($siteId) -and $vs -is [System.Collections.IDictionary]) { $siteId = $vs['voiceSiteId']; $siteName = $vs['technicalSiteName'] } }
                            $row["OrangeSite"] = $siteName; $row["OrangeSiteId"] = $siteId; $row["OrangeStatus"] = $oObj.status; $row["OrangeUsage"] = $oObj.usage; $processedKeys[$tKey] = $true
                        }
                        if ($i % 200 -eq 0) {
                            Update-ProgressUI -Current (60 + (($i / $dt.Rows.Count) * 30)) -Total 100 -Activity "Merging Data ($i/$($dt.Rows.Count))"
                        }
                    }

                    # Add Orange-only numbers
                    Write-Log "  > Adding Orange-only numbers..."
                    $missing = $orangeIndex.Keys | Where-Object { -not $processedKeys.ContainsKey($_) }
                    foreach ($key in $missing) {
                        $oObj = $orangeIndex[$key]; $row = $dt.NewRow(); $row["TelephoneNumber"] = "+" + $key; $siteName = ""; $siteId = ""; $vs = $oObj.voiceSite
                        if ($null -ne $vs) { $siteId = $vs.voiceSiteId; $siteName = $vs.technicalSiteName; if ([string]::IsNullOrWhiteSpace($siteId) -and $vs -is [System.Collections.IDictionary]) { $siteId = $vs['voiceSiteId']; $siteName = $vs['technicalSiteName'] } }
                        $row["OrangeSite"] = $siteName; $row["OrangeSiteId"] = $siteId; $row["OrangeStatus"] = $oObj.status; $row["OrangeUsage"] = $oObj.usage; $dt.Rows.Add($row)
                    }
                }
            }
        } else {
            Write-Log "Step 4/5: Skipping Orange Sync (API Key not populated)."
        }

        # 4. RENDER
        Write-Log "Step 5/5: Rendering Grid..."
        Update-ProgressUI -Current 95 -Total 100 -Activity "Rendering Grid"
        $dataGridView.DataSource = $global:masterDataTable
        foreach ($hc in $global:defaultHiddenCols) { if ($dataGridView.Columns[$hc]) { $dataGridView.Columns[$hc].Visible = $false } }
        $btnSyncOrange.Enabled = $true

        # 5. DYNAMIC FILTER UPDATE
        Write-Log "Updating Filter Tags..."
        Update-FilterTags

        Write-Log "--- PROCESS COMPLETE ---"
        Update-Stats
        Update-TagStatistics
        Update-ProgressUI -Current 100 -Total 100 -Activity "Ready"
        Reset-ProgressUI
    } catch {
        Write-Log "ERROR: $($_.Exception.Message)"
    } finally {
        $global:form.Cursor = [System.Windows.Forms.Cursors]::Default
    }
    })
#endregion

#region Orange Sync Logic
# --- AUTH RETRY LOOP ---
$btnSyncOrange.Add_Click({
        # GUARD: Key Check
        if ([string]::IsNullOrWhiteSpace($global:txtOrangeKey.Text)) {
        [System.Windows.Forms.MessageBox]::Show("Orange API Key is not populated. Feature disabled.", "Feature Disabled", "OK", "Warning"); return
    }

    if ($global:masterDataTable -eq $null) { Write-Log "No data."; return }
    if ([string]::IsNullOrWhiteSpace($global:txtOrangeAuth.Text) -or [string]::IsNullOrWhiteSpace($global:txtOrangeCust.Text)) {
        [System.Windows.Forms.MessageBox]::Show("You must populate all Orange Configuration fields.", "Validation Error", "OK", "Warning"); return
    }

    $orangeData = $null
    try {
        $global:form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor; Reset-ProgressUI; $dataGridView.DataSource = $null;
        $orangeData = Get-AllOrangeCloudNumbers
    }
    catch { Write-Log "Error: $($_.Exception.Message)" }
    finally { $global:form.Cursor = [System.Windows.Forms.Cursors]::Default }

    if ($orangeData) {
        Write-Log "Indexing Orange..."; $global:orangeHistoryMap = @{}; $orangeIndex = @{}
        foreach ($oNum in $orangeData) { $key = [string]$oNum.number.Replace("+", "").Trim(); $orangeIndex[$key] = $oNum; $global:orangeHistoryMap["+" + $key] = $oNum.history }
        Write-Log "Merging..."; $dt = $global:masterDataTable; $processedKeys = @{}
        for ($i = 0; $i -lt $dt.Rows.Count; $i++) {
            $row = $dt.Rows[$i]; $tKey = [string]$row["TelephoneNumber"].Replace("+", "").Trim()
            if ($orangeIndex.ContainsKey($tKey)) {
                $oObj = $orangeIndex[$tKey]; $siteName = ""; $siteId = ""; $vs = $oObj.voiceSite; if ($null -ne $vs) { $siteId = $vs.voiceSiteId; $siteName = $vs.technicalSiteName; if ([string]::IsNullOrWhiteSpace($siteId) -and $vs -is [System.Collections.IDictionary]) { $siteId = $vs['voiceSiteId']; $siteName = $vs['technicalSiteName'] } }
                $row["OrangeSite"] = $siteName; $row["OrangeSiteId"] = $siteId; $row["OrangeStatus"] = $oObj.status; $row["OrangeUsage"] = $oObj.usage; $processedKeys[$tKey] = $true
            }
            if ($i % 500 -eq 0) { Update-ProgressUI -Current $i -Total $dt.Rows.Count -Activity "Merging Orange Data" }
        }
        Write-Log "Adding new..."; $missing = $orangeIndex.Keys | Where-Object { -not $processedKeys.ContainsKey($_) }; $idx = 0
        foreach ($key in $missing) {
            $oObj = $orangeIndex[$key]; $row = $dt.NewRow(); $row["TelephoneNumber"] = "+" + $key; $siteName = ""; $siteId = ""; $vs = $oObj.voiceSite; if ($null -ne $vs) { $siteId = $vs.voiceSiteId; $siteName = $vs.technicalSiteName; if ([string]::IsNullOrWhiteSpace($siteId) -and $vs -is [System.Collections.IDictionary]) { $siteId = $vs['voiceSiteId']; $siteName = $vs['technicalSiteName'] } }
            $row["OrangeSite"] = $siteName; $row["OrangeSiteId"] = $siteId; $row["OrangeStatus"] = $oObj.status; $row["OrangeUsage"] = $oObj.usage; $dt.Rows.Add($row); $idx++
            if ($idx % 100 -eq 0) { Update-ProgressUI -Current $idx -Total $missing.Count -Activity "Adding New Numbers" }
        }
        $dataGridView.DataSource = $global:masterDataTable; Write-Log "Sync Complete."

        foreach ($hc in $global:defaultHiddenCols) { if ($dataGridView.Columns[$hc]) { $dataGridView.Columns[$hc].Visible = $false } }

        # Update Dynamic Filters
        Update-FilterTags

        Update-Stats
        Update-TagStatistics
        Update-ProgressUI -Current 100 -Total 100 -Activity "Done"
        Reset-ProgressUI
    }
    })
#endregion
