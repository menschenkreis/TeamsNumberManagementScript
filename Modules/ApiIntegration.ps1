# =============================================================================
# API INTEGRATION MODULE - Orange API and Proxy Logic
# =============================================================================

# --- HELPER: Get Proxy Parameters Dictionary ---
function Get-ProxyParams {
    $proxy = $global:txtProxy.Text.Trim()
    $params = @{}
    if (-not [string]::IsNullOrWhiteSpace($proxy)) {
        $params['Proxy'] = $proxy
        $params['ProxyUseDefaultCredentials'] = $true
    }
    return $params
}

function Get-OrangeHeaders {
    $authHeaderB64 = $global:txtOrangeAuth.Text.Trim()
    $apiKey = $global:txtOrangeKey.Text.Trim()

    # Validation
    if ([string]::IsNullOrWhiteSpace($apiKey)) { Throw "Orange API Key is missing in configuration." }
    if ([string]::IsNullOrWhiteSpace($authHeaderB64)) { Throw "Orange Auth Header is missing in configuration." }

    $proxyParams = Get-ProxyParams

    [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
    try {
        $authUrl = "https://api.orange.com/oauth/v3/token"
        # SPLATTING FIX
        $authResponse = Invoke-WebRequest -Uri $authUrl -Method POST -Headers @{ 'Authorization' = $authHeaderB64; 'Accept' = 'application/json' } -Body @{ 'grant_type' = 'client_credentials' } -ContentType 'application/x-www-form-urlencoded' @proxyParams -UseBasicParsing
        $accessToken = ($authResponse.Content | ConvertFrom-Json).access_token
        return @{ 'X-BT-API-KEY' = $apiKey; 'Authorization' = "Bearer $accessToken"; 'Accept' = 'application/json' }
    }
    catch {
        $errBody = "Unknown"; if ($_.Exception.Response) { try { $reader = New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream()); $errBody = $reader.ReadToEnd() } catch {} }
        Throw "Orange Auth Failed: $($_.Exception.Message) | Body: $errBody"
    }
}

function Get-AllOrangeCloudNumbers {
    [CmdletBinding()] param()
    $customerCode = $global:txtOrangeCust.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($customerCode)) { Throw "Customer Code is missing." }

    $proxyParams = Get-ProxyParams

    $Service = "opc"; $apiHeaders = Get-OrangeHeaders
    $allNumbers = [System.Collections.Generic.List[Object]]::new()
    $offset = 0; $limit = 1000; $hasMore = $true
    Write-Log "Orange: Fetching data from API ($customerCode)..."
    do {
        $url = "https://api.orange.com/btalk/v1/customers/$customerCode/$Service/cloudvoicenumbers?limit=$limit&offset=$offset"
        try {
            # SPLATTING FIX
            $response = Invoke-WebRequest -Uri $url -Method Get -Headers $apiHeaders -ContentType 'application/json' @proxyParams -UseBasicParsing
            $batch = $response.Content | ConvertFrom-Json
            if ($batch) {
                $batchArray = @($batch)
                if ($batchArray.Count -gt 0) {
                    $allNumbers.AddRange($batchArray)
                    Write-Log "  > Retrieved $($batchArray.Count) numbers (Offset $offset)."
                    $offset += $limit
                    if ($batchArray.Count -lt $limit) { $hasMore = $false }
                } else { $hasMore = $false }
            } else { $hasMore = $false }
        } catch { Write-Log "Orange Error: $($_.Exception.Message)"; $hasMore = $false }
        [System.Windows.Forms.Application]::DoEvents()
    } while ($hasMore)
    Write-Log "Orange: Total fetched: $($allNumbers.Count)"
    return $allNumbers
}

function Get-SingleOrangeNumber {
    param([string]$PhoneNumber)
    $customerCode = $global:txtOrangeCust.Text.Trim(); $Service = "opc"; $apiHeaders = Get-OrangeHeaders

    $proxyParams = Get-ProxyParams

    $cleanNum = $PhoneNumber.Replace("+", "").Trim()
    $url = "https://api.orange.com/btalk/v1/customers/$customerCode/$Service/cloudvoicenumbers?number=$cleanNum"
    try {
        # SPLATTING FIX
        $response = Invoke-WebRequest -Uri $url -Method Get -Headers $apiHeaders -ContentType 'application/json' @proxyParams -UseBasicParsing
        $array = $response.Content | ConvertFrom-Json
        if ($array -and $array.Count -gt 0) { return $array[0] }
        return $null
    } catch { Write-Log "Orange Single Fetch Error ($cleanNum): $($_.Exception.Message)"; return $null }
}

function Update-SelectedRows {
    param([System.Collections.IList]$Rows, [string]$NewOrangeStatus = $null, [string]$NewOrangeUsage = $null, [switch]$ClearUserData)
    if ($null -eq $Rows -or $Rows.Count -eq 0) { return }
    foreach ($row in $Rows) {
        if (-not [string]::IsNullOrWhiteSpace($NewOrangeStatus)) { $row.Cells["OrangeStatus"].Value = $NewOrangeStatus }
        if (-not [string]::IsNullOrWhiteSpace($NewOrangeUsage)) { $row.Cells["OrangeUsage"].Value = $NewOrangeUsage }
        if ($ClearUserData) {
            $row.Cells["UserPrincipalName"].Value = ""; $row.Cells["DisplayName"].Value = ""; $row.Cells["OnlineVoiceRoutingPolicy"].Value = ""; $row.Cells["EnterpriseVoiceEnabled"].Value = ""; $row.Cells["ActivationState"].Value = "Unassigned"
            # Clear new columns as well
            if ($row.DataGridView.Columns.Contains("AccountEnabled")) { $row.Cells["AccountEnabled"].Value = "" }
            if ($row.DataGridView.Columns.Contains("PreferredDataLocation")) { $row.Cells["PreferredDataLocation"].Value = "" }
            if ($row.DataGridView.Columns.Contains("UsageLocation")) { $row.Cells["UsageLocation"].Value = "" }
            if ($row.DataGridView.Columns.Contains("TeamsMeetingPolicy")) { $row.Cells["TeamsMeetingPolicy"].Value = "" }
            if ($row.DataGridView.Columns.Contains("FeatureTypes")) { $row.Cells["FeatureTypes"].Value = "" }
        }
    }
}

function Group-PhoneNumbers {
    param([int64[]]$NumbersList)
    $sortedNumbers = $NumbersList | Sort-Object | Get-Unique; $ranges = @(); $start = $null; $end = $null
    foreach ($number in $sortedNumbers) {
        if ($null -eq $start) { $start = $number; $end = $number } elseif ($number -eq $end + 1) { $end = $number } else { $ranges += [PSCustomObject]@{ Start = $start; End = $end }; $start = $number; $end = $number }
    }
    if ($null -ne $start) { $ranges += [PSCustomObject]@{ Start = $start; End = $end } }
    return $ranges
}

function Publish-OrangeNumbers {
    param([System.Collections.ArrayList]$RowObjects, [string]$UsageType)
    $customerCode = $global:txtOrangeCust.Text.Trim()
    $Service = "opc"; $apiHeaders = Get-OrangeHeaders

    $proxyParams = Get-ProxyParams

    $successfulRows = New-Object System.Collections.ArrayList

    $counter = 0
    $total = $RowObjects.Count

    foreach ($row in $RowObjects) {
        $counter++
        Update-ProgressUI -Current $counter -Total $total -Activity "Publishing to OC"

        $siteId = $row.Cells["OrangeSiteId"].Value
        $rawNum = [string]$row.Cells["TelephoneNumber"].Value
        $cleanNum = $rawNum.Replace("+", "").Trim().Split(';')[0]

        if ([string]::IsNullOrWhiteSpace($siteId)) { Write-Log "SKIP: $rawNum - No SiteId."; continue }

        $numbersPayload = @( @{ usage = $UsageType; rangeStart = $cleanNum; rangeStop = $cleanNum } )
        $body = @{ voiceOffer = "btlvs"; voiceSiteId = $siteId; numbers = $numbersPayload } | ConvertTo-Json -Depth 5 -Compress
        $url = "https://api.orange.com/btalk/v1/customers/$customerCode/$Service/cloudvoicenumbers"

        try {
            # SPLATTING FIX
            Invoke-WebRequest -Uri $url -Method POST -Headers $apiHeaders -Body $body -ContentType 'application/json' @proxyParams -UseBasicParsing | Out-Null
            Write-Log "Success: Published $cleanNum"
            [void]$successfulRows.Add($row)
        }
        catch {
            $errBody = "Unknown"; if ($_.Exception.Response) { try { $reader = New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream()); $errBody = $reader.ReadToEnd() } catch {} }
            Write-Log "Failed ($cleanNum): $($_.Exception.Message) | Resp: $errBody"
        }
    }
    return $successfulRows
}

function Unpublish-OrangeNumbersBatch {
    param([System.Collections.ArrayList]$RowObjects)
    $customerCode = $global:txtOrangeCust.Text.Trim(); $Service = "opc"; $apiHeaders = Get-OrangeHeaders

    $proxyParams = Get-ProxyParams

    $groupedBySite = $RowObjects | Group-Object -Property { $_.Cells["OrangeSiteId"].Value }
    $successfulRows = New-Object System.Collections.ArrayList

    $grpCounter = 0
    foreach ($group in $groupedBySite) {
        $grpCounter++
        Update-ProgressUI -Current $grpCounter -Total $groupedBySite.Count -Activity "Releasing Batch (By Site)"

        $siteId = $group.Name; if ([string]::IsNullOrWhiteSpace($siteId)) { Write-Log "SKIP: No SiteId."; continue }
        $intNumbers = @(); foreach ($r in $group.Group) { $raw = [string]$r.Cells["TelephoneNumber"].Value; $clean = $raw.Replace("+", "").Trim(); if ($clean -match "^\d+$") { $intNumbers += [int64]$clean } }
        $ranges = Group-PhoneNumbers -NumbersList $intNumbers
        $numbersPayload = @(); foreach ($rng in $ranges) { $numbersPayload += @{ usage = "CallingUserAssignment"; rangeStart = [string]$rng.Start; rangeStop = [string]$rng.End } }
        $body = @{ voiceOffer = "btlvs"; voiceSiteId = $siteId; numbers = $numbersPayload } | ConvertTo-Json -Depth 5 -Compress
        $url = "https://api.orange.com/btalk/v1/customers/$customerCode/$Service/cloudvoicenumbers/release"
        Write-Log "Orange API (POST): Releasing $($group.Count) numbers for Site $siteId..."
        try {
            # SPLATTING FIX
            Invoke-WebRequest -Uri $url -Method POST -Headers $apiHeaders -Body $body -ContentType 'application/json' @proxyParams -UseBasicParsing | Out-Null;
            Write-Log "Success: Batch Released."; [void]$successfulRows.AddRange($group.Group)
        }
        catch { $errBody = "Unknown"; if ($_.Exception.Response) { try { $reader = New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream()); $errBody = $reader.ReadToEnd() } catch {} }; Write-Log "API Failed: $($_.Exception.Message) | Response: $errBody" }
    }
    return $successfulRows
}
