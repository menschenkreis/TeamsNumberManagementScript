# =============================================================================
# TEAMS PHONE MANAGER v56.4 (UI Spacing & Centering)
# =============================================================================

#region 1. Cleanup & Assemblies
# --- CLEANUP SECTION ---
if ($global:form -and !$global:form.IsDisposed) {
    try {
        $global:form.Close()
        $global:form.Dispose()
    } catch {
        Write-Host "Warning: Could not cleanly dispose previous form - $($_.Exception.Message)"
    }
}
$global:form = $null

# Load Assemblies
try { Add-Type -AssemblyName System.Windows.Forms } catch { Write-Host "Warning: Failed to load System.Windows.Forms - $($_.Exception.Message)" }
try { Add-Type -AssemblyName System.Drawing } catch { Write-Host "Warning: Failed to load System.Drawing - $($_.Exception.Message)" }
try { Add-Type -AssemblyName System.Data } catch { Write-Host "Warning: Failed to load System.Data - $($_.Exception.Message)" }
try { Add-Type -AssemblyName Microsoft.VisualBasic } catch { Write-Host "Warning: Failed to load Microsoft.VisualBasic - $($_.Exception.Message)" }
#endregion

#region 2. Global Variables & Logging
# --- Global Data ---
$global:teamsUsersMap = @{}
$global:orangeHistoryMap = @{} 
$global:masterDataTable = $null 
$global:form = $null 
$global:voiceRoutingPolicies = @()
$global:teamsMeetingPolicies = @()
$global:hideUnassigned = $false
$global:appVersion = "v56.4"

# --- SETTINGS GLOBALS ---
$global:settingsXmlPath = $null
$global:lowStockThreshold = 5 # Default

# UPDATED: Default Tags Removed. Now relies solely on XML or Manual Input.
$global:allowedTags = @() 

# --- LOGGING SETUP ---
$global:logFilePath = $null

try {
    # Determine script location (fallback to current dir if running unsaved)
    $scriptPath = if ($PSScriptRoot) { $PSScriptRoot } else { $PWD.Path }
    $logDir = Join-Path $scriptPath "ExecutionLogs"
    
    # Create Directory
    if (-not (Test-Path $logDir)) {
        New-Item -ItemType Directory -Path $logDir -Force | Out-Null
    }

    # Generate Filename
    $logName = "Log_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
    $global:logFilePath = Join-Path $logDir $logName
    
    # Set Default XML Path
    $global:defaultXmlPath = Join-Path $scriptPath "Settings.xml"

    # Initialize File
    $initMsg = @"
===========================================================
TEAMS PHONE MANAGER EXECUTION LOG
User:      $([Environment]::UserName)
Date:      $(Get-Date)
Machine:   $([Environment]::MachineName)
===========================================================
"@
    $initMsg | Out-File -FilePath $global:logFilePath -Encoding UTF8
} catch {
    Write-Host "Warning: Could not initialize log file. $($_.Exception.Message)"
}
#endregion

#region 3. UI References & Column Definitions
# --- GLOBAL UI REFERENCES ---
$global:txtOrangeAuth = $null
$global:txtOrangeCust = $null
$global:txtOrangeKey = $null
$global:txtProxy = $null
$global:dgvTagStats = $null 
$global:lblValTotal = $null
$global:lblValDisp = $null
$global:lblValSel = $null
$global:cmbFilterTag = $null 
$global:cmbTag = $null # Reference for the Action Tag dropdown

# Define Table Columns
$global:tableColumns = @(
    "TelephoneNumber", 
    "NumberType", 
    "ActivationState", 
    "City", 
    "IsoCountryCode", 
    "IsoSubdivision", 
    "NumberSource", 
    "Tag", 
    "UserPrincipalName", 
    "DisplayName", 
    "OnlineVoiceRoutingPolicy", 
    "TeamsMeetingPolicy", 
    "EnterpriseVoiceEnabled", 
    "PreferredDataLocation", 
    "UsageLocation", 
    "FeatureTypes", 
    "OrangeSite", 
    "OrangeStatus", 
    "OrangeUsage", 
    "OrangeSiteId"
)

# Columns to hide by default
$global:defaultHiddenCols = @(
    "ActivationState", 
    "City", 
    "IsoCountryCode", 
    "IsoSubdivision", 
    "TeamsMeetingPolicy", 
    "PreferredDataLocation", 
    "UsageLocation", 
    "FeatureTypes", 
    "OrangeUsage", 
    "OrangeSiteId",
    "EnterpriseVoiceEnabled"
)
#endregion

# =============================================================================
# FUNCTIONS
# =============================================================================

#region 4. Functions - Helper & UI
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
        $global:form.Text = "Teams Phone Manager $($global:appVersion) - $Activity ($Current / $Total)"
    }
    [System.Windows.Forms.Application]::DoEvents()
}

function Reset-ProgressUI {
    if ($progressBar) { $progressBar.Value = 0 }
    if ($global:form) { $global:form.Text = "Teams Phone Manager $($global:appVersion)" }
    [System.Windows.Forms.Application]::DoEvents()
}

function Update-Stats {
    $total = if ($global:masterDataTable) { $global:masterDataTable.Rows.Count } else { 0 }
    $disp = if ($dataGridView) { $dataGridView.Rows.Count } else { 0 }
    $sel = if ($dataGridView) { $dataGridView.SelectedRows.Count } else { 0 }

    if ($global:lblValTotal) { $global:lblValTotal.Text = "$total" }
    if ($global:lblValDisp)  { $global:lblValDisp.Text  = "$disp" }
    if ($global:lblValSel)   { $global:lblValSel.Text   = "$sel" }
}

function Get-SimpleInput {
    param([string]$Title = "Input", [string]$Prompt = "Please enter value:")

    $f = New-Object System.Windows.Forms.Form
    $f.Width = 400; $f.Height = 180; $f.Text = $Title
    $f.StartPosition = "CenterParent"; $f.FormBorderStyle = "FixedDialog"
    $f.MaximizeBox = $false; $f.MinimizeBox = $false

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Location = New-Object System.Drawing.Point(20, 20)
    $lbl.Size = New-Object System.Drawing.Size(340, 20)
    $lbl.Text = $Prompt
    $f.Controls.Add($lbl)

    $txt = New-Object System.Windows.Forms.TextBox
    $txt.Location = New-Object System.Drawing.Point(20, 50)
    $txt.Size = New-Object System.Drawing.Size(340, 20)
    $f.Controls.Add($txt)

    $btnOk = New-Object System.Windows.Forms.Button
    $btnOk.Text = "OK"; $btnOk.DialogResult = "OK"
    $btnOk.Location = New-Object System.Drawing.Point(200, 90)
    $f.Controls.Add($btnOk)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Cancel"; $btnCancel.DialogResult = "Cancel"
    $btnCancel.Location = New-Object System.Drawing.Point(280, 90)
    $f.Controls.Add($btnCancel)

    $f.AcceptButton = $btnOk; $f.CancelButton = $btnCancel
    if ($f.ShowDialog() -eq "OK") { return $txt.Text }
    return $null
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
    $sfd = New-Object System.Windows.Forms.SaveFileDialog
    $sfd.Filter = "CSV Files (*.csv)|*.csv"
    $sfd.Title = "Save Export As"
    $sfd.FileName = "TeamsPhoneExport_$(Get-Date -Format 'yyyyMMdd_HHmm').csv"

    if ($sfd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $path = $sfd.FileName
        Write-Log "Exporting $($sel.Count) rows to CSV..."
        $global:form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        try {
            $exportList = [System.Collections.Generic.List[PSCustomObject]]::new()
            foreach ($row in $sel) {
                $obj = [Ordered]@{}
                foreach ($colName in $global:tableColumns) {
                    $val = $row.Cells[$colName].Value
                    if ($null -eq $val) { $val = "" }
                    $obj[$colName] = [string]$val
                }
                $exportList.Add([PSCustomObject]$obj)
            }
            $exportList | Export-Csv -Path $path -NoTypeInformation -Delimiter "," -Encoding UTF8
            Write-Log "Export saved to: $path"
            [System.Windows.Forms.MessageBox]::Show("Export successful!", "Done")
        } catch {
            Write-Log "Export Error: $($_.Exception.Message)"
            [System.Windows.Forms.MessageBox]::Show("Export Failed: $($_.Exception.Message)", "Error")
        } finally {
            $global:form.Cursor = [System.Windows.Forms.Cursors]::Default
        }
    }
}
#endregion

#region 5. Functions - XML Settings
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
        if ($settings.OrangeApiKey)     { $global:txtOrangeKey.Text = $settings.OrangeApiKey }
        if ($settings.Proxy)            { $global:txtProxy.Text = $settings.Proxy }
        
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
        $xml.Settings.OrangeApiKey     = $global:txtOrangeKey.Text
        $xml.Settings.Proxy            = $global:txtProxy.Text
        
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
#endregion

#region 6. Functions - Tags & Stats Logic
function Update-TagStatistics {
    # UPDATED: This now scans the actual data to determine what tags to show in the Stats Grid.
    if ($null -eq $global:dgvTagStats) { return }
    if ($null -eq $global:masterDataTable) { return }

    $stats = @{}
    
    # 1. Scan Data for Unique Tags and Calculate Counts
    foreach ($row in $global:masterDataTable.Rows) {
        $tagStr = [string]$row["Tag"]
        $upn = [string]$row["UserPrincipalName"]
        $isAssigned = -not [string]::IsNullOrWhiteSpace($upn)
        $isReserved = $tagStr -match "Reserved" 
        
        if (-not [string]::IsNullOrWhiteSpace($tagStr)) {
            $tagsInRow = $tagStr.Split(',')
            foreach ($rawTag in $tagsInRow) {
                $t = $rawTag.Trim()
                if (-not [string]::IsNullOrWhiteSpace($t)) {
                    # Initialize if new
                    if (-not $stats.ContainsKey($t)) {
                        $stats[$t] = @{ Total = 0; Free = 0 }
                    }
                    
                    # Increment
                    $stats[$t].Total++
                    if (-not $isAssigned -and -not $isReserved) {
                        $stats[$t].Free++
                    }
                }
            }
        }
    }

    # 2. Rebuild Grid Rows based on Discovered Tags
    $global:dgvTagStats.Rows.Clear()
    
    $sortedTags = $stats.Keys | Sort-Object
    
    foreach ($tagKey in $sortedTags) {
        $total = $stats[$tagKey].Total
        $free  = $stats[$tagKey].Free
        
        $rowIndex = $global:dgvTagStats.Rows.Add($tagKey, $total, $free)
        $row = $global:dgvTagStats.Rows[$rowIndex]

        # Dynamic Threshold Coloring
        $limit = $global:lowStockThreshold
        $warningLimit = $limit * 2 

        if ($free -lt $limit) {
            $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::LightCoral
            $row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::Black
            $row.DefaultCellStyle.SelectionBackColor = [System.Drawing.Color]::FromArgb(200, 200, 200)
            $row.DefaultCellStyle.SelectionForeColor = [System.Drawing.Color]::Black
        } elseif ($free -lt $warningLimit) {
            $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::Orange
            $row.DefaultCellStyle.SelectionBackColor = [System.Drawing.Color]::FromArgb(200, 200, 200)
            $row.DefaultCellStyle.SelectionForeColor = [System.Drawing.Color]::Black
        } else {
            $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::White
            $row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::Black
        }
    }
    
    $global:dgvTagStats.ClearSelection()
}

# --- Dynamically Update Filter Tags ---
function Update-FilterTags {
    if ($null -eq $global:masterDataTable) { return }
    if ($null -eq $global:cmbFilterTag) { return }

    # Save current selection if possible
    $currentSelection = $global:cmbFilterTag.SelectedItem

    $uniqueTags = New-Object System.Collections.Generic.HashSet[string]
    
    foreach ($row in $global:masterDataTable.Rows) {
        $rawTag = [string]$row["Tag"]
        if (-not [string]::IsNullOrWhiteSpace($rawTag)) {
            $parts = $rawTag.Split(',')
            foreach ($p in $parts) {
                $trimmed = $p.Trim()
                if (-not [string]::IsNullOrWhiteSpace($trimmed)) {
                    [void]$uniqueTags.Add($trimmed)
                }
            }
        }
    }

    $sortedTags = $uniqueTags | Sort-Object

    $global:cmbFilterTag.Items.Clear()
    $global:cmbFilterTag.Items.Add("All")
    
    foreach ($t in $sortedTags) {
        [void]$global:cmbFilterTag.Items.Add($t)
    }

    # Restore selection or default to All
    if ($currentSelection -and $global:cmbFilterTag.Items.Contains($currentSelection)) {
        $global:cmbFilterTag.SelectedItem = $currentSelection
    } else {
        $global:cmbFilterTag.SelectedIndex = 0
    }
}
#endregion

#region 7. Functions - API Logic (Orange & Proxy)
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
        if (-not [string]::IsNullOrWhiteSpace($NewOrangeUsage))  { $row.Cells["OrangeUsage"].Value = $NewOrangeUsage }
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

function Merge-OrangeData {
    param([System.Collections.Generic.List[Object]]$OrangeData)

    Write-Log "Indexing Orange..."
    $global:orangeHistoryMap = @{}
    $orangeIndex = @{}
    foreach ($oNum in $OrangeData) {
        $key = [string]$oNum.number.Replace("+","").Trim()
        $orangeIndex[$key] = $oNum
        $global:orangeHistoryMap["+" + $key] = $oNum.history
    }

    Write-Log "Merging..."
    $dt = $global:masterDataTable
    $processedKeys = @{}

    for ($i = 0; $i -lt $dt.Rows.Count; $i++) {
        $row = $dt.Rows[$i]
        $tKey = [string]$row["TelephoneNumber"].Replace("+","").Trim()
        if ($orangeIndex.ContainsKey($tKey)) {
            $oObj = $orangeIndex[$tKey]; $siteName = ""; $siteId = ""; $vs = $oObj.voiceSite
            if ($null -ne $vs) { $siteId = $vs.voiceSiteId; $siteName = $vs.technicalSiteName; if ([string]::IsNullOrWhiteSpace($siteId) -and $vs -is [System.Collections.IDictionary]) { $siteId = $vs['voiceSiteId']; $siteName = $vs['technicalSiteName'] } }
            $row["OrangeSite"] = $siteName; $row["OrangeSiteId"] = $siteId; $row["OrangeStatus"] = $oObj.status; $row["OrangeUsage"] = $oObj.usage
            $processedKeys[$tKey] = $true
        }
        if ($i % 200 -eq 0) {
            Update-ProgressUI -Current $i -Total $dt.Rows.Count -Activity "Merging Orange Data ($i/$($dt.Rows.Count))"
        }
    }

    Write-Log "Adding Orange-only numbers..."
    $missing = $orangeIndex.Keys | Where-Object { -not $processedKeys.ContainsKey($_) }
    $idx = 0
    foreach ($key in $missing) {
        $oObj = $orangeIndex[$key]; $row = $dt.NewRow(); $row["TelephoneNumber"] = "+" + $key; $siteName = ""; $siteId = ""; $vs = $oObj.voiceSite
        if ($null -ne $vs) { $siteId = $vs.voiceSiteId; $siteName = $vs.technicalSiteName; if ([string]::IsNullOrWhiteSpace($siteId) -and $vs -is [System.Collections.IDictionary]) { $siteId = $vs['voiceSiteId']; $siteName = $vs['technicalSiteName'] } }
        $row["OrangeSite"] = $siteName; $row["OrangeSiteId"] = $siteId; $row["OrangeStatus"] = $oObj.status; $row["OrangeUsage"] = $oObj.usage
        $dt.Rows.Add($row)
        $idx++
        if ($idx % 100 -eq 0) { Update-ProgressUI -Current $idx -Total $missing.Count -Activity "Adding New Numbers" }
    }
    Write-Log "Merge complete."
}
#endregion

# =============================================================================
# GUI SETUP
# =============================================================================

#region 8. GUI Construction
$global:form = New-Object System.Windows.Forms.Form
$global:form.Text = "Teams Phone Manager $($global:appVersion)"
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
# Vertical Centering (Y=31 for Labels, Y=28 for Controls)
# Horizontal Spacing (10px gap between Label end and Input start)

$lblAuth = New-Object System.Windows.Forms.Label
$lblAuth.Location = New-Object System.Drawing.Point(10, 31)
$lblAuth.Size = New-Object System.Drawing.Size(80, 20); $lblAuth.Text = "Auth Header:"

$global:txtOrangeAuth = New-Object System.Windows.Forms.TextBox
$global:txtOrangeAuth.Location = New-Object System.Drawing.Point(100, 28)
$global:txtOrangeAuth.Size = New-Object System.Drawing.Size(200, 20); $global:txtOrangeAuth.PasswordChar = "*"

$lblCust = New-Object System.Windows.Forms.Label
$lblCust.Location = New-Object System.Drawing.Point(320, 31)
$lblCust.Size = New-Object System.Drawing.Size(80, 20); $lblCust.Text = "Customer Id:"

$global:txtOrangeCust = New-Object System.Windows.Forms.TextBox
$global:txtOrangeCust.Location = New-Object System.Drawing.Point(410, 28)
$global:txtOrangeCust.Size = New-Object System.Drawing.Size(60, 20); $global:txtOrangeCust.Text = ""

$lblKey = New-Object System.Windows.Forms.Label
$lblKey.Location = New-Object System.Drawing.Point(490, 31)
$lblKey.Size = New-Object System.Drawing.Size(60, 20); $lblKey.Text = "API Key:"

$global:txtOrangeKey = New-Object System.Windows.Forms.TextBox
$global:txtOrangeKey.Location = New-Object System.Drawing.Point(560, 28)
$global:txtOrangeKey.Size = New-Object System.Drawing.Size(120, 20); $global:txtOrangeKey.PasswordChar = "*"

$lblProxy = New-Object System.Windows.Forms.Label
$lblProxy.Location = New-Object System.Drawing.Point(700, 31)
$lblProxy.Size = New-Object System.Drawing.Size(50, 20); $lblProxy.Text = "Proxy:"

$global:txtProxy = New-Object System.Windows.Forms.TextBox
$global:txtProxy.Location = New-Object System.Drawing.Point(760, 28)
$global:txtProxy.Size = New-Object System.Drawing.Size(150, 20); $global:txtProxy.Text = ""

# Buttons (Y=28)
$btnLoadXml = New-Object System.Windows.Forms.Button; $btnLoadXml.Location = New-Object System.Drawing.Point(940, 28); $btnLoadXml.Size = New-Object System.Drawing.Size(80, 25); $btnLoadXml.Text = "Load XML"
$btnSaveXml = New-Object System.Windows.Forms.Button; $btnSaveXml.Location = New-Object System.Drawing.Point(1030, 28); $btnSaveXml.Size = New-Object System.Drawing.Size(80, 25); $btnSaveXml.Text = "Save XML"

# UPDATED: Moved Help Button to the far right (1120) to utilize new width
$btnOrangeHelp = New-Object System.Windows.Forms.Button
$btnOrangeHelp.Location = New-Object System.Drawing.Point(1120, 28)
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
# UPDATED: Label changed from 'Text:' to 'Filter:'
$lblFilter = New-Object System.Windows.Forms.Label; $lblFilterY = $innerY + 5; 
$lblFilter.Location = New-Object System.Drawing.Point(535, $lblFilterY); $lblFilter.Size = New-Object System.Drawing.Size(35, 20); $lblFilter.Text = "Filter:"

# UPDATED: Textbox width increased by 30px (100 -> 130)
$txtFilter = New-Object System.Windows.Forms.TextBox; $txtFilterY = $innerY + 2; 
$txtFilter.Location = New-Object System.Drawing.Point(570, $txtFilterY); $txtFilter.Size = New-Object System.Drawing.Size(130, 20)

# Tag Filter Dropdown - UPDATED: Shifted X right by 30px (675 -> 705)
$lblFilterTag = New-Object System.Windows.Forms.Label; 
$lblFilterTag.Location = New-Object System.Drawing.Point(705, $lblFilterY); $lblFilterTag.Size = New-Object System.Drawing.Size(30, 20); $lblFilterTag.Text = "Tag:"

# UPDATED: Shifted X right by 30px (705 -> 735)
$global:cmbFilterTag = New-Object System.Windows.Forms.ComboBox; 
$global:cmbFilterTag.Location = New-Object System.Drawing.Point(735, $txtFilterY); $global:cmbFilterTag.Size = New-Object System.Drawing.Size(90, 20); $global:cmbFilterTag.DropDownStyle = "DropDownList"
$global:cmbFilterTag.Items.Add("All")
$global:cmbFilterTag.SelectedIndex = 0

# UPDATED: Shifted X right by 30px (800 -> 830)
$btnApplyFilter = New-Object System.Windows.Forms.Button; 
$btnApplyFilter.Location = New-Object System.Drawing.Point(830, $innerY); $btnApplyFilter.Size = New-Object System.Drawing.Size(80, 30); $btnApplyFilter.Text = "Apply Filter"; $btnApplyFilter.BackColor = "#A9A9A9"; $btnApplyFilter.ForeColor = "White"

# 4. Toggles & Columns (Shifted Right)
$btnToggleUnassigned = New-Object System.Windows.Forms.Button
# UPDATED: Shifted X right by 30px (900 -> 930)
$btnToggleUnassigned.Location = New-Object System.Drawing.Point(930, $innerY) 
$btnToggleUnassigned.Size = New-Object System.Drawing.Size(120, 30)
$btnToggleUnassigned.Text = "Hide Unassigned"
$btnToggleUnassigned.BackColor = "#e0e0e0"

$btnSelectCols = New-Object System.Windows.Forms.Button
# UPDATED: Shifted X right by 30px (1040 -> 1070)
$btnSelectCols.Location = New-Object System.Drawing.Point(1070, $innerY)
$btnSelectCols.Size = New-Object System.Drawing.Size(70, 30); $btnSelectCols.Text = "Columns"; $btnSelectCols.BackColor = "#808080"; $btnSelectCols.ForeColor = "White"

# Help Button
$btnHelp = New-Object System.Windows.Forms.Button
# UPDATED: Shifted X right by 30px (1130 -> 1160)
$btnHelp.Location = New-Object System.Drawing.Point(1160, $innerY)
$btnHelp.Size = New-Object System.Drawing.Size(60, 30)
$btnHelp.Text = "Help"
$btnHelp.BackColor = "#17a2b8"
$btnHelp.ForeColor = "White"

# --- Re-Add Logic Blocks ---
$actionApplyFilter = {
    if ($dataGridView.DataSource -is [System.Data.DataTable]) {
        $view = $dataGridView.DataSource.DefaultView
        $filterParts = New-Object System.Collections.ArrayList
        $rawVal = $txtFilter.Text.Trim()

        if (-not [string]::IsNullOrWhiteSpace($rawVal)) {
            $safeVal = $rawVal.Replace("'", "''").Replace("[", "[[]").Replace("*", "[*]").Replace("%", "[%]")
            $cols = $dataGridView.DataSource.Columns
            $textFilterParts = New-Object System.Collections.ArrayList
            foreach ($col in $cols) {
                [void]$textFilterParts.Add("[$($col.ColumnName)] LIKE '%$safeVal%'")
            }
            [void]$filterParts.Add("(" + ($textFilterParts -join " OR ") + ")")
        }

        if ($global:hideUnassigned) {
            [void]$filterParts.Add("(UserPrincipalName IS NOT NULL AND UserPrincipalName <> '')")
        }

        $selectedTag = $global:cmbFilterTag.SelectedItem
        if ($selectedTag -and $selectedTag -ne "All") {
            [void]$filterParts.Add("([Tag] LIKE '%$selectedTag%')")
        }

        if ($filterParts.Count -gt 0) {
            try { $view.RowFilter = $filterParts -join " AND " } catch {}
        } else {
            $view.RowFilter = ""
        }
        Update-Stats
    }
}

$btnApplyFilter.Add_Click($actionApplyFilter)
$txtFilter.Add_KeyDown({ param($s, $e)
    if ($e.KeyCode -eq [System.Windows.Forms.Keys]::Enter) { & $actionApplyFilter; $e.SuppressKeyPress = $true }
})
$global:cmbFilterTag.Add_SelectionChangeCommitted($actionApplyFilter)

$btnToggleUnassigned.Add_Click({
    $global:hideUnassigned = -not $global:hideUnassigned
    if ($global:hideUnassigned) {
        $btnToggleUnassigned.Text = "Show All"
        $btnToggleUnassigned.BackColor = "#b3d9ff"
    } else {
        $btnToggleUnassigned.Text = "Hide Unassigned"
        $btnToggleUnassigned.BackColor = "#e0e0e0"
    }
    & $actionApplyFilter
})

$btnGetFree.Add_Click({
    if ($dataGridView.Rows.Count -eq 0) { return }
    $startIndex = 0; if ($dataGridView.SelectedRows.Count -gt 0) { $startIndex = $dataGridView.SelectedRows[$dataGridView.SelectedRows.Count - 1].Index + 1 }
    $foundIndex = -1; for ($i = $startIndex; $i -lt $dataGridView.Rows.Count; $i++) { $upn = [string]$dataGridView.Rows[$i].Cells["UserPrincipalName"].Value; if ([string]::IsNullOrWhiteSpace($upn)) { $foundIndex = $i; break } }
    if ($foundIndex -eq -1) { for ($i = 0; $i -lt $startIndex; $i++) { $upn = [string]$dataGridView.Rows[$i].Cells["UserPrincipalName"].Value; if ([string]::IsNullOrWhiteSpace($upn)) { $foundIndex = $i; break } } }
    if ($foundIndex -ne -1) { $dataGridView.ClearSelection(); $dataGridView.Rows[$foundIndex].Selected = $true; $dataGridView.FirstDisplayedScrollingRowIndex = $foundIndex; Update-Stats } else { [System.Windows.Forms.MessageBox]::Show("No unassigned numbers found.", "Info") }
})

$btnHelp.Add_Click({
    $fHelp = New-Object System.Windows.Forms.Form
    $fHelp.Text = "Help"
    $fHelp.Size = New-Object System.Drawing.Size(800, 600)
    $fHelp.StartPosition = "CenterParent"

    $txtHelp = New-Object System.Windows.Forms.TextBox
    $txtHelp.Multiline = $true; $txtHelp.ReadOnly = $true
    $txtHelp.Location = New-Object System.Drawing.Point(10, 10)
    $txtHelp.Size = New-Object System.Drawing.Size(760, 490)
    $txtHelp.ScrollBars = "Vertical"; $txtHelp.BackColor = "White"
    $txtHelp.Font = New-Object System.Drawing.Font("Consolas", 10)

    $basePath = if ($PSScriptRoot) { $PSScriptRoot } else { $PWD.Path }
    $helpPath = Join-Path $basePath "help.txt"
    if (Test-Path $helpPath) {
        try { $txtHelp.Text = Get-Content $helpPath -Raw -Encoding UTF8 }
        catch { $txtHelp.Text = "Error: $($_.Exception.Message)" }
    } else {
        $txtHelp.Text = "help.txt not found."
    }

    $btnClose = New-Object System.Windows.Forms.Button
    $btnClose.Text = "Close"
    $btnClose.Location = New-Object System.Drawing.Point(350, 520)
    $btnClose.DialogResult = "OK"

    $fHelp.Controls.AddRange(@($txtHelp, $btnClose))
    $fHelp.Add_Shown({ $txtHelp.Select(0, 0); $btnClose.Focus() })
    $fHelp.ShowDialog()
})

$grpTopActions.Controls.AddRange(@($btnConnect, $btnFetchData, $btnSyncOrange, $sepTop1, $btnExport, $btnGetFree, $sepFilter, $lblFilter, $txtFilter, $lblFilterTag, $global:cmbFilterTag, $btnApplyFilter, $btnToggleUnassigned, $btnSelectCols, $btnHelp))

# -- Grid (Y Position shifted down to 160 to accommodate groupbox) --
$gridY = 160
$dataGridView = New-Object System.Windows.Forms.DataGridView
$dataGridView.Location = New-Object System.Drawing.Point(20, $gridY)
$dataGridView.Size = New-Object System.Drawing.Size(1240, 610)
$dataGridView.Anchor = "Top, Bottom, Left, Right"
$dataGridView.AllowUserToAddRows = $false
$dataGridView.SelectionMode = "FullRowSelect"
$dataGridView.MultiSelect = $true
$dataGridView.ReadOnly = $true
$dataGridView.AutoSizeColumnsMode = "AllCells"
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

$grpTag.Controls.AddRange(@($lblTagInput, $global:cmbTag, $cbBlacklist, $cbReserved, $cbPremium, $btnApplyTag, $btnRemoveTags, $sepAction1, $btnAssign, $btnUnassign, $sepAction2, $btnRemove, $sepAction3, $btnReleaseOC, $btnPublishOC, $btnManualPublish))
#endregion

# =============================================================================
# LOGIC
# =============================================================================

#region 9. Logic - General UI
$btnSelectCols.Add_Click({
    if ($dataGridView.DataSource -eq $null) {
        [System.Windows.Forms.MessageBox]::Show("Please load data first.", "Info"); return
    }

    $frmCols = New-Object System.Windows.Forms.Form
    $frmCols.Text = "Select Columns"
    $frmCols.Size = New-Object System.Drawing.Size(300, 500)
    $frmCols.StartPosition = "CenterParent"
    $frmCols.FormBorderStyle = "FixedDialog"

    $clb = New-Object System.Windows.Forms.CheckedListBox
    $clb.Location = New-Object System.Drawing.Point(10, 10)
    $clb.Size = New-Object System.Drawing.Size(260, 400)
    $clb.CheckOnClick = $true

    foreach ($col in $dataGridView.Columns) {
        $state = if ($col.Visible) { [System.Windows.Forms.CheckState]::Checked } else { [System.Windows.Forms.CheckState]::Unchecked }
        [void]$clb.Items.Add($col.Name, $state)
    }

    $btnOkCols = New-Object System.Windows.Forms.Button
    $btnOkCols.Text = "OK"
    $btnOkCols.DialogResult = "OK"
    $btnOkCols.Location = New-Object System.Drawing.Point(100, 420)
    $frmCols.Controls.Add($clb)
    $frmCols.Controls.Add($btnOkCols)
    $frmCols.AcceptButton = $btnOkCols

    if ($frmCols.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $global:form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        $dataGridView.SuspendLayout()
        for ($i = 0; $i -lt $clb.Items.Count; $i++) {
            $colName = $clb.Items[$i]
            $isVisible = $clb.GetItemChecked($i)
            if ($colName -eq "TelephoneNumber") { $isVisible = $true }
            $dataGridView.Columns[$colName].Visible = $isVisible
        }
        $dataGridView.ResumeLayout()
        $global:form.Cursor = [System.Windows.Forms.Cursors]::Default
    }
    $frmCols.Dispose()
})
#endregion

#region 10. Logic - Connection & Fetching
# UPDATED CONNECT BUTTON LOGIC (User Request: No Minimize)
$btnConnect.Add_Click({
    try {
        # UPDATED: Immediate feedback
        Write-Log "Authenticating..."
        
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
        $users = Get-CsOnlineUser -ResultSize 20000 -ErrorAction Stop
        
        Write-Log "  > Found $($users.Count) users. Building Index..."
        $global:teamsUsersMap = @{}
        foreach ($u in $users) { if ($u.Identity) { $global:teamsUsersMap[$u.Identity] = $u } }

        # 2. GET NUMBERS
        Write-Log "Step 3/5: Fetching Phone Numbers (Batched)..."
        Update-ProgressUI -Current 30 -Total 100 -Activity "Fetch Numbers"
        $allNumbers = New-Object System.Collections.ArrayList
        $batchSize = 1000; $skip = 0
        while ($true) {
            $batch = Get-CsPhoneNumberAssignment -Skip $skip -Top $batchSize -ErrorAction Stop
            if (!$batch) { break }
            [void]$allNumbers.AddRange($batch)
            $skip += $batchSize
            Write-Log "  > Fetched batch. Total so far: $($allNumbers.Count)"
            $pct = [Math]::Min(55, 30 + [int]($allNumbers.Count / 100))
            Update-ProgressUI -Current $pct -Total 100 -Activity "Fetching Numbers ($($allNumbers.Count))"
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

                if ($orangeData) { Merge-OrangeData -OrangeData $orangeData }
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
        Merge-OrangeData -OrangeData $orangeData

        $dataGridView.DataSource = $global:masterDataTable
        foreach ($hc in $global:defaultHiddenCols) { if ($dataGridView.Columns[$hc]) { $dataGridView.Columns[$hc].Visible = $false } }

        Update-FilterTags
        Update-Stats
        Update-TagStatistics
        Update-ProgressUI -Current 100 -Total 100 -Activity "Done"
        Reset-ProgressUI
        Write-Log "Sync Complete."
    }
})
#endregion

#region 11. Logic - Context Menu Actions
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
            $numData = Get-CsPhoneNumberAssignment -TelephoneNumber $ph -ErrorAction Stop
            
            if ($numData) {
                $row.Cells["ActivationState"].Value = $numData.ActivationState; $row.Cells["NumberType"].Value = $numData.NumberType; $row.Cells["City"].Value = $numData.City
                $row.Cells["IsoCountryCode"].Value = $numData.IsoCountryCode; $row.Cells["IsoSubdivision"].Value = $numData.IsoSubdivision; $row.Cells["NumberSource"].Value = $numData.NumberSource
                $row.Cells["Tag"].Value = if ($numData.Tag) { $numData.Tag -join ", " } else { "" }

                $userId = $numData.AssignedPstnTargetId
                if (-not [string]::IsNullOrWhiteSpace($userId)) {
                    try {
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
            $safeInput = $inputUser.Replace("'", "''")
            $adUser = Get-ADUser -Filter "SamAccountName -eq '$safeInput'" -Properties UserPrincipalName -ErrorAction Stop
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
        Set-CsPhoneNumberAssignment -Identity $upn -PhoneNumber $ph -PhoneNumberType $type -ErrorAction Stop; Write-Log "Teams Assignment Successful."
        $row.Cells["UserPrincipalName"].Value = $userObj.UserPrincipalName; $row.Cells["DisplayName"].Value = $userObj.DisplayName; $row.Cells["ActivationState"].Value = "Assigned"
        Write-Log "Syncing to On-Prem AD..."
        try {
            $safeUpn = $upn.Replace("'", "''")
            $adUser = Get-ADUser -Filter "UserPrincipalName -eq '$safeUpn'" -ErrorAction Stop
            if ($adUser) {
                Set-ADUser -Identity $adUser -OfficePhone $ph -ErrorAction Stop
                Write-Log "AD OfficePhone updated for $upn."
            }
        } catch {
            Write-Log "AD Sync Warning: $($_.Exception.Message)"
        }
        Update-TagStatistics
    } catch {
        Write-Log "Assignment Failed: $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show("Assignment Failed: $($_.Exception.Message)", "Error")
    } finally {
        $global:form.Cursor = [System.Windows.Forms.Cursors]::Default
    }
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
                Remove-CsPhoneNumberAssignment -Identity $upn -PhoneNumber $phone -PhoneNumberType $numType -ErrorAction Stop; Write-Log "Unassigned $phone."
                
                # NEW: Try to sync to AD (Clear OfficePhone)
                Write-Log "Syncing to On-Prem AD (Clearing OfficePhone)..."
                try {
                    $safeUpn = $upn.Replace("'", "''")
                    $adUser = Get-ADUser -Filter "UserPrincipalName -eq '$safeUpn'" -ErrorAction Stop
                    if ($adUser) {
                        Set-ADUser -Identity $adUser -Clear "telephoneNumber" -ErrorAction Stop
                        Write-Log "AD OfficePhone cleared for $upn."
                    }
                } catch {
                    Write-Log "AD Sync Warning: $($_.Exception.Message)"
                }

            } catch {
                $errMsg = $_.Exception.Message
                if ($errMsg -match "on-premises Active Directory" -or $errMsg -match "synchronized to the cloud") {
                    Write-Log "Detected On-Prem user. Clearing AD attributes..."
                    try {
                        Set-ADUser -Identity $upn -Clear "msRTCSIP-Line", "telephoneNumber" -ErrorAction Stop
                        Write-Log "Success: AD attributes cleared."
                    } catch {
                        Write-Log "Failed to clear AD attributes: $($_.Exception.Message)."
                    }
                } else {
                    Write-Log "Failed unassign: ${phone}: $errMsg"
                }
                continue
            } }
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

$dataGridView.Add_CellFormatting({ param($s, $e)
    if ($e.ColumnIndex -ge 0 -and $dataGridView.Columns[$e.ColumnIndex].Name -match "^Orange") {
        $e.CellStyle.BackColor = [System.Drawing.Color]::Bisque
    }
})

$dataGridView.Add_CellDoubleClick({ param($s, $e)
    if ($e.RowIndex -ge 0) {
        $p = $dataGridView.Rows[$e.RowIndex].Cells["TelephoneNumber"].Value
        Write-Log "-- $p --"
        if ($global:orangeHistoryMap[$p]) {
            $global:orangeHistoryMap[$p] | ForEach-Object { Write-Log "$($_.date) | $($_.status)" }
        }
    }
})
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
        $tagsToRemove = @()
        $tagsToAdd = @()

        # Determine location tag swap
        $newLocationTag = $desiredTags | Where-Object { $global:allowedTags -contains $_ } | Select-Object -First 1
        if ($newLocationTag) {
            $oldLocationTag = $currentTags | Where-Object { $global:allowedTags -contains $_ } | Select-Object -First 1
            if ($oldLocationTag -and $oldLocationTag -ne $newLocationTag) { $tagsToRemove += $oldLocationTag }
            if (-not ($currentTags -contains $newLocationTag)) { $tagsToAdd += $newLocationTag }
        }

        # Determine special tag changes
        $specialTags = @("Blacklist", "Reserved", "Premium")
        foreach ($st in $specialTags) {
            $isDesired = $desiredTags -contains $st
            $isCurrent = $currentTags -contains $st
            if ($isCurrent -and -not $isDesired) { $tagsToRemove += $st }
            if ($isDesired -and -not $isCurrent) { $tagsToAdd += $st }
        }

        # Apply tag changes via API
        try {
            foreach ($t in $tagsToRemove) {
                Write-Log "Removing tag '$t' from $ph..."
                Remove-CsPhoneNumberTag -PhoneNumber $ph -Tag $t -ErrorAction Stop
            }
            foreach ($t in $tagsToAdd) {
                Write-Log "Adding tag '$t' to $ph..."
                Set-CsPhoneNumberTag -PhoneNumber $ph -Tag $t -ErrorAction Stop
            }
            $finalTags = $currentTags | Where-Object { -not ($tagsToRemove -contains $_) }
            $finalTags += $tagsToAdd
            $r.Cells["Tag"].Value = ($finalTags | Sort-Object | Get-Unique) -join ", "
            if ($tagsToAdd.Count -eq 0 -and $tagsToRemove.Count -eq 0) {
                Write-Log "No tag changes needed for $ph."
            } else {
                Write-Log "Tags updated for $ph."
            }
        } catch {
            Write-Log "Failed to update tags for ${ph}: $($_.Exception.Message)"
        }
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
                    Remove-CsPhoneNumberAssignment -Identity $u -PhoneNumber $p -PhoneNumberType $t -ErrorAction Stop;
                    Write-Log "Unassigned $p."
                    
                    # NEW: Try to sync to AD (Clear OfficePhone)
                    Write-Log "Syncing to On-Prem AD (Clearing OfficePhone)..."
                    try {
                        $safeU = $u.Replace("'", "''")
                        $adUser = Get-ADUser -Filter "UserPrincipalName -eq '$safeU'" -ErrorAction Stop
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
                            $safeU = $u.Replace("'", "''")
                            Get-ADUser -Filter "UserPrincipalName -eq '$safeU'" -Properties msRTCSIP-Line, telephoneNumber -ErrorAction Stop | Set-ADUser -Clear "msRTCSIP-Line", "telephoneNumber" -ErrorAction Stop;
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
        $toRemove = New-Object System.Collections.ArrayList; Write-Log "Removing $($rowsToProcess.Count) number(s)..."
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
                        Remove-CsPhoneNumberAssignment -Identity $u -PhoneNumber $p -PhoneNumberType $t -ErrorAction Stop
                    } catch { 
                        $errMsg = $_.Exception.Message; 
                        if ($errMsg -match "on-premises Active Directory" -or $errMsg -match "synchronized to the cloud") { 
                            Write-Log "Detected On-Prem user. Clearing AD attributes..."
                            try {
                                $safeU = $u.Replace("'", "''")
                                Get-ADUser -Filter "UserPrincipalName -eq '$safeU'" -Properties msRTCSIP-Line, telephoneNumber -ErrorAction Stop | Set-ADUser -Clear "msRTCSIP-Line", "telephoneNumber" -ErrorAction Stop;
                                Write-Log "Success: AD attributes cleared."
                            } catch { Write-Log "Failed to clear AD attributes: $($_.Exception.Message)." }
                        } else { throw $_ }
                    }
                }; 
                Remove-CsOnlineTelephoneNumber -TelephoneNumber ([string[]]@($p)) -ErrorAction Stop; 
                [void]$toRemove.Add($r);
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

#region 15. Final Assembly & Execution
# --- Render ---
# REPLACE THE LAST FORM.CONTROLS LINE WITH THIS:
$global:form.Controls.AddRange(@($grpConfig, $grpTopActions, $dataGridView, $grpLog, $grpStats, $grpTag, $grpTagStats, $progressBar))
$global:form.Add_Shown({ 
    $global:form.Activate()
    # Auto-load default XML if present
    if (Test-Path $global:defaultXmlPath) {
        Import-SettingsXml -Path $global:defaultXmlPath
    }
    Write-Log "Ready." 
})
[void]$global:form.ShowDialog()
$global:form.Dispose()
#endregion
