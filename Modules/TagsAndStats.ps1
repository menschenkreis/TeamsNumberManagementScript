# =============================================================================
# TAGS AND STATS MODULE - Tag Statistics and Filtering
# =============================================================================

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
        $free = $stats[$tagKey].Free

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
