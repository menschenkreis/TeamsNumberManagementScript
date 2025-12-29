# =============================================================================
# TEAMS PHONE MANAGER v56.2 (Modular Edition)
# =============================================================================
# Main script that loads all modules
# =============================================================================

# Determine script location
$ScriptRoot = if ($PSScriptRoot) { $PSScriptRoot } else { $PWD.Path }
$ModulesPath = Join-Path $ScriptRoot "Modules"

Write-Host "==============================================================================" -ForegroundColor Cyan
Write-Host "TEAMS PHONE MANAGER v56.2 (Modular Edition)" -ForegroundColor Cyan
Write-Host "==============================================================================" -ForegroundColor Cyan
Write-Host "Loading modules from: $ModulesPath" -ForegroundColor Yellow
Write-Host ""

# Load modules in order
$modules = @(
    "Core.ps1"
    "HelperFunctions.ps1"
    "SettingsManagement.ps1"
    "TagsAndStats.ps1"
    "ApiIntegration.ps1"
    "UIComponents.ps1"
    "TeamsOperations.ps1"
    "ContextMenu.ps1"
    "Actions.ps1"
)

foreach ($module in $modules) {
    $modulePath = Join-Path $ModulesPath $module
    if (Test-Path $modulePath) {
        Write-Host "[Loading] $module..." -ForegroundColor Green
        . $modulePath
    } else {
        Write-Host "[ERROR] Module not found: $module" -ForegroundColor Red
        Write-Host "Expected path: $modulePath" -ForegroundColor Red
        exit 1
    }
}

Write-Host ""
Write-Host "All modules loaded successfully!" -ForegroundColor Green
Write-Host "==============================================================================" -ForegroundColor Cyan
Write-Host ""

#region Final Assembly & Execution
# --- Render ---
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
