# =============================================================================
# CORE MODULE - Global Variables, Assemblies, and Logging Setup
# =============================================================================

#region Cleanup & Assemblies
# --- CLEANUP SECTION ---
if ($global:form -and !$global:form.IsDisposed) {
    $global:form.Close()
    $global:form.Dispose()
}

# Load Assemblies
try { Add-Type -AssemblyName System.Windows.Forms } catch {}
try { Add-Type -AssemblyName System.Drawing } catch {}
try { Add-Type -AssemblyName System.Data } catch {}
try { Add-Type -AssemblyName Microsoft.VisualBasic } catch {}
#endregion

#region Global Variables & Logging
# --- Global Data ---
$global:teamsUsersMap = @{}
$global:orangeHistoryMap = @{}
$global:masterDataTable = $null
$global:form = $null
$global:voiceRoutingPolicies = @()
$global:teamsMeetingPolicies = @()
$global:hideUnassigned = $false

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
    # Go up one level from Modules directory to get the main script directory
    $mainScriptPath = Split-Path -Parent $scriptPath
    $logDir = Join-Path $mainScriptPath "ExecutionLogs"

    # Create Directory
    if (-not (Test-Path $logDir)) {
        New-Item -ItemType Directory -Path $logDir -Force | Out-Null
    }

    # Generate Filename
    $logName = "Log_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
    $global:logFilePath = Join-Path $logDir $logName

    # Set Default XML Path
    $global:defaultXmlPath = Join-Path $mainScriptPath "Settings.xml"

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

#region UI References & Column Definitions
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
