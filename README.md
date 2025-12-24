# Teams Phone Manager

**Teams Phone Manager** is a comprehensive GUI-based PowerShell tool designed to streamline the management of Microsoft Teams Phone numbers. It integrates directly with the **Orange Business Talk (OC)** API to synchronize inventory status between Microsoft Teams and the Orange Cloud.

## Features

* **Inventory Management:** Track number stock, assignment status, and location tags in a unified grid.
* **User Assignment:** Assign/Unassign numbers using UPN or SamAccountName.
* **AD Synchronization:** Automatically updates `telephoneNumber` and `OfficePhone` attributes in On-Premises Active Directory when numbers are assigned or removed.
* **Orange Business Talk Integration:** Publish, Release, and Sync number statuses directly with the Orange API.
* **Policy Management:** Bulk or single assignment of Voice Routing and Meeting Policies.
* **Visual Alerts:** Color-coded tracking for low number stock based on location tags.

---

## Prerequisites

To use this tool effectively, ensure the following requirements are met:

### 1. Permissions
* **Microsoft Teams:** You must hold at least the **Teams Communication Administrator** role.
* **Active Directory:** To use the AD Sync features (updating user phone attributes), the account running the script must have **Write/Update permissions** on user objects in the On-Premises Active Directory.

### 2. PowerShell Modules
* **MicrosoftTeams:** Required. The script will attempt to install this if missing.
* **ActiveDirectory:** Required only if you wish to use SamAccountName lookup or AD attribute syncing.

---

## Installation & Usage

1.  **Download:** Download the `.ps1` script file from this repository.
2.  **Run:** Execute the script using PowerShell.
    ```powershell
    .\TeamsPhoneManager_v56.4.ps1
    ```

### Configuration (Settings.xml)
While the script can run standalone, it is **highly recommended** to create a `Settings.xml` file in the same directory as the script. This allows you to persist API credentials, proxy settings, and tag lists.

**Example `Settings.xml`:**
```xml
<Settings>
  <OrangeCustomerID>YOUR_CUSTOMER_ID</OrangeCustomerID>
  <OrangeAuthHeader>Basic YOUR_BASE64_AUTH_HEADER</OrangeAuthHeader>
  <OrangeApiKey>YOUR_API_KEY</OrangeApiKey>
  <Proxy>[http://proxy.address:8080](http://proxy.address:8080)</Proxy>

  <LowStockAlertThreshold>5</LowStockAlertThreshold>
  <SelectTagList>London, Paris, New York, HQ, Remote</SelectTagList>
</Settings>
```

> **Note:** If you do not provide an XML file, you can enter credentials manually within the application GUI. However, these settings will not be saved after the application closes.

## User Guide

### 1. Workflow Steps
* **Connect:** Click to authenticate with the Microsoft Teams PowerShell module.
* **Get Data:** Fetch Voice/Meeting policies, Teams Users, and Phone Numbers.
    * *Note:* If Orange API keys are populated, the tool will also fetch numbers from Orange Cloud and merge the datasets.
* **Re-Sync (Optional):** Refresh Orange statuses without re-downloading the entire Teams dataset.

### 2. Main Grid & Search
* **Filtering:**
    * **Text Box:** Search by Number, User, City, etc. Supports wildcards (`*`).
    * **Tag Filter:** Filter by specific location tags.
    * **Hide Unassigned:** Toggles visibility of empty/available numbers.
* **Columns:** Click to show/hide specific data fields.
* **Export:** Select rows and click **Export** to save the selection as a CSV file.
* **Get Free Number:** Automatically finds and highlights the next available unassigned number in the current view.

### 3. Actions (Bottom Panel)

#### A. Tagging (Smart Tags)
Select a **Location Tag** or check flags (**Blacklist**, **Reserved**, **Premium**) and click **Apply Tag**.
* **Logic:** Replaces existing location tags with the new selection but preserves other flags (e.g., Reserved) unless explicitly unchecked.

#### B. Teams Assignment
* **Assign:** Assigns the selected number to a user.
    * *Input:* UserPrincipalName (email) OR SamAccountName.
    * *Action:* Updates `OfficePhone` in On-Premises AD.
* **Unassign:** Removes the user from the number in Teams.
    * *Action:* Attempts to clear `telephoneNumber` in On-Premises AD.

#### C. Removal
* **Remove:** Permanently removes the number from your Teams Tenant.

#### D. Orange Cloud Actions
* **Release from OC:** Unassigns the user in Teams **and** releases the number batch in Orange Cloud.
* **Publish to OC:** Publishes the selected range to Orange Cloud.
* **Manual Publish:** Allows manual entry of a number range to publish to a specific Voice Site.

### 4. Context Menu (Right-Click)
* **Refresh Orange/Teams Info:** Updates data for just the selected row(s).
* **Change Voice Routing Policy:** Grants a new voice policy to the assigned user.
* **Grant Meeting Policy:** Grants a new meeting policy to the assigned user.
* **Enable Enterprise Voice:** Toggles the EV flag for the user.
* **Force Publish to OC:** Forces a publish command for a specific site ID.

### 5. Statistics & Logging
* **Tag Statistics:** Displays **Total** vs. **Free** numbers per location tag.
    * 🟠 **Orange:** Low Stock (Warning).
    * 🔴 **Red:** Critical Stock.
* **Logging:** A comprehensive log is displayed in the GUI and saved automatically to the `\ExecutionLogs` folder.

---

### Orange Business Talk API
For details regarding the Orange Business Talk API required for the integration features, please refer to the official documentation:
[Orange Developer: Business Talk API Getting Started](https://developer.orange.com/apis/businesstalk-fr/getting-started)

---

### Disclaimer
> This script is provided "as-is". Please test thoroughly in a non-production environment before using it to manage live numbers. The author is not responsible for any data loss or configuration errors resulting from the use of this tool.
