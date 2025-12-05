
<#
.SYNOPSIS
  Export Microsoft Purview Information Protection (labels & label policies)
  and Data Loss Prevention (DLP policies & rules), then generate Word and
  PowerPoint documentation.

.PARAMETER OutputFolder
  Folder to write JSON/CSV/Word/PowerPoint reports.

.PARAMETER TenantDisplayName
  Friendly tenant name for report titles.

.PARAMETER IncludeGraphEnrichment
  Optional. If specified, attempts Microsoft Graph (beta) enrichment for IP/DLP.

.PARAMETER CreateWord
  Optional. Generate Word (.docx) report via COM.

.PARAMETER CreatePowerPoint
  Optional. Generate PowerPoint (.pptx) report via COM.

.PARAMETER CreateMarkdown
  Optional. Generate Markdown (.md) report with portal links.

.NOTES
  - Requires ExchangeOnlineManagement module and IPPSSession for Compliance Center.
  - Word/PPT generation uses COM automation (Windows + Office required).
  - Markdown generation is cross-platform compatible.
  - Graph enrichment uses Microsoft.Graph and beta endpoints with admin consent.

#>

param(
  [Parameter(Mandatory=$true)]
  [string]$OutputFolder,
  [Parameter(Mandatory=$true)]
  [string]$TenantDisplayName,
  [switch]$IncludeGraphEnrichment,
  [switch]$CreateWord,
  [switch]$CreatePowerPoint,
  [switch]$CreateMarkdown
)

# ---------------------------
# Helpers & Setup
# ---------------------------
$ErrorActionPreference = 'Stop'
$null = New-Item -ItemType Directory -Path $OutputFolder -Force

function Write-Info($msg) { Write-Host "[INFO] $msg" -ForegroundColor Cyan }
function Write-Warn($msg) { Write-Host "[WARN] $msg" -ForegroundColor Yellow }
function Write-Err($msg)  { Write-Host "[ERROR] $msg" -ForegroundColor Red }

$timestamp = (Get-Date).ToString("yyyy-MM-dd_HH-mm-ss")
$reportBase = Join-Path $OutputFolder "PurviewDoc_$($timestamp)"
$paths = @{
  LabelsJson           = "$reportBase.labels.json"
  LabelPoliciesJson    = "$reportBase.labelPolicies.json"
  LabelsCsv            = "$reportBase.labels.csv"
  LabelPoliciesCsv     = "$reportBase.labelPolicies.csv"
  DlpPoliciesJson      = "$reportBase.dlpPolicies.json"
  DlpRulesJson         = "$reportBase.dlpRules.json"
  DlpPoliciesCsv       = "$reportBase.dlpPolicies.csv"
  DlpRulesCsv          = "$reportBase.dlpRules.csv"
  WordDoc              = "$reportBase.docx"
  PowerPoint           = "$reportBase.pptx"
  MarkdownDoc          = "$reportBase.md"
  MetaJson             = "$reportBase.meta.json"
}

$meta = [ordered]@{
  Tenant               = $TenantDisplayName
  GeneratedOnUtc       = (Get-Date).ToUniversalTime().ToString("o")
  Tools                = @('IPPSSession (Compliance PowerShell)')
  GraphEnrichment      = [bool]$IncludeGraphEnrichment
  Sections             = @('Sensitivity Labels','Label Policies','Auto-Labeling Policies','DLP Policies','DLP Rules')
}

# ---------------------------
# Connect to IPPSSession (Compliance Center)
# ---------------------------
Write-Info "Connecting to Compliance PowerShell (IPPSSession)..."
Import-Module ExchangeOnlineManagement -ErrorAction Stop

try {
  # This opens modern auth prompt
  Connect-IPPSSession | Out-Null
  Write-Info "Connected to IPPSSession."
}
catch {
  Write-Err "Failed to connect to IPPSSession. $_"
  throw
}

# ---------------------------
# Collect Information Protection: Labels & Label Policies
# ---------------------------
Write-Info "Retrieving Sensitivity Labels..."
$labels = @()
try {
  # Get-Label returns label metadata defined in Purview
  $labels = Get-Label -ResultSize Unlimited
  Write-Info "Found $($labels.Count) labels."
}
catch {
  Write-Err "Get-Label failed. $_"
}

Write-Info "Retrieving Label Policies..."
$labelPolicies = @()
try {
  # Published label policies (who sees what labels)
  $labelPolicies = Get-LabelPolicy -ResultSize Unlimited
  Write-Info "Found $($labelPolicies.Count) label policies."
}
catch {
  Write-Err "Get-LabelPolicy failed. $_"
}

# Normalize labels for export
$labelsProjection = $labels | ForEach-Object {
  $encryptionEnabled = $false
  $encryptionSummary = "None"
  if ($_.Encryption -and $_.Encryption.EncryptionEnabled) {
    $encryptionEnabled = $true
    $encryptionSummary = "Template-based RMS"
    if ($_.Encryption.EncryptionRightsDefinitions) {
      $encryptionSummary += " ($($_.Encryption.EncryptionRightsDefinitions.Count) permission assignments)"
    }
  }
  
  $parentLabel = $null
  if ($_.ParentId) {
    $parent = $labels | Where-Object { $_.ImmutableId -eq $_.ParentId }
    if ($parent) { $parentLabel = $parent.Name }
  }
  
  [PSCustomObject]@{
    Name               = $_.Name
    Guid               = $_.ImmutableId
    ParentLabelName    = $parentLabel
    ParentLabelId      = $_.ParentId
    Priority           = $_.Priority
    Enabled            = $_.Enabled
    Tooltip            = $_.Tooltip
    Description        = $_.Comment
    EncryptionEnabled  = $encryptionEnabled
    EncryptionSummary  = $encryptionSummary
    ContentMarking     = ($_.ContentMarking   | ConvertTo-Json -Depth 6 -Compress)
    Encryption         = ($_.Encryption       | ConvertTo-Json -Depth 6 -Compress)
    EndpointProtection = ($_.EndpointProtection | ConvertTo-Json -Depth 6 -Compress)
    AutoLabeling       = ($_.AutoLabeling     | ConvertTo-Json -Depth 6 -Compress)
    LocaleSettings     = ($_.LocaleSettings   | ConvertTo-Json -Depth 6 -Compress)
    ScopeSettings      = ($_.Settings         | ConvertTo-Json -Depth 6 -Compress)
    ModifiedTime       = $_.WhenChanged
    CreatedTime        = $_.WhenCreated
    PortalUrl          = "https://compliance.microsoft.com/informationprotection/labels/$($_.ImmutableId)"
  }
}

# Normalize label policies for export
$labelPoliciesProjection = $labelPolicies | ForEach-Object {
  [PSCustomObject]@{
    Name                 = $_.Name
    Guid                 = $_.ImmutableId
    Enabled              = $_.Enabled
    Priority             = $_.Priority
    PublisherType        = $_.PublisherType
    ApplyTo              = ($_.ApplyTo | ConvertTo-Json -Depth 6)      # users/groups
    LabelsPublished      = ($_.Labels | ConvertTo-Json -Depth 6)       # label references
    AdvancedSettings     = ($_.AdvancedSettings | ConvertTo-Json -Depth 6)
    Mode                 = $_.Mode
    CreatedTime          = $_.WhenCreated
    ModifiedTime         = $_.WhenChanged
  }
}

# ---------------------------
# Collect Auto-Labeling Policies (Service-Side)
# ---------------------------
Write-Info "Retrieving Auto-Labeling Policies (service-side)..."
$autoLabelPolicies = @()
$autoLabelRules = @()
try {
  $autoLabelPolicies = Get-AutoSensitivityLabelPolicy -ErrorAction Stop
  Write-Info "Found $($autoLabelPolicies.Count) auto-labeling policies."
  
  if ($autoLabelPolicies.Count -gt 0) {
    $autoLabelRules = Get-AutoSensitivityLabelRule -ErrorAction Stop
    Write-Info "Found $($autoLabelRules.Count) auto-labeling rules."
  }
}
catch {
  Write-Warn "Auto-labeling policy retrieval failed (may not be configured). $_"
}

# Normalize auto-label policies
$autoLabelPoliciesProjection = $autoLabelPolicies | ForEach-Object {
  [PSCustomObject]@{
    Name                = $_.Name
    Guid                = $_.Identity
    Mode                = $_.Mode
    Comment             = $_.Comment
    ApplySensitivityLabel = $_.ApplySensitivityLabel
    Locations           = ($_.Locations | ConvertTo-Json -Depth 6 -Compress)
    ExchangeLocation    = ($_.ExchangeLocation | ConvertTo-Json -Depth 6 -Compress)
    SharePointLocation  = ($_.SharePointLocation | ConvertTo-Json -Depth 6 -Compress)
    OneDriveLocation    = ($_.OneDriveLocation | ConvertTo-Json -Depth 6 -Compress)
    Priority            = $_.Priority
    CreatedTime         = $_.WhenCreated
    ModifiedTime        = $_.WhenChanged
  }
}

# Normalize auto-label rules
$autoLabelRulesProjection = $autoLabelRules | ForEach-Object {
  [PSCustomObject]@{
    Name             = $_.Name
    Guid             = $_.Identity
    Policy           = $_.Policy
    Conditions       = ($_.Conditions | ConvertTo-Json -Depth 10 -Compress)
    CreatedTime      = $_.WhenCreated
    ModifiedTime     = $_.WhenChanged
  }
}

# ---------------------------
# Collect DLP Policies & Rules
# ---------------------------
Write-Info "Retrieving DLP Policies..."
$dlpPolicies = @()
try {
  # Modern cmdlets:
  $dlpPolicies = Get-DlpCompliancePolicy -ErrorAction Stop
  Write-Info "Found $($dlpPolicies.Count) DLP policies."
}
catch {
  Write-Warn "Get-DlpCompliancePolicy failed. Attempting legacy Get-DlpPolicy..."
  try {
    $dlpPolicies = Get-DlpPolicy -ErrorAction Stop
    Write-Info "Found $($dlpPolicies.Count) (legacy) DLP policies."
  }
  catch {
    Write-Err "DLP policy retrieval failed. $_"
  }
}

Write-Info "Retrieving DLP Rules..."
$dlpRules = @()
try {
  $dlpRules = Get-DlpComplianceRule -ErrorAction Stop
  Write-Info "Found $($dlpRules.Count) DLP rules."
}
catch {
  Write-Warn "Get-DlpComplianceRule failed. Attempting legacy Get-TransportRule..."
  try {
    # Fallback only partially relevant; DLP rules are not transport rules,
    # included here for completeness where tenants use legacy mail flow.
    $dlpRules = Get-TransportRule -ErrorAction Stop
    Write-Info "Found $($dlpRules.Count) transport rules (legacy fallback)."
  }
  catch {
    Write-Err "DLP rule retrieval failed. $_"
  }
}

# Flatten DLP policy projection
$dlpPoliciesProjection = $dlpPolicies | ForEach-Object {
  [PSCustomObject]@{
    Name            = $_.Name
    Guid            = $_.Identity
    Enabled         = $_.Enabled
    Mode            = $_.Mode  # test/enable etc.
    Comment         = $_.Comment
    Workload        = ($_.Workload    | ConvertTo-Json -Depth 6)   # Exchange/SharePoint/OneDrive/Teams
    ExchangeLocation = ($_.ExchangeLocation | ConvertTo-Json -Depth 6)
    SharePointLocation = ($_.SharePointLocation | ConvertTo-Json -Depth 6)
    OneDriveLocation   = ($_.OneDriveLocation | ConvertTo-Json -Depth 6)
    TeamsLocation      = ($_.TeamsLocation    | ConvertTo-Json -Depth 6)
    Exclusions         = ($_.ExchangeSenderNotify | ConvertTo-Json -Depth 6)
    CreatedTime        = $_.WhenCreated
    ModifiedTime       = $_.WhenChanged
  }
}

# Flatten DLP rules projection (conditions/actions as JSON strings for readability)
$dlpRulesProjection = $dlpRules | ForEach-Object {
  $conditionsJson = $null
  $exceptionsJson = $null
  $actionsJson    = $null

  if ($_.Conditions) { $conditionsJson = ($_.Conditions | ConvertTo-Json -Depth 10) }
  if ($_.Exceptions) { $exceptionsJson = ($_.Exceptions | ConvertTo-Json -Depth 10) }
  if ($_.Actions)    { $actionsJson    = ($_.Actions    | ConvertTo-Json -Depth 10) }

  [PSCustomObject]@{
    Name             = $_.Name
    Guid             = $_.Identity
    Policy           = $_.Policy  # policy association
    Enabled          = $_.Enabled
    Priority         = $_.Priority
    Mode             = $_.Mode
    Conditions       = $conditionsJson
    Exceptions       = $exceptionsJson
    Actions          = $actionsJson
    CreatedTime      = $_.WhenCreated
    ModifiedTime     = $_.WhenChanged
  }
}

# ---------------------------
# Optional: Graph enrichment (beta)
# ---------------------------
if ($IncludeGraphEnrichment) {
  Write-Info "Attempting Microsoft Graph enrichment (beta)..."
  try {
    Import-Module Microsoft.Graph -ErrorAction Stop
    # Connect with scopes—admin consent required
    Connect-MgGraph -Scopes @(
      "SecurityEvents.Read.All",
      "Policy.Read.All"
    ) | Out-Null

    # Use beta profile
    Select-MgProfile -Name "beta"

    # NOTE: Endpoints subject to change; wrap in try/catch
    # Example endpoints (illustrative; may vary by GA state):
    # GET /security/informationProtection/policyLabels
    # GET /security/dataLossPreventionPolicies
    $graphLabels   = @()
    $graphDlpPols  = @()

    try {
      $graphLabels = Invoke-MgGraphRequest -Method GET -Uri "/security/informationProtection/policyLabels"
      Write-Info "Graph labels enrichment returned: $($graphLabels.value.Count)"
    } catch { Write-Warn "Graph labels endpoint failed/enforced. $_" }

    try {
      $graphDlpPols = Invoke-MgGraphRequest -Method GET -Uri "/security/dataLossPreventionPolicies"
      Write-Info "Graph DLP policies enrichment returned: $($graphDlpPols.value.Count)"
    } catch { Write-Warn "Graph DLP endpoint failed/enforced. $_" }

    $meta.GraphNotes = "Graph enrichment attempted (beta). Some endpoints may be preview/subject to change."
    Disconnect-MgGraph | Out-Null
  }
  catch {
    Write-Warn "Microsoft Graph enrichment not available. $_"
  }
}

# ---------------------------
# Persist: JSON & CSV
# ---------------------------
Write-Info "Writing JSON and CSV outputs to $OutputFolder..."

# JSON
$labelsProjection        | ConvertTo-Json -Depth 10 | Out-File -Encoding utf8 $paths.LabelsJson
$labelPoliciesProjection | ConvertTo-Json -Depth 10 | Out-File -Encoding utf8 $paths.LabelPoliciesJson
$autoLabelPoliciesProjection | ConvertTo-Json -Depth 10 | Out-File -Encoding utf8 "$reportBase.autoLabelPolicies.json"
$autoLabelRulesProjection    | ConvertTo-Json -Depth 10 | Out-File -Encoding utf8 "$reportBase.autoLabelRules.json"
$dlpPoliciesProjection   | ConvertTo-Json -Depth 10 | Out-File -Encoding utf8 $paths.DlpPoliciesJson
$dlpRulesProjection      | ConvertTo-Json -Depth 10 | Out-File -Encoding utf8 $paths.DlpRulesJson
$meta                    | ConvertTo-Json -Depth 6  | Out-File -Encoding utf8 $paths.MetaJson

# CSV
$labelsProjection        | Export-Csv -Path $paths.LabelsCsv -NoTypeInformation -Encoding UTF8
$labelPoliciesProjection | Export-Csv -Path $paths.LabelPoliciesCsv -NoTypeInformation -Encoding UTF8
$autoLabelPoliciesProjection | Export-Csv -Path "$reportBase.autoLabelPolicies.csv" -NoTypeInformation -Encoding UTF8
$autoLabelRulesProjection    | Export-Csv -Path "$reportBase.autoLabelRules.csv" -NoTypeInformation -Encoding UTF8
$dlpPoliciesProjection   | Export-Csv -Path $paths.DlpPoliciesCsv -NoTypeInformation -Encoding UTF8
$dlpRulesProjection      | Export-Csv -Path $paths.DlpRulesCsv -NoTypeInformation -Encoding UTF8

Write-Info "Data export complete."

# ---------------------------
# Markdown Report
# ---------------------------
if ($CreateMarkdown) {
  Write-Info "Generating Markdown report: $($paths.MarkdownDoc)"
  try {
    $md = @()
    $md += "# Microsoft Purview Information Protection Configuration"
    $md += ""
    $md += "**Tenant:** $TenantDisplayName  "
    $md += "**Generated:** $((Get-Date).ToString('f'))  "
    $md += "**Report Version:** 1.0"
    $md += ""
    $md += "---"
    $md += ""
    $md += "## Summary"
    $md += ""
    $md += "- **Total Labels:** $($labelsProjection.Count)"
    $md += "- **Labels with Encryption:** $(($labelsProjection | Where-Object { $_.EncryptionEnabled }).Count)"
    $md += "- **Parent Labels:** $(($labelsProjection | Where-Object { -not $_.ParentLabelId }).Count)"
    $md += "- **Sublabels:** $(($labelsProjection | Where-Object { $_.ParentLabelId }).Count)"
    $md += "- **Label Policies:** $($labelPoliciesProjection.Count)"
    $md += "- **Auto-Labeling Policies:** $($autoLabelPoliciesProjection.Count)"
    $md += "- **DLP Policies:** $($dlpPoliciesProjection.Count)"
    $md += "- **DLP Rules:** $($dlpRulesProjection.Count)"
    $md += ""
    $md += "---"
    $md += ""
    $md += "## Sensitivity Labels"
    $md += ""
    $md += "📘 [Learn about sensitivity labels](https://learn.microsoft.com/microsoft-365/compliance/sensitivity-labels)"
    $md += ""
    
    # Group labels by parent/child
    $parentLabels = $labelsProjection | Where-Object { -not $_.ParentLabelId } | Sort-Object Priority
    
    foreach ($label in $parentLabels) {
      $md += "### $($label.Name)"
      $md += ""
      $md += "- **GUID:** ``$($label.Guid)``"
      $md += "- **Status:** $(if ($label.Enabled) { '✅ Enabled' } else { '❌ Disabled' })"
      $md += "- **Priority:** $($label.Priority)"
      if ($label.Tooltip) { $md += "- **User Description:** $($label.Tooltip)" }
      if ($label.Description) { $md += "- **Admin Notes:** $($label.Description)" }
      $md += ""
      
      if ($label.EncryptionEnabled) {
        $md += "#### 🔒 Protection"
        $md += ""
        $md += "- **Encryption:** ✅ Enabled"
        $md += "- **Type:** $($label.EncryptionSummary)"
        $md += ""
        $md += "📘 [Learn about encryption settings](https://learn.microsoft.com/microsoft-365/compliance/encryption-sensitivity-labels)"
        $md += ""
      }
      
      $md += "**Portal:** [Configure this label]($($label.PortalUrl))"
      $md += ""
      
      # Add sublabels
      $sublabels = $labelsProjection | Where-Object { $_.ParentLabelId -eq $label.Guid } | Sort-Object Priority
      if ($sublabels.Count -gt 0) {
        $md += "#### Sublabels"
        $md += ""
        foreach ($sub in $sublabels) {
          $md += "- **$($sub.Name)**"
          $md += "  - GUID: ``$($sub.Guid)``"
          $md += "  - Status: $(if ($sub.Enabled) { '✅' } else { '❌' })"
          if ($sub.EncryptionEnabled) { $md += "  - Protection: 🔒 Encrypted" }
          $md += "  - [Portal]($($sub.PortalUrl))"
        }
        $md += ""
      }
      
      $md += "---"
      $md += ""
    }
    
    $md += "## Label Policies"
    $md += ""
    $md += "📘 [Learn about label policies](https://learn.microsoft.com/microsoft-365/compliance/sensitivity-labels#what-label-policies-can-do)"
    $md += ""
    
    foreach ($policy in ($labelPoliciesProjection | Sort-Object Priority)) {
      $md += "### $($policy.Name)"
      $md += ""
      $md += "- **GUID:** ``$($policy.Guid)``"
      $md += "- **Status:** $(if ($policy.Enabled) { '✅ Enabled' } else { '❌ Disabled' })"
      $md += "- **Priority:** $($policy.Priority)"
      $md += "- **Mode:** $($policy.Mode)"
      
      $labelCount = 0
      try {
        $publishedLabels = $policy.LabelsPublished | ConvertFrom-Json
        $labelCount = $publishedLabels.Count
      } catch { }
      $md += "- **Published Labels:** $labelCount"
      
      $md += ""
      $md += "---"
      $md += ""
    }
    
    if ($autoLabelPoliciesProjection.Count -gt 0) {
      $md += "## Auto-Labeling Policies (Service-Side)"
      $md += ""
      $md += "📘 [Learn about auto-labeling](https://learn.microsoft.com/microsoft-365/compliance/apply-sensitivity-label-automatically)"
      $md += ""
      
      foreach ($policy in ($autoLabelPoliciesProjection | Sort-Object Priority)) {
        $md += "### $($policy.Name)"
        $md += ""
        $md += "- **Mode:** $($policy.Mode)"
        $md += "- **Applies Label:** $($policy.ApplySensitivityLabel)"
        $md += "- **Priority:** $($policy.Priority)"
        if ($policy.Comment) { $md += "- **Description:** $($policy.Comment)" }
        $md += ""
        $md += "---"
        $md += ""
      }
    }
    
    $md += "## Data Loss Prevention (DLP)"
    $md += ""
    $md += "📘 [Learn about DLP](https://learn.microsoft.com/microsoft-365/compliance/dlp-learn-about-dlp)"
    $md += ""
    
    foreach ($policy in ($dlpPoliciesProjection | Sort-Object Priority)) {
      $md += "### $($policy.Name)"
      $md += ""
      $md += "- **GUID:** ``$($policy.Guid)``"
      $md += "- **Status:** $(if ($policy.Enabled) { '✅ Enabled' } else { '❌ Disabled' })"
      $md += "- **Mode:** $($policy.Mode)"
      if ($policy.Comment) { $md += "- **Description:** $($policy.Comment)" }
      $md += ""
      
      # Find associated rules
      $policyRules = $dlpRulesProjection | Where-Object { $_.Policy -eq $policy.Name }
      if ($policyRules.Count -gt 0) {
        $md += "#### Rules ($($policyRules.Count))"
        $md += ""
        foreach ($rule in $policyRules) {
          $md += "- **$($rule.Name)** - Priority: $($rule.Priority), Status: $(if ($rule.Enabled) { '✅' } else { '❌' })"
        }
        $md += ""
      }
      
      $md += "**Portal:** [View policy](https://compliance.microsoft.com/datalossprevention/policies)"
      $md += ""
      $md += "---"
      $md += ""
    }
    
    $md += "## Additional Resources"
    $md += ""
    $md += "### Microsoft Purview Portal"
    $md += ""
    $md += "- [Information Protection Home](https://compliance.microsoft.com/informationprotection)"
    $md += "- [Sensitivity Labels](https://compliance.microsoft.com/informationprotection?viewid=sensitivitylabels)"
    $md += "- [Label Policies](https://compliance.microsoft.com/informationprotection?viewid=sensitivitylabelpolicies)"
    $md += "- [Auto-Labeling](https://compliance.microsoft.com/informationprotection?viewid=autolabeling)"
    $md += "- [DLP Policies](https://compliance.microsoft.com/datalossprevention?viewid=policies)"
    $md += ""
    $md += "### Microsoft Learn Documentation"
    $md += ""
    $md += "- [Get started with sensitivity labels](https://learn.microsoft.com/microsoft-365/compliance/get-started-with-sensitivity-labels)"
    $md += "- [Create and configure sensitivity labels](https://learn.microsoft.com/microsoft-365/compliance/create-sensitivity-labels)"
    $md += "- [DLP policy reference](https://learn.microsoft.com/microsoft-365/compliance/dlp-policy-reference)"
    $md += ""
    $md += "---"
    $md += ""
    $md += "*Report generated by Purview Configuration Documentation Tool*"
    
    $md | Out-File -FilePath $paths.MarkdownDoc -Encoding utf8
    Write-Info "Markdown report created."
  }
  catch {
    Write-Err "Failed to create Markdown report. $_"
  }
}

# ---------------------------
# Word Report (COM automation)
# ---------------------------
if ($CreateWord) {
  Write-Info "Generating Word report: $($paths.WordDoc)"
  try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $doc = $word.Documents.Add()
    $selection = $word.Selection

    function Add-Heading($text, [int]$level = 1) {
      $selection.Style = "Heading $level"
      $selection.TypeText($text)
      $selection.TypeParagraph()
    }
    function Add-Body($text) {
      $selection.Style = "Normal"
      $selection.TypeText($text)
      $selection.TypeParagraph()
    }
    function Add-TableFromObjects($objects, $columns) {
      if (-not $objects -or $objects.Count -eq 0) {
        Add-Body "No data."
        return
      }
      # Insert a table with header row
      $range = $selection.Range
      $table = $doc.Tables.Add($range, ($objects.Count + 1), $columns.Count)
      $table.Style = "Table Grid"

      # Header
      for ($c=0; $c -lt $columns.Count; $c++) {
        $table.Cell(1, $c+1).Range.Text = $columns[$c]
      }

      # Rows
      for ($r=0; $r -lt $objects.Count; $r++) {
        $rowObj = $objects[$r]
        for ($c=0; $c -lt $columns.Count; $c++) {
          $col = $columns[$c]
          $val = $rowObj.$col
          if ($val -is [string]) { $text = $val }
          else { $text = ($val | Out-String).Trim() }
          $table.Cell($r+2, $c+1).Range.Text = $text
        }
      }

      $selection.MoveDown()
      $selection.TypeParagraph()
    }

    # Title
    Add-Heading "$TenantDisplayName – Purview Information Protection & DLP Report" 1
    Add-Body "Generated: $((Get-Date).ToString('f')) (local)."

    # Labels
    Add-Heading "Sensitivity Labels" 2
    Add-Body "Total Labels: $($labelsProjection.Count) | With Encryption: $(($labelsProjection | Where-Object { $_.EncryptionEnabled }).Count)"
    Add-Body "Portal: https://compliance.microsoft.com/informationprotection?viewid=sensitivitylabels"
    Add-Body "Learn More: https://learn.microsoft.com/microsoft-365/compliance/sensitivity-labels"
    Add-TableFromObjects $labelsProjection @("Name","Guid","Enabled","Priority","EncryptionEnabled","EncryptionSummary","ParentLabelName","Tooltip","ModifiedTime")

    # Label Policies
    Add-Heading "Label Policies (Publishing)" 2
    Add-Body "Portal: https://compliance.microsoft.com/informationprotection?viewid=sensitivitylabelpolicies"
    Add-Body "Learn More: https://learn.microsoft.com/microsoft-365/compliance/sensitivity-labels#what-label-policies-can-do"
    Add-TableFromObjects $labelPoliciesProjection @("Name","Guid","Enabled","Priority","PublisherType","ApplyTo","LabelsPublished","AdvancedSettings","ModifiedTime")

    # Auto-Labeling Policies (Service-Side)
    if ($autoLabelPoliciesProjection.Count -gt 0) {
      Add-Heading "Auto-Labeling Policies (Service-Side)" 2
      Add-Body "Portal: https://compliance.microsoft.com/informationprotection?viewid=autolabeling"
      Add-Body "Learn More: https://learn.microsoft.com/microsoft-365/compliance/apply-sensitivity-label-automatically"
      Add-TableFromObjects $autoLabelPoliciesProjection @("Name","Mode","ApplySensitivityLabel","Priority","Comment","Locations","ModifiedTime")
    }

    # DLP Policies
    Add-Heading "DLP Policies" 2
    Add-Body "Portal: https://compliance.microsoft.com/datalossprevention?viewid=policies"
    Add-Body "Learn More: https://learn.microsoft.com/microsoft-365/compliance/dlp-learn-about-dlp"
    Add-TableFromObjects $dlpPoliciesProjection @("Name","Guid","Enabled","Mode","Comment","Workload","ExchangeLocation","SharePointLocation","OneDriveLocation","TeamsLocation","ModifiedTime")

    # DLP Rules
    Add-Heading "DLP Rules" 2
    Add-Body "Learn More: https://learn.microsoft.com/microsoft-365/compliance/dlp-policy-reference"
    Add-TableFromObjects $dlpRulesProjection @("Name","Guid","Policy","Enabled","Priority","Mode","Conditions","Exceptions","Actions","ModifiedTime")

    # Save & close
    $doc.SaveAs([ref]$paths.WordDoc)
    $doc.Close()
    $word.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
    Write-Info "Word report created."
  }
  catch {
    Write-Err "Failed to create Word report. $_"
  }
}

# ---------------------------
# PowerPoint Report (COM automation)
# ---------------------------
if ($CreatePowerPoint) {
  Write-Info "Generating PowerPoint: $($paths.PowerPoint)"
  try {
    $ppt = New-Object -ComObject PowerPoint.Application
    $ppt.Visible = [Microsoft.Office.Core.MsoTriState]::msoTrue
    $pres = $ppt.Presentations.Add()

    function AddTitleAndContentSlide($title, $bullets) {
      $layout = [Microsoft.Office.Interop.PowerPoint.PpSlideLayout]::ppLayoutText
      $slide = $pres.Slides.Add($pres.Slides.Count + 1, $layout)
      $slide.Shapes[1].TextFrame.TextRange.Text = $title
      $content = $slide.Shapes[2].TextFrame.TextRange
      $content.Text = ""
      foreach ($b in $bullets) {
        $content.Paragraphs($content.Paragraphs.Count + 1).Text = $b
      }
    }

    # Title slide
    AddTitleAndContentSlide "$TenantDisplayName – Purview IP & DLP Overview", @("Generated: $((Get-Date).ToString('f'))")

    # Labels slide
    $encryptedCount = ($labelsProjection | Where-Object { $_.EncryptionEnabled }).Count
    $parentCount = ($labelsProjection | Where-Object { -not $_.ParentLabelId }).Count
    $sublabelCount = ($labelsProjection | Where-Object { $_.ParentLabelId }).Count
    $labelSummary = @(
      "Total Labels: $($labelsProjection.Count)"
      "Parent Labels: $parentCount | Sublabels: $sublabelCount"
      "Labels with Encryption: $encryptedCount"
      ""
      "Portal: https://compliance.microsoft.com/informationprotection"
    )
    AddTitleAndContentSlide "Sensitivity Labels Overview", $labelSummary

    # Encryption Protection slide
    $protectionSummary = @(
      "Encryption (Rights Management):"
      "- Labels with encryption: $encryptedCount of $($labelsProjection.Count)"
      "- Protection includes access controls and usage rights"
      ""
      "Learn: https://learn.microsoft.com/microsoft-365/compliance/encryption-sensitivity-labels"
    )
    AddTitleAndContentSlide "Protection & Encryption", $protectionSummary

    # Label Policies slide
    $lpSummary = @(
      "Label Policies: $($labelPoliciesProjection.Count)"
      "Control who sees which labels"
      "User/group assignments captured"
      ""
      "Portal: https://compliance.microsoft.com/informationprotection?viewid=sensitivitylabelpolicies"
    )
    AddTitleAndContentSlide "Label Policies", $lpSummary

    # Auto-Labeling slide
    if ($autoLabelPoliciesProjection.Count -gt 0) {
      $alSummary = @(
        "Auto-Labeling Policies: $($autoLabelPoliciesProjection.Count)"
        "Service-side automatic classification"
        "Scans SharePoint, OneDrive, Exchange"
        ""
        "Portal: https://compliance.microsoft.com/informationprotection?viewid=autolabeling"
      )
      AddTitleAndContentSlide "Auto-Labeling (Service-Side)", $alSummary
    }

    # DLP Policies slide
    $dpSummary = @(
      "DLP Policies: $($dlpPoliciesProjection.Count)"
      "Locations: Exchange, SharePoint, OneDrive, Teams."
    )
    AddTitleAndContentSlide "DLP Policies", $dpSummary

    # DLP Rules slide
    $drSummary = @(
      "DLP Rules: $($dlpRulesProjection.Count)"
      "Conditions/Actions summarized."
    )
    AddTitleAndContentSlide "DLP Rules", $drSummary

    $pres.SaveAs($paths.PowerPoint)
    $pres.Close()
    $ppt.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($ppt) | Out-Null
    Write-Info "PowerPoint created."
  }
  catch {
    Write-Err "Failed to create PowerPoint. $_"
  }
}

Write-Info "All done. Files written:"
$paths.GetEnumerator() | Sort-Object Name | ForEach-Object { Write-Host " - $($_.Value)" }

# ---------------------------
# Disconnect IPPSSession
# ---------------------------
try { Disconnect-ExchangeOnline -Confirm:$false | Out-Null } catch { }
