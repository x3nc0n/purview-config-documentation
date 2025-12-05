# Microsoft Purview Configuration Documentation - Technical Guide

**Document Version:** 1.0  
**Last Updated:** December 5, 2025  
**Purpose:** Detailed technical documentation for extracting, formatting, and documenting Microsoft Purview Information Protection configurations

---

## Table of Contents

1. [Overview](#overview)
2. [Information Protection Architecture](#information-protection-architecture)
3. [Script Workflow Details](#script-workflow-details)
4. [Data Collection Methods](#data-collection-methods)
5. [Output Formats and Usage](#output-formats-and-usage)
6. [Configuration Elements Reference](#configuration-elements-reference)
7. [Portal and Documentation Links](#portal-and-documentation-links)
8. [Manual Configuration Guide](#manual-configuration-guide)
9. [Troubleshooting](#troubleshooting)

---

## Overview

This PowerShell automation tool extracts Microsoft Purview Information Protection and Data Loss Prevention configurations from a Microsoft 365 tenant and generates comprehensive documentation in multiple formats:

- **JSON** - Machine-readable configuration backup
- **CSV** - Spreadsheet-compatible tabular data
- **Markdown** - Human-readable documentation suitable for HTML conversion
- **Word (.docx)** - Detailed configuration report with hyperlinks
- **PowerPoint (.pptx)** - Executive summary presentation

### Key Features

- ✅ Extracts sensitivity labels with full protection settings
- ✅ Documents label policies and user/group assignments
- ✅ Captures DLP policies and rules with conditions/actions
- ✅ Preserves encryption and access control configurations
- ✅ Generates documentation with portal and Learn links
- ✅ Suitable for configuration drift detection and tenant migration

---

## Information Protection Architecture

### Components Overview

Microsoft Purview Information Protection consists of several interconnected components:

```
┌─────────────────────────────────────────────────────────────┐
│                    INFORMATION PROTECTION                    │
├─────────────────────────────────────────────────────────────┤
│                                                               │
│  ┌────────────────┐     ┌──────────────────┐               │
│  │ Sensitivity    │────▶│ Label Policies   │               │
│  │ Labels         │     │ (Publishing)     │               │
│  │                │     │                  │               │
│  │ - Protection   │     │ - User Scopes    │               │
│  │ - Marking      │     │ - Label Lists    │               │
│  │ - Conditions   │     │ - Settings       │               │
│  └────────────────┘     └──────────────────┘               │
│         │                                                    │
│         │                                                    │
│         ▼                                                    │
│  ┌────────────────┐     ┌──────────────────┐               │
│  │ Auto-Labeling  │     │ Label Analytics  │               │
│  │ Policies       │     │ (Activity Data)  │               │
│  └────────────────┘     └──────────────────┘               │
│                                                               │
└─────────────────────────────────────────────────────────────┘
```

### Label Hierarchy

Sensitivity labels support parent-child relationships (sublabels):

```
Public (Parent)
Confidential (Parent)
├── Confidential \ Finance (Sublabel)
├── Confidential \ HR (Sublabel)
└── Confidential \ Legal (Sublabel)
Highly Confidential (Parent)
├── Highly Confidential \ All Employees (Sublabel)
└── Highly Confidential \ Executive Only (Sublabel)
```

**Key Concepts:**
- **Parent labels** can have protection settings that apply to all sublabels
- **Sublabels** inherit parent's visual markings but can have unique encryption
- **Priority** determines which label applies when conflicts occur (lower number = higher priority)

---

## Script Workflow Details

### 1. Connection and Authentication

The script connects to the Security & Compliance Center PowerShell (IPPSSession):

```powershell
Connect-IPPSSession
```

**What happens:**
- Opens modern authentication browser window
- Requires account with Compliance Administrator role or higher
- Establishes session to protection.office.com endpoint
- Validates admin consent for required permissions

**Required Permissions:**
- Information Protection Administrator (Entra ID role)
- Or Compliance Administrator
- Or Global Administrator

**Microsoft Learn Reference:**  
📘 [Connect to Security & Compliance PowerShell](https://learn.microsoft.com/powershell/exchange/connect-to-scc-powershell)

---

### 2. Sensitivity Label Collection

#### 2.1 Get-Label Cmdlet

Retrieves all sensitivity labels configured in the tenant:

```powershell
$labels = Get-Label -ResultSize Unlimited
```

**Data Retrieved:**
- `Name` - Display name of the label
- `ImmutableId` (GUID) - Unique identifier that persists across renames
- `Priority` - Numeric priority (0 = highest)
- `Enabled` - Whether label is active
- `Tooltip` - User-facing description
- `Comment` - Admin notes
- `ParentId` - GUID of parent label (for sublabels)
- `ContentMarking` - Visual markings configuration
- `Encryption` - Protection and access controls
- `SiteAndGroupProtectionSettings` - Teams/SharePoint/M365 Groups settings
- `AutoLabeling` - Client-side auto-classification rules
- `LocaleSettings` - Multi-language display names

**Portal Location:**  
🌐 [Microsoft Purview portal > Information protection > Labels](https://compliance.microsoft.com/informationprotection?viewid=sensitivitylabels)

**Microsoft Learn Reference:**  
📘 [Learn about sensitivity labels](https://learn.microsoft.com/microsoft-365/compliance/sensitivity-labels)

---

#### 2.2 Understanding Label Protection Settings

##### **Encryption (Rights Management)**

When a label has encryption enabled, it provides:

**Access Controls:**
- **Users and Groups** - Specific people who can access content
- **Rights/Permissions** - What users can do (View, Edit, Copy, Print, Forward, etc.)
- **Expiration** - When access expires
- **Offline Access** - Duration allowed without connectivity
- **Double Key Encryption (DKE)** - Customer-controlled encryption key

**Example Encryption Object:**
```json
{
  "EncryptionEnabled": true,
  "EncryptionProtectionType": "Template",
  "EncryptionRightsDefinitions": [
    {
      "Identity": "AllStaff@contoso.com",
      "Rights": ["VIEW", "EDIT", "DOCEDIT", "PRINT"]
    },
    {
      "Identity": "Executives@contoso.com", 
      "Rights": ["OWNER"]
    }
  ],
  "EncryptionContentExpiredOnDateInDaysOrNever": "Never",
  "EncryptionOfflineAccessDays": 30
}
```

**Common Rights:**
- `VIEW` - Open and read
- `EDIT` / `DOCEDIT` - Modify content
- `PRINT` - Print document
- `COPY` - Copy content to clipboard
- `EXPORT` / `EXTRACT` - Save as different format
- `OWNER` - Full control including removing protection
- `VIEWRIGHTSDATA` - View document rights
- `REPLY` / `REPLYALL` / `FORWARD` - Email-specific rights

**Portal Configuration:**  
When creating/editing a label, encryption is configured under:
1. Navigate to label settings
2. **"Encryption"** section
3. Configure permissions assignments
4. Set access restrictions

**Microsoft Learn References:**  
📘 [Restrict access to content using encryption](https://learn.microsoft.com/microsoft-365/compliance/encryption-sensitivity-labels)  
📘 [Configure encryption settings](https://learn.microsoft.com/microsoft-365/compliance/encryption-sensitivity-labels#configure-encryption-settings)

---

##### **Content Marking**

Visual indicators applied to documents:

**Types:**
- **Header** - Text at top of every page
- **Footer** - Text at bottom of every page
- **Watermark** - Diagonal background text

**Example ContentMarking Object:**
```json
{
  "ContentMarkingEnabled": true,
  "Header": {
    "Enabled": true,
    "Text": "CONFIDENTIAL - Internal Use Only",
    "FontSize": 12,
    "FontColor": "#FF0000",
    "Alignment": "Center"
  },
  "Footer": {
    "Enabled": true,
    "Text": "Classification: Confidential | ${User.DisplayName} | ${Date}",
    "FontSize": 10,
    "FontColor": "#000000",
    "Alignment": "Left"
  },
  "Watermark": {
    "Enabled": true,
    "Text": "CONFIDENTIAL",
    "FontSize": 48,
    "FontColor": "#D3D3D3",
    "Layout": "Diagonal"
  }
}
```

**Dynamic Variables:**
- `${User.DisplayName}` - User who applied label
- `${User.PrincipalName}` - User's email
- `${Date}` - Current date
- `${Time}` - Current time
- `${Label.Name}` - Label name

**Microsoft Learn Reference:**  
📘 [Add content marking to documents](https://learn.microsoft.com/microsoft-365/compliance/sensitivity-labels-office-apps#dynamic-markings-with-variables)

---

##### **Auto-Labeling (Client-Side)**

Conditions that trigger automatic label application:

**Sensitive Information Types (SITs):**
- Credit card numbers
- Social security numbers
- Bank account numbers
- Passport numbers
- Custom regex patterns

**Trainable Classifiers:**
- Resumes
- Source code
- Harassment
- Profanity
- Financial statements

**Example AutoLabeling Configuration:**
```json
{
  "AutoLabelingEnabled": true,
  "Conditions": [
    {
      "ConditionType": "SensitiveInformationType",
      "Value": "Credit Card Number",
      "MinCount": 1,
      "MaxCount": 10
    },
    {
      "ConditionType": "SensitiveInformationType", 
      "Value": "U.S. Social Security Number",
      "MinCount": 1
    }
  ],
  "Operator": "And",
  "AutoApplyType": "Recommend",
  "PolicyTip": "This document contains sensitive financial information. Apply Confidential label?"
}
```

**Auto-Apply Modes:**
- **Recommend** - Suggests label, user can dismiss
- **Automatic** - Applies label silently
- **AutoApply with User Override** - User can change

**Microsoft Learn Reference:**  
📘 [Apply a sensitivity label automatically](https://learn.microsoft.com/microsoft-365/compliance/apply-sensitivity-label-automatically)

---

##### **Endpoint Protection**

Microsoft Purview Data Loss Prevention integration for labeled files:

**Device Controls:**
- Restrict copying to removable media
- Prevent printing
- Restrict copy/paste to non-corporate apps
- Block upload to cloud services
- Prevent screen capture

**Example:**
```json
{
  "EndpointProtectionEnabled": true,
  "BlockPrint": true,
  "BlockCopyToClipboard": false,
  "BlockScreenshot": true,
  "BlockCloudUpload": true,
  "AllowedDomains": ["contoso.com", "sharepoint.com"]
}
```

**Microsoft Learn Reference:**  
📘 [Endpoint data loss prevention](https://learn.microsoft.com/microsoft-365/compliance/endpoint-dlp-learn-about)

---

### 3. Label Policy Collection

Label policies control **who sees which labels** and **policy settings**.

#### 3.1 Get-LabelPolicy Cmdlet

```powershell
$labelPolicies = Get-LabelPolicy -ResultSize Unlimited
```

**Data Retrieved:**
- `Name` - Policy name
- `Priority` - When multiple policies apply, higher priority wins
- `Labels` - List of label GUIDs published by this policy
- `ApplyTo` - Users and groups who receive this policy
- `Settings` - Policy-level configurations
- `AdvancedSettings` - Custom key-value configurations

**Key Policy Settings:**

##### **Mandatory Labeling**
```json
{
  "RequireDocumentLabel": true,
  "RequireEmailLabel": true
}
```
Forces users to apply a label before saving/sending.

##### **Default Label**
```json
{
  "DefaultLabelId": "8faca7b8-8d20-48a3-8ea2-0f96310a848e"
}
```
Auto-applies specified label to new documents.

##### **Label Justification**
```json
{
  "JustificationRequired": true
}
```
Requires users to explain why they're downgrading or removing a label.

##### **Custom Permissions**
```json
{
  "EnableCustomPermissions": true
}
```
Allows users to define custom protection settings.

##### **Outlook-Specific**
```json
{
  "OutlookDefaultLabel": "8faca7b8-8d20-48a3-8ea2-0f96310a848e",
  "DisableMandatoryInOutlook": false,
  "OutlookBlockTrustedDomains": ["external.com"],
  "OutlookBlockUntrustedCollaborationLabel": "ConfidentialGuid"
}
```

**Portal Location:**  
🌐 [Microsoft Purview portal > Information protection > Label policies](https://compliance.microsoft.com/informationprotection?viewid=sensitivitylabelpolicies)

**Microsoft Learn Reference:**  
📘 [Create and configure sensitivity label policies](https://learn.microsoft.com/microsoft-365/compliance/sensitivity-labels#what-label-policies-can-do)

---

#### 3.2 Label Policy Scopes

Policies can target specific user populations:

**Scope Types:**

1. **All Users in Organization**
   ```json
   { "ApplyTo": ["All"] }
   ```

2. **Specific Users**
   ```json
   { "ApplyTo": ["user1@contoso.com", "user2@contoso.com"] }
   ```

3. **Security Groups / M365 Groups**
   ```json
   { 
     "ApplyTo": ["Finance-Team@contoso.com", "Legal-Department@contoso.com"],
     "ExceptApplyTo": ["External-Contractors@contoso.com"]
   }
   ```

4. **Adaptive Scopes** (Recommended)
   Dynamic groups based on attributes:
   ```json
   {
     "ApplyToAdaptiveScopes": ["All-Finance-Users", "All-Executive-Users"]
   }
   ```

**Best Practice:**  
Use adaptive scopes for automatic user inclusion based on department, location, or other attributes.

**Microsoft Learn Reference:**  
📘 [Learn about adaptive scopes](https://learn.microsoft.com/microsoft-365/compliance/purview-adaptive-scopes)

---

### 4. Auto-Labeling Policies (Service-Side)

**Important:** These are different from client-side auto-labeling on labels!

#### 4.1 Get-AutoSensitivityLabelPolicy

```powershell
$autoLabelPolicies = Get-AutoSensitivityLabelPolicy
$autoLabelRules = Get-AutoSensitivityLabelRule
```

**Service-Side vs Client-Side:**

| Feature | Client-Side (Label) | Service-Side (Policy) |
|---------|-------------------|----------------------|
| **Applies To** | Office apps (Word, Excel, PowerPoint, Outlook) | SharePoint, OneDrive, Exchange |
| **Timing** | When user creates/edits document | Scans existing and new content |
| **Simulation Mode** | No | Yes - test before enforcing |
| **Exchange** | Outlook client only | Server-side email scanning |

**Example Auto-Label Policy:**
```json
{
  "Name": "Auto-label Credit Card Documents",
  "LabelId": "ConfidentialFinanceGuid",
  "Locations": [
    "SharePointOnline",
    "OneDriveForBusiness"
  ],
  "Conditions": [
    {
      "Type": "SensitiveInfoType",
      "Value": "Credit Card Number",
      "MinCount": 5
    }
  ],
  "Mode": "TestWithNotifications",
  "Comment": "Pilot: Auto-classify documents with 5+ credit cards"
}
```

**Policy Modes:**
1. **Simulation** - Reports what would be labeled (no changes)
2. **TestWithNotifications** - Labels content, sends notifications
3. **Enable** - Fully operational

**Microsoft Learn Reference:**  
📘 [Automatically apply a sensitivity label in Microsoft 365](https://learn.microsoft.com/microsoft-365/compliance/apply-sensitivity-label-automatically)

---

### 5. Data Loss Prevention (DLP) Collection

DLP prevents sensitive information from being shared inappropriately.

#### 5.1 Get-DlpCompliancePolicy

```powershell
$dlpPolicies = Get-DlpCompliancePolicy
```

**Data Retrieved:**
- `Name` - Policy name
- `Mode` - Enable, TestWithNotifications, TestWithoutNotifications, Disable
- `Workload` - Where policy applies (Exchange, SharePoint, OneDrive, Teams, Devices)
- `ExchangeLocation` - Mailboxes included (All or specific)
- `SharePointLocation` - Sites included
- `OneDriveLocation` - OneDrive accounts
- `TeamsLocation` - Teams chats and channels
- `EndpointDlpLocation` - Windows devices
- `Priority` - When multiple policies match, determines which applies

**Example DLP Policy:**
```json
{
  "Name": "Protect Credit Card Information",
  "Mode": "Enable",
  "Workload": ["Exchange", "SharePoint", "OneDrive", "Teams"],
  "ExchangeLocation": "All",
  "SharePointLocation": "All",
  "OneDriveLocation": "All",
  "TeamsLocation": "All",
  "Priority": 1,
  "Comment": "Prevents sharing of financial data"
}
```

**Portal Location:**  
🌐 [Microsoft Purview portal > Data loss prevention > Policies](https://compliance.microsoft.com/datalossprevention?viewid=policies)

**Microsoft Learn Reference:**  
📘 [Learn about data loss prevention](https://learn.microsoft.com/microsoft-365/compliance/dlp-learn-about-dlp)

---

#### 5.2 Get-DlpComplianceRule

Rules define the **conditions** and **actions** for DLP policies:

```powershell
$dlpRules = Get-DlpComplianceRule
```

**Rule Structure:**

##### **Conditions** (When to trigger)
```json
{
  "Conditions": {
    "ContentContainsSensitiveInformation": [
      {
        "Name": "Credit Card Number",
        "MinCount": 1,
        "MaxCount": 10,
        "MinConfidence": 85
      }
    ],
    "ContentIsSharedWith": "ExternalUsers",
    "DocumentIsUnsupportedFileType": false,
    "HasSensitivityLabel": {
      "Labels": ["ConfidentialGuid", "HighlyConfidentialGuid"]
    }
  }
}
```

**Common Conditions:**
- Sensitive information types detected
- Document has specific sensitivity label
- Shared with external users
- Document size exceeds threshold
- User is in specific group
- File extension matches pattern

##### **Actions** (What to do)
```json
{
  "Actions": [
    {
      "Type": "BlockAccess",
      "BlockAccessScope": "All"
    },
    {
      "Type": "NotifyUser",
      "NotifyUserType": "NotSet",
      "NotifyEmail": ["SecurityTeam@contoso.com"]
    },
    {
      "Type": "GenerateIncidentReport",
      "IncidentReportContent": ["DocumentName", "Sender", "MatchedRules"],
      "ReportTo": ["ComplianceTeam@contoso.com"]
    },
    {
      "Type": "Quarantine",
      "QuarantineTag": "AdminOnlyTag"
    }
  ]
}
```

**Available Actions:**
- **BlockAccess** - Prevent sharing/sending
- **BlockAccessWithOverride** - Block but allow business justification
- **NotifyUser** - Email policy tip
- **GenerateIncidentReport** - Alert administrators
- **GenerateAlert** - Send alert to Security & Compliance center
- **Quarantine** - Move email to quarantine
- **SetHeader** - Add X-header to email
- **RemoveHeader** - Remove email header
- **RedirectMessageTo** - Reroute email
- **DeleteMessage** - Remove email (audit logged)
- **Encrypt** - Apply O365 Message Encryption

##### **Exceptions** (When NOT to trigger)
```json
{
  "Exceptions": {
    "RecipientDomainIs": ["contoso.com"],
    "SenderIpRange": ["10.0.0.0/8"],
    "DocumentCreatedBy": ["AutomatedService@contoso.com"]
  }
}
```

**Microsoft Learn Reference:**  
📘 [Data loss prevention policy reference](https://learn.microsoft.com/microsoft-365/compliance/dlp-policy-reference)

---

### 6. Graph API Enrichment (Optional)

The script optionally uses Microsoft Graph beta endpoints for additional data:

```powershell
Connect-MgGraph -Scopes "SecurityEvents.Read.All", "Policy.Read.All"
```

**Graph Endpoints Used:**
- `/security/informationProtection/policyLabels` - Label definitions
- `/security/informationProtection/labelPolicies` - Policy configurations
- `/security/dataLossPreventionPolicies` - DLP configurations

**Benefits:**
- Cross-platform compatibility (not Windows-only)
- Additional metadata and usage analytics
- Integration with Azure AD groups

**Limitations:**
- Beta endpoints may change
- Requires additional admin consent
- Some settings only available via Compliance PowerShell

**Microsoft Learn Reference:**  
📘 [Microsoft Graph API for information protection](https://learn.microsoft.com/graph/api/resources/informationprotection)

---

## Output Formats and Usage

### 1. JSON Format

**Purpose:** Machine-readable backup, configuration as code, version control

**Files Generated:**
- `PurviewDoc_TIMESTAMP.labels.json` - All sensitivity labels
- `PurviewDoc_TIMESTAMP.labelPolicies.json` - All label policies
- `PurviewDoc_TIMESTAMP.dlpPolicies.json` - DLP policies
- `PurviewDoc_TIMESTAMP.dlpRules.json` - DLP rules
- `PurviewDoc_TIMESTAMP.meta.json` - Metadata about export

**Use Cases:**
- Configuration drift detection (compare JSON files)
- Automated compliance audits
- Infrastructure-as-code implementations
- API integration with SIEM/SOAR tools

**Example Label JSON:**
```json
{
  "Name": "Confidential - Finance",
  "Guid": "8faca7b8-8d20-48a3-8ea2-0f96310a848e",
  "Priority": 10,
  "Enabled": true,
  "Tooltip": "Financial data requiring encryption",
  "Description": "Used for sensitive financial documents",
  "ContentMarking": {
    "Header": {
      "Enabled": true,
      "Text": "CONFIDENTIAL - FINANCE DEPARTMENT"
    }
  },
  "Encryption": {
    "EncryptionEnabled": true,
    "Rights": ["Finance-Team@contoso.com|VIEW,EDIT", "Executives@contoso.com|OWNER"]
  },
  "ModifiedTime": "2025-11-15T14:30:00Z"
}
```

---

### 2. CSV Format

**Purpose:** Spreadsheet analysis, pivot tables, mail merge

**Files Generated:**
- `PurviewDoc_TIMESTAMP.labels.csv`
- `PurviewDoc_TIMESTAMP.labelPolicies.csv`
- `PurviewDoc_TIMESTAMP.dlpPolicies.csv`
- `PurviewDoc_TIMESTAMP.dlpRules.csv`

**Use Cases:**
- Quick Excel analysis
- Create label inventory reports
- Identify labels without encryption
- Find policies with no users assigned

**Note:** Complex nested objects (like Encryption settings) are JSON-stringified within CSV cells.

---

### 3. Markdown Format (.md)

**Purpose:** Human-readable documentation, static site generation, wikis

**Recommended Structure:**

````markdown
# Tenant: Contoso Corporation - Sensitivity Labels

**Generated:** 2025-12-05 10:30:00 UTC

## Summary

- **Total Labels:** 15
- **Labels with Encryption:** 8
- **Labels with Auto-Labeling:** 3
- **Active Label Policies:** 4

---

## Sensitivity Labels

### 1. Public

- **GUID:** `a1b2c3d4-e5f6-7890-abcd-ef1234567890`
- **Status:** ✅ Enabled
- **Priority:** 100
- **Protection:** None
- **Description:** Information approved for public release

**Portal:** [View in Purview](https://compliance.microsoft.com/informationprotection/labels/a1b2c3d4-e5f6-7890-abcd-ef1234567890)

---

### 2. Confidential - Finance

- **GUID:** `8faca7b8-8d20-48a3-8ea2-0f96310a848e`
- **Status:** ✅ Enabled
- **Priority:** 10
- **Protection:** 🔒 **Encryption Enabled**

#### Encryption Details

- **Type:** Template-based
- **Access Control:**
  - `Finance-Team@contoso.com` - View, Edit, Print
  - `Executives@contoso.com` - Owner (Full Control)
- **Expiration:** Never
- **Offline Access:** 30 days

#### Visual Markings

- **Header:** "CONFIDENTIAL - FINANCE DEPARTMENT" (Red, 12pt, Center)
- **Footer:** "Classification: Confidential | ${User.DisplayName} | ${Date}"
- **Watermark:** "CONFIDENTIAL" (Diagonal, Gray)

#### Auto-Labeling

- **Trigger:** Document contains ≥5 credit card numbers
- **Mode:** Recommend to user

**Portal:** [View in Purview](https://compliance.microsoft.com/informationprotection/labels/8faca7b8-8d20-48a3-8ea2-0f96310a848e)  
**Learn More:** [Restrict access using encryption](https://learn.microsoft.com/microsoft-365/compliance/encryption-sensitivity-labels)

---
````

---

### 4. Word Document (.docx)

**Purpose:** Comprehensive configuration report with navigation, suitable for audits

**Document Structure:**

1. **Cover Page**
   - Tenant name and logo
   - Report generation date
   - Document version
   - Prepared by information

2. **Table of Contents** (Auto-generated)

3. **Executive Summary**
   - Configuration overview
   - Key statistics
   - Recent changes

4. **Sensitivity Labels Section**
   - Each label on separate page
   - Full settings documentation
   - Hyperlinks to portal and Learn docs
   - Screenshots (if available)

5. **Label Policies Section**
   - Policy-by-policy breakdown
   - User/group assignments
   - Settings explained

6. **DLP Section**
   - Policies and rules
   - Condition/action explanations
   - Affected locations

7. **Appendix**
   - Cmdlets used
   - Export timestamp
   - Raw data tables

**Hyperlink Examples:**
- **Portal Links:** Each label/policy links to its configuration page
- **Learn Links:** Inline references to Microsoft Learn documentation
- **Cross-References:** Internal document navigation

**Microsoft Learn Reference for Manual Setup:**  
📘 [Sensitivity labels quick start](https://learn.microsoft.com/microsoft-365/compliance/get-started-with-sensitivity-labels)

---

### 5. PowerPoint Presentation (.pptx)

**Purpose:** Executive briefing, stakeholder communication, less technical detail

**Slide Structure:**

1. **Title Slide**
   - Tenant name
   - Report date
   - Purview Information Protection Overview

2. **Configuration Summary Slide**
   - Metrics cards (labels count, policies count, DLP policies)
   - Pie chart: Labels with/without encryption
   - Bar chart: Labels by priority

3. **Sensitivity Labels Overview**
   - Table: Top 10 labels by priority
   - Icons indicating encryption status
   - Usage frequency (if available)

4. **Label Policies Overview**
   - User coverage statistics
   - Policy scopes (All users vs. specific groups)
   - Settings highlights (mandatory labeling, default labels)

5. **Protection Highlights**
   - Encryption-enabled labels
   - Access control summary
   - Visual marking examples

6. **DLP Overview**
   - Policy count by workload
   - Top sensitive info types protected
   - Action summary (block, notify, alert)

7. **Next Steps / Recommendations**
   - Configuration gaps identified
   - Suggested improvements
   - Contact information

**Design Notes:**
- Use Microsoft Purview color scheme (purple/blue)
- Icons for encryption, marking, auto-labeling
- Charts for quantitative data
- Minimal text per slide (5-7 bullets max)

---

## Configuration Elements Reference

### Sensitivity Label Properties

| Property | Type | Description | Example |
|----------|------|-------------|---------|
| `Name` | String | Display name | "Confidential - Finance" |
| `ImmutableId` | GUID | Permanent identifier | "8faca7b8-..." |
| `ParentId` | GUID | Parent label (if sublabel) | "a1b2c3d4-..." |
| `Priority` | Integer | 0 = highest priority | 10 |
| `Enabled` | Boolean | Whether label is active | true |
| `Tooltip` | String | User-facing description | "For financial data" |
| `Comment` | String | Admin notes | "Created for Q4 2025" |
| `ContentMarking` | Object | Headers/footers/watermarks | {...} |
| `Encryption` | Object | Protection settings | {...} |
| `SiteAndGroupProtectionSettings` | Object | Teams/SharePoint privacy | {...} |
| `AutoLabelingSettings` | Object | Auto-apply conditions | {...} |
| `LocaleSettings` | Array | Multi-language names | [...] |
| `WhenCreated` | DateTime | Creation timestamp | "2025-01-15T..." |
| `WhenChanged` | DateTime | Last modification | "2025-11-20T..." |

---

### Label Policy Properties

| Property | Type | Description |
|----------|------|-------------|
| `Name` | String | Policy name |
| `ImmutableId` | GUID | Permanent identifier |
| `Priority` | Integer | Policy precedence |
| `Labels` | Array[GUID] | Published label IDs |
| `ApplyTo` | Array[String] | Users/groups |
| `ExceptApplyTo` | Array[String] | Excluded users |
| `Settings` | Object | Policy configurations |
| `AdvancedSettings` | Object | Custom key-value pairs |
| `Mode` | String | Enable/Test/Disable |

**Advanced Settings Examples:**
```json
{
  "outlookdefaultlabel": "8faca7b8-...",
  "enablecontainertelemetry": "true",
  "powerbimandatory": "true",
  "disablemandatorybeforemetadata": "true"
}
```

---

### DLP Policy Properties

| Property | Type | Description |
|----------|------|-------------|
| `Name` | String | Policy name |
| `Identity` | String | Policy ID |
| `Mode` | String | Enable/Test/Disable |
| `Workload` | Array | Exchange/SharePoint/OneDrive/Teams/Devices |
| `ExchangeLocation` | Array | Mailboxes (All or specific) |
| `SharePointLocation` | Array | Sites |
| `OneDriveLocation` | Array | OneDrive accounts |
| `TeamsLocation` | Array | Teams chats/channels |
| `EndpointDlpLocation` | Array | Windows devices |
| `Priority` | Integer | Rule precedence |

---

### DLP Rule Properties

| Property | Type | Description |
|----------|------|-------------|
| `Name` | String | Rule name |
| `Policy` | String | Parent policy name |
| `Priority` | Integer | Rule order within policy |
| `Conditions` | Object | When to trigger |
| `Exceptions` | Object | When NOT to trigger |
| `Actions` | Array | What to do |

**Condition Types:**
- `ContentContainsSensitiveInformation`
- `ContentIsSharedWith`
- `HasSensitivityLabel`
- `DocumentSizeOver`
- `SenderDomainIs`
- `RecipientDomainIs`

**Action Types:**
- `BlockAccess`
- `NotifyUser`
- `GenerateIncidentReport`
- `Quarantine`
- `Encrypt`

---

## Portal and Documentation Links

### Microsoft Purview Portal

| Component | Portal URL |
|-----------|-----------|
| **Information Protection Home** | https://compliance.microsoft.com/informationprotection |
| **Sensitivity Labels** | https://compliance.microsoft.com/informationprotection?viewid=sensitivitylabels |
| **Label Policies** | https://compliance.microsoft.com/informationprotection?viewid=sensitivitylabelpolicies |
| **Auto-Labeling Policies** | https://compliance.microsoft.com/informationprotection?viewid=autolabeling |
| **DLP Policies** | https://compliance.microsoft.com/datalossprevention?viewid=policies |
| **DLP Alerts** | https://compliance.microsoft.com/datalossprevention?viewid=dlpalerts |
| **Sensitive Info Types** | https://compliance.microsoft.com/classificationdatatypes |
| **Trainable Classifiers** | https://compliance.microsoft.com/classificationtrainableclassifiers |

### Microsoft Learn Documentation

| Topic | Learn URL |
|-------|-----------|
| **Sensitivity Labels Overview** | https://learn.microsoft.com/microsoft-365/compliance/sensitivity-labels |
| **Get Started Guide** | https://learn.microsoft.com/microsoft-365/compliance/get-started-with-sensitivity-labels |
| **Encryption Configuration** | https://learn.microsoft.com/microsoft-365/compliance/encryption-sensitivity-labels |
| **Auto-Labeling** | https://learn.microsoft.com/microsoft-365/compliance/apply-sensitivity-label-automatically |
| **DLP Overview** | https://learn.microsoft.com/microsoft-365/compliance/dlp-learn-about-dlp |
| **DLP Policy Reference** | https://learn.microsoft.com/microsoft-365/compliance/dlp-policy-reference |
| **Compliance PowerShell** | https://learn.microsoft.com/powershell/exchange/connect-to-scc-powershell |
| **Label Policy Settings** | https://learn.microsoft.com/microsoft-365/compliance/sensitivity-labels#what-label-policies-can-do |

### Direct Label Configuration URL Pattern

To link directly to a specific label configuration page:

```
https://compliance.microsoft.com/informationprotection/labels/{ImmutableId}
```

Example:
```
https://compliance.microsoft.com/informationprotection/labels/8faca7b8-8d20-48a3-8ea2-0f96310a848e
```

---

## Manual Configuration Guide

### Scenario: Tenant Migration or Configuration Drift Recovery

When you need to manually recreate labels based on this documentation:

#### Step 1: Create Sensitivity Labels

1. **Navigate to Portal:**  
   🌐 https://compliance.microsoft.com/informationprotection?viewid=sensitivitylabels

2. **Click "Create a label"**

3. **Configure Basic Settings:**
   - **Name:** Use exact name from documentation
   - **Display name:** Same as Name
   - **Description for users:** Copy from `Tooltip` field
   - **Description for admins:** Copy from `Comment` field

4. **Configure Scope:**
   - ☑ Files & emails
   - ☑ Groups & sites (if `SiteAndGroupProtectionSettings` present)
   - ☑ Schematized data assets (if applicable)

5. **Configure Encryption** (if present in documentation):
   - Select "Configure encryption settings"
   - Choose "Assign permissions now"
   - Click "Assign permissions"
   - Add users/groups with appropriate rights:
     - Extract from `Encryption.EncryptionRightsDefinitions`
     - Match rights to permission levels (Co-Author, Reviewer, Viewer, etc.)
   - Set offline access duration from `EncryptionOfflineAccessDays`
   - Configure expiration if not "Never"

6. **Configure Content Marking** (if present):
   - **Header:**
     - Enable if `ContentMarking.Header.Enabled = true`
     - Text: Copy from `Header.Text`
     - Font size: From `Header.FontSize`
     - Color: From `Header.FontColor`
     - Alignment: From `Header.Alignment`
   - **Footer:** Same process
   - **Watermark:** Same process

7. **Configure Auto-Labeling** (if present):
   - Choose "Define auto-labeling conditions"
   - Add sensitive info types from `AutoLabelingSettings.Conditions`
   - Set instance counts (MinCount, MaxCount)
   - Choose recommendation vs. automatic application

8. **Configure Teams & Groups Settings** (if applicable):
   - Privacy: Public vs. Private
   - External user access
   - External sharing

9. **Review and Submit**

#### Step 2: Create Label Policies

1. **Navigate to Portal:**  
   🌐 https://compliance.microsoft.com/informationprotection?viewid=sensitivitylabelpolicies

2. **Click "Publish labels"**

3. **Select Labels:**
   - Choose labels from `LabelsPublished` array (by name or GUID)

4. **Assign Users/Groups:**
   - Add users/groups from `ApplyTo` array
   - Exclude users from `ExceptApplyTo` if present

5. **Configure Policy Settings:**
   - **Require users to apply label:** From `Settings.RequireDocumentLabel`
   - **Provide default label:** From `Settings.DefaultLabelId`
   - **Require justification:** From `Settings.JustificationRequired`
   - **Provide custom help page URL:** From `AdvancedSettings.customurl`
   - **Outlook default label:** From `AdvancedSettings.outlookdefaultlabel`

6. **Name the Policy:**
   - Use exact name from documentation

7. **Submit**

#### Step 3: Create Auto-Labeling Policies (Service-Side)

1. **Navigate to Portal:**  
   🌐 https://compliance.microsoft.com/informationprotection?viewid=autolabeling

2. **Click "Create auto-labeling policy"**

3. **Choose Information Type:**
   - Select templates or custom conditions
   - Add sensitive info types with instance counts

4. **Choose Locations:**
   - Select from: SharePoint, OneDrive, Exchange
   - Configure "All" or specific sites/mailboxes

5. **Set Label:**
   - Choose label to auto-apply

6. **Configure Mode:**
   - **Simulation:** Test first (recommended)
   - **Enable:** Apply automatically

7. **Name and Submit**

#### Step 4: Create DLP Policies

1. **Navigate to Portal:**  
   🌐 https://compliance.microsoft.com/datalossprevention?viewid=policies

2. **Click "Create policy"**

3. **Choose Template or Custom:**
   - Templates: Financial, Medical, Privacy
   - Custom: Build from scratch

4. **Configure Locations:**
   - From `Workload` and location arrays in documentation

5. **Define Policy Rules:**
   - **Conditions:** From `DlpRule.Conditions`
     - Sensitive info types
     - Label requirements
     - Sharing scope
   - **Actions:** From `DlpRule.Actions`
     - Block access
     - Notify users
     - Generate incident report
     - Quarantine

6. **Configure Notifications:**
   - Policy tips text
   - Email notifications
   - Incident reports

7. **Test Mode:**
   - Start in "Test it out first" mode
   - Review DLP alerts before enforcing

8. **Name and Submit**

---

## Troubleshooting

### Common Issues

#### 1. Connect-IPPSSession Fails

**Symptoms:**
```
Connect-IPPSSession : Access Denied. You do not have permission to connect to this endpoint.
```

**Solutions:**
- Verify account has Compliance Administrator or higher
- Check MFA is enabled and can authenticate
- Ensure no Conditional Access policies blocking connection
- Try: `Connect-IPPSSession -UserPrincipalName admin@contoso.com`

**Reference:**  
📘 [Troubleshoot connection issues](https://learn.microsoft.com/powershell/exchange/connect-to-scc-powershell#troubleshoot-connection-issues)

---

#### 2. Get-Label Returns Empty

**Symptoms:**
```
Found 0 labels.
```

**Solutions:**
- Labels may not be created yet
- User may lack permission to view labels
- Wait 24 hours after creating first label (propagation delay)
- Verify license: E3/E5 or standalone Information Protection license required

---

#### 3. Word/PowerPoint Generation Fails

**Symptoms:**
```
[ERROR] Failed to create Word report. Cannot create object: "Word.Application"
```

**Solutions:**
- Requires Windows OS with Office installed
- Office must be properly activated
- Run PowerShell as Administrator
- Check COM security settings
- Alternative: Use only JSON/CSV/MD outputs on Mac/Linux

---

#### 4. Encryption Settings Unreadable in CSV

**Symptoms:**
CSV shows: `{"EncryptionEnabled":true,"EncryptionRightsDefinitions":[...]}`

**Solutions:**
- Complex objects are JSON-stringified in CSV
- Use JSON output for full fidelity
- In Excel: Use "Text to Columns" or JSON parsing functions
- Recommended: Import JSON into PowerShell and re-export specific fields

---

#### 5. Graph Enrichment Fails

**Symptoms:**
```
[WARN] Graph labels endpoint failed/enforced.
```

**Solutions:**
- Graph beta endpoints subject to change
- Not all tenants have features enabled
- Admin consent may be required for additional scopes
- This is optional - core functionality doesn't depend on it

---

### Permissions Reference

**Minimum Required Permissions:**

| Task | Azure AD Role | Alternative Role |
|------|--------------|------------------|
| Read sensitivity labels | Compliance Data Administrator | Information Protection Reader |
| Read label policies | Compliance Data Administrator | Information Protection Reader |
| Read DLP policies | Compliance Data Administrator | Security Reader |
| Create labels | Information Protection Administrator | Compliance Administrator |
| Create policies | Information Protection Administrator | Compliance Administrator |
| Graph API access | Security Administrator | Compliance Administrator + API permissions |

**PowerShell Module Versions:**

```powershell
# Check installed versions
Get-Module -ListAvailable ExchangeOnlineManagement
Get-Module -ListAvailable Microsoft.Graph

# Update if needed
Update-Module ExchangeOnlineManagement -Force
Update-Module Microsoft.Graph -Force
```

**Recommended Versions:**
- `ExchangeOnlineManagement` ≥ 3.0.0
- `Microsoft.Graph` ≥ 2.0.0

---

## Script Enhancements Needed

Based on this documentation review, recommended improvements:

### High Priority

1. **Add Markdown Export Function**
   - Generate structured .md files with portal links
   - Include encryption/marking callouts
   - Embed Microsoft Learn references

2. **Capture Parent/Child Relationships**
   - Build label hierarchy tree
   - Show inheritance of protection settings
   - Visualize in Word/PowerPoint

3. **Add Auto-Labeling Policy Collection**
   - `Get-AutoSensitivityLabelPolicy`
   - `Get-AutoSensitivityLabelRule`
   - Distinguish from client-side auto-labeling

4. **Enhance Encryption Documentation**
   - Parse rights into readable format (not just JSON)
   - Highlight Owner vs. View-only permissions
   - Flag labels without encryption

5. **Add Portal Links to Word/PowerPoint**
   - Hyperlink each label/policy name to its portal page
   - Add "Learn More" links to Microsoft Learn docs

### Medium Priority

6. **Collect Label Analytics** (if available)
   - Label usage frequency
   - Most applied labels
   - User activity data

7. **Add Scope Information**
   - Identify labels scoped to Exchange only, etc.
   - Document adaptive scopes

8. **Improve DLP Rule Documentation**
   - Parse conditions into human-readable format
   - Create condition/action summary tables

9. **Add Configuration Comparison**
   - Compare two exports to identify drift
   - Highlight new/modified/deleted items

10. **Validation Checks**
    - Labels published but not in any policy
    - Policies with no users assigned
    - DLP policies in test mode > 30 days

---

## Conclusion

This documentation provides comprehensive guidance for:

✅ Understanding how the PowerShell script works  
✅ Interpreting Information Protection configurations  
✅ Using exported documentation formats  
✅ Manually recreating configurations from documentation  
✅ Troubleshooting common issues  

For questions or issues, consult the Microsoft Learn references throughout this document or engage Microsoft Support.

**Document Maintained By:** Microsoft Purview Configuration Team  
**Last Review Date:** December 5, 2025  
**Next Review:** Quarterly (March 2026)

