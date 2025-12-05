# Microsoft Purview Configuration Documentation Tool

A PowerShell-based tool for documenting and managing Microsoft Purview configuration settings across Microsoft 365 tenants (Entra ID directories).

> **Note:** This is an independent project by John Spaid and is **not an official Microsoft product**. While the author works at Microsoft, this tool is provided as-is under the MIT License without warranty or official support from Microsoft Corporation.

## Overview

This tool helps organizations document, track, and manage their Microsoft Purview configuration by leveraging PowerShell for Microsoft 365, Microsoft Purview, and the Microsoft Graph API. It provides automated configuration documentation and backup capabilities for compliance and governance requirements.

## Current Scope

The initial release focuses on two critical Purview components:

- **Information Protection Labels** - Sensitivity labels used to classify and protect organizational data
- **Data Loss Prevention (DLP) Policies** - Policies that prevent sensitive information from leaving the organization

## Planned Features

Future releases will expand to cover additional Microsoft Purview products and capabilities:

- Insider Risk Management
- Communication Compliance
- Information Barriers
- Records Management
- eDiscovery
- Audit solutions
- Compliance Manager assessments

## Prerequisites

- PowerShell 5.1 or PowerShell 7+
- Microsoft 365 administrator access (appropriate roles for Purview)
- Required PowerShell modules:
  - Microsoft.Graph
  - ExchangeOnlineManagement
  - PnP.PowerShell (for SharePoint/Teams configurations)

### Cross-Platform Compatibility

| Feature | Windows | macOS | Linux |
|---------|---------|-------|-------|
| **Core Functionality** | ✅ Full | ✅ Full | ✅ Full |
| JSON/CSV Export | ✅ | ✅ | ✅ |
| Markdown Generation | ✅ | ✅ | ✅ |
| Word (.docx) Export | ✅ | ❌ | ❌ |
| PowerPoint (.pptx) Export | ✅ | ❌ | ❌ |
| Graph API Enrichment | ✅ | ✅ | ✅ |

**Notes:**
- **macOS/Linux Users:** All core export functionality works perfectly, including JSON, CSV, and Markdown generation
- **Word/PowerPoint:** These formats require Windows with Microsoft Office installed (uses COM automation)
- **Recommended for macOS/Linux:** Use `-CreateMarkdown` parameter for human-readable documentation
- **All Modules Available:** `ExchangeOnlineManagement`, `Microsoft.Graph`, and `PnP.PowerShell` are all available on macOS and Linux via PowerShell Gallery

### macOS/Linux Example

```powershell
# Install modules (one-time setup)
Install-Module -Name Microsoft.Graph -Scope CurrentUser -Force
Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser -Force

# Run export with Markdown (no Office required)
./Export-PurviewConfiguration.ps1 `
  -OutputFolder "./output" `
  -TenantDisplayName "Your Organization" `
  -CreateMarkdown
```

## Installation

```powershell
# Clone the repository
git clone https://github.com/YOUR_USERNAME/purview-config-documentation.git
cd purview-config-documentation

# Install required modules
Install-Module -Name Microsoft.Graph -Scope CurrentUser -Force
Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser -Force
Install-Module -Name PnP.PowerShell -Scope CurrentUser -Force
```

## Usage

### Common Examples

#### Example 1: Quick Export (JSON and CSV only)

Export configurations to JSON and CSV files for backup or version control:

```powershell
.\Export-PurviewConfiguration.ps1 `
  -OutputFolder ".\backup\2025-12-05" `
  -TenantDisplayName "Contoso Corporation"
```

**What it does:**
- Connects to Security & Compliance PowerShell
- Exports sensitivity labels, label policies, auto-labeling policies, DLP policies and rules
- Generates JSON files for configuration-as-code
- Generates CSV files for spreadsheet analysis
- Creates metadata file with export details

---

#### Example 2: Markdown Documentation for Wiki

Generate human-readable documentation with portal links for your internal wiki or documentation site:

```powershell
.\Export-PurviewConfiguration.ps1 `
  -OutputFolder ".\docs\purview" `
  -TenantDisplayName "Contoso Corporation" `
  -CreateMarkdown
```

**What it does:**
- All JSON/CSV exports (Example 1)
- **Plus:** Markdown (.md) file with:
  - Hierarchical label structure
  - Encryption status indicators
  - Direct links to Purview portal pages
  - Microsoft Learn documentation references
  - Suitable for static site generators or wikis

---

#### Example 3: Audit Report (Word Document)

Create a comprehensive audit report for compliance reviews:

```powershell
.\Export-PurviewConfiguration.ps1 `
  -OutputFolder ".\reports\audit-2025-Q4" `
  -TenantDisplayName "Contoso Corporation" `
  -CreateWord
```

**What it does:**
- All JSON/CSV exports (Example 1)
- **Plus:** Word (.docx) document with:
  - Formatted tables for all configurations
  - Portal and Learn links embedded
  - Suitable for printing and formal audits
  - **Note:** Requires Windows with Microsoft Office installed

---

#### Example 4: Executive Presentation

Generate a PowerPoint presentation for leadership briefings:

```powershell
.\Export-PurviewConfiguration.ps1 `
  -OutputFolder ".\presentations" `
  -TenantDisplayName "Contoso Corporation" `
  -CreatePowerPoint
```

**What it does:**
- All JSON/CSV exports (Example 1)
- **Plus:** PowerPoint (.pptx) with:
  - Summary slides with key metrics
  - Encryption and protection highlights
  - High-level overview (less technical detail)
  - **Note:** Requires Windows with Microsoft Office installed

---

#### Example 5: Complete Documentation Suite

Generate all formats for comprehensive documentation:

```powershell
.\Export-PurviewConfiguration.ps1 `
  -OutputFolder ".\complete-export" `
  -TenantDisplayName "Contoso Corporation" `
  -CreateMarkdown `
  -CreateWord `
  -CreatePowerPoint
```

**What it does:**
- All JSON/CSV exports
- Markdown documentation
- Word audit report
- PowerPoint presentation
- **Best for:** Tenant migration preparation, annual compliance audits, configuration baselines

---

#### Example 6: With Microsoft Graph Enrichment

Include additional data from Microsoft Graph beta endpoints:

```powershell
.\Export-PurviewConfiguration.ps1 `
  -OutputFolder ".\enriched-export" `
  -TenantDisplayName "Contoso Corporation" `
  -CreateMarkdown `
  -IncludeGraphEnrichment
```

**What it does:**
- All standard exports
- **Plus:** Attempts to retrieve additional data via Microsoft Graph API
- Requires admin consent for Graph scopes
- **Note:** Beta endpoints may change; use for supplemental data only

---

### Parameters Reference

| Parameter | Required | Type | Description |
|-----------|----------|------|-------------|
| `-OutputFolder` | Yes | String | Path where output files will be saved (created if doesn't exist) |
| `-TenantDisplayName` | Yes | String | Friendly name for your tenant (used in report titles) |
| `-CreateMarkdown` | No | Switch | Generate Markdown (.md) documentation with portal links |
| `-CreateWord` | No | Switch | Generate Word (.docx) report (requires Windows + Office) |
| `-CreatePowerPoint` | No | Switch | Generate PowerPoint (.pptx) presentation (requires Windows + Office) |
| `-IncludeGraphEnrichment` | No | Switch | Attempt to retrieve additional data via Microsoft Graph beta |

### Output Files

The script generates timestamped files including:

- **JSON Files** - Machine-readable configuration backup
  - `PurviewDoc_TIMESTAMP.labels.json`
  - `PurviewDoc_TIMESTAMP.labelPolicies.json`
  - `PurviewDoc_TIMESTAMP.autoLabelPolicies.json`
  - `PurviewDoc_TIMESTAMP.dlpPolicies.json`
  - `PurviewDoc_TIMESTAMP.dlpRules.json`

- **CSV Files** - Spreadsheet-compatible exports
  - `PurviewDoc_TIMESTAMP.labels.csv`
  - `PurviewDoc_TIMESTAMP.labelPolicies.csv`
  - And corresponding CSV files for all policy types

- **Markdown** - Human-readable documentation with portal links
  - `PurviewDoc_TIMESTAMP.md`

- **Word Document** - Comprehensive report with hyperlinks
  - `PurviewDoc_TIMESTAMP.docx`

- **PowerPoint** - Executive summary presentation
  - `PurviewDoc_TIMESTAMP.pptx`

### Key Features

✅ **Comprehensive Coverage**
- Sensitivity labels with parent/child relationships
- Label policies and user assignments
- Auto-labeling policies (service-side)
- DLP policies and rules with conditions/actions

✅ **Encryption & Protection Highlights**
- Identifies labels with encryption enabled
- Documents access controls and rights management
- Shows visual markings (headers, footers, watermarks)

✅ **Documentation Links**
- Direct links to Purview portal configuration pages
- Microsoft Learn documentation references
- Enables manual configuration from documentation

✅ **Multiple Output Formats**
- JSON for automation and version control
- CSV for spreadsheet analysis
- Markdown for wikis and static sites
- Word for comprehensive audits
- PowerPoint for executive briefings

## Output Formats

Configuration data is exported in multiple structured formats:

- **JSON** - Machine-readable configuration backup and version control
- **CSV** - Spreadsheet-compatible tabular data for analysis
- **Markdown** - Human-readable documentation with hyperlinks to portal and Learn docs
- **Word (.docx)** - Comprehensive configuration report suitable for audits
- **PowerPoint (.pptx)** - Executive summary with key metrics and highlights

## Security Considerations

- Credentials are never stored in scripts or output files
- Uses modern authentication with Azure AD
- Supports Conditional Access policies and MFA
- Output files may contain sensitive configuration data - handle appropriately

## Contributing

Contributions are welcome! Please submit pull requests or open issues for bugs and feature requests.

## License

GNU General Public License v3.0

Copyright (C) 2025 John Spaid

This program is free software: you can redistribute it and/or modify
it under the terms of the GNU General Public License as published by
the Free Software Foundation, either version 3 of the License, or
(at your option) any later version.

This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
GNU General Public License for more details.

You should have received a copy of the GNU General Public License
along with this program. If not, see <https://www.gnu.org/licenses/>.

### Copyleft Notice

This software is licensed under GPL v3, which means:
- ✅ You can freely use, modify, and distribute this software
- ✅ You can use it for commercial purposes
- ⚠️ **Any derivative works must also be released under GPL v3**
- ⚠️ **Source code must be made available for any distributed modifications**

This ensures the software and all improvements remain free and open source forever.

## Author

**John Spaid**

This is an independent, personal project. The author is employed by Microsoft but this tool is not affiliated with, endorsed by, or supported by Microsoft Corporation. It is provided as a community contribution under the MIT License.

## Project Status

🚧 **Initial Development** - This project is in early development stages. The core functionality for Information Protection labels and DLP policies is being implemented first.
