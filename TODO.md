# TODO List - Purview Configuration Documentation Tool

**Last Updated:** December 5, 2025

---

## High Priority

### 1. Configuration-as-Code Export for CI/CD Pipelines

**Status:** 🔴 Not Started  
**Priority:** High  
**Estimated Effort:** 3-5 days

**Description:**
Add functionality to export Purview configurations in a format optimized for source control (Git) and CI/CD deployment pipelines. This would enable Infrastructure-as-Code (IaC) practices for Purview configurations.

**Requirements:**
- Export configurations in a structured, source-control-friendly format
- Support for declarative configuration files (e.g., YAML, JSON schema)
- Generate deployment scripts that can apply configurations to target tenants
- Include validation and drift detection capabilities
- Support for environment-specific overrides (dev, test, prod)

**Proposed Implementation:**

#### Export Format Options:
1. **JSON Schema-based** - Structured JSON with schema validation
2. **YAML** - Human-readable, merge-friendly format
3. **Terraform/OpenTofu** - HashiCorp Configuration Language (HCL)
4. **PowerShell DSC** - Desired State Configuration

#### File Structure:
```
/purview-config/
├── labels/
│   ├── parent-labels.yaml
│   ├── sublabels/
│   │   ├── confidential-finance.yaml
│   │   ├── confidential-hr.yaml
│   │   └── confidential-legal.yaml
│   └── schemas/
│       └── label-schema.json
├── policies/
│   ├── label-policies/
│   │   ├── global-policy.yaml
│   │   └── finance-dept-policy.yaml
│   ├── auto-labeling-policies/
│   │   └── credit-card-auto-label.yaml
│   └── dlp-policies/
│       ├── financial-data-dlp.yaml
│       └── pii-protection-dlp.yaml
├── environments/
│   ├── dev.env.yaml
│   ├── test.env.yaml
│   └── prod.env.yaml
└── deploy/
    ├── Import-PurviewConfiguration.ps1
    ├── Test-PurviewConfiguration.ps1
    └── Compare-PurviewConfiguration.ps1
```

#### Key Features:

**Export Functionality:**
- `-ExportForSourceControl` parameter
- Breaks down configurations into modular files
- Removes environment-specific data (GUIDs, timestamps)
- Normalizes sensitive data (user/group references by name, not GUID)
- Includes dependency tracking (labels referenced by policies)

**Import/Deploy Script:**
```powershell
.\Import-PurviewConfiguration.ps1 `
  -ConfigPath ".\purview-config" `
  -Environment "prod" `
  -WhatIf
```

**Drift Detection:**
```powershell
.\Compare-PurviewConfiguration.ps1 `
  -SourcePath ".\purview-config" `
  -TargetTenant "contoso.onmicrosoft.com" `
  -GenerateReport
```

**Validation:**
- JSON schema validation before deployment
- Dependency resolution (ensure parent labels exist before sublabels)
- Conflict detection (duplicate names, priority conflicts)
- Permission verification (user/group existence)

#### CI/CD Integration Examples:

**GitHub Actions:**
```yaml
name: Deploy Purview Configuration
on:
  push:
    branches: [main]
    paths: ['purview-config/**']

jobs:
  deploy:
    runs-on: windows-latest
    steps:
      - uses: actions/checkout@v3
      - name: Validate Configuration
        run: .\deploy\Test-PurviewConfiguration.ps1 -ConfigPath .\purview-config
      - name: Deploy to Production
        run: .\deploy\Import-PurviewConfiguration.ps1 -ConfigPath .\purview-config -Environment prod
        env:
          PURVIEW_CLIENT_ID: ${{ secrets.PURVIEW_CLIENT_ID }}
          PURVIEW_CLIENT_SECRET: ${{ secrets.PURVIEW_CLIENT_SECRET }}
```

**Azure DevOps Pipeline:**
```yaml
trigger:
  branches:
    include:
      - main
  paths:
    include:
      - purview-config/*

pool:
  vmImage: 'windows-latest'

steps:
- task: PowerShell@2
  displayName: 'Validate Purview Configuration'
  inputs:
    filePath: 'deploy/Test-PurviewConfiguration.ps1'
    arguments: '-ConfigPath $(Build.SourcesDirectory)/purview-config'

- task: PowerShell@2
  displayName: 'Deploy Purview Configuration'
  inputs:
    filePath: 'deploy/Import-PurviewConfiguration.ps1'
    arguments: '-ConfigPath $(Build.SourcesDirectory)/purview-config -Environment prod'
  env:
    PURVIEW_CLIENT_ID: $(PurviewClientId)
    PURVIEW_CLIENT_SECRET: $(PurviewClientSecret)
```

#### Benefits:
- ✅ Version control for compliance configurations
- ✅ Audit trail of all configuration changes
- ✅ Automated deployment across environments
- ✅ Configuration drift detection and remediation
- ✅ Disaster recovery (restore from Git)
- ✅ Multi-tenant management (same config, different tenants)

#### Challenges to Address:
- **GUIDs:** Parent label references, policy assignments use GUIDs - need name-based resolution
- **User/Group References:** Must resolve by name/UPN, not GUID
- **Encryption Templates:** RMS templates may need manual creation/mapping
- **Conditional Access:** External dependencies that can't be auto-created
- **Permissions:** Service principal vs. user authentication for CI/CD

#### Related Tools/References:
- [Microsoft365DSC](https://microsoft365dsc.com/) - Existing DSC resource for M365
- [Export-M365DSCConfiguration](https://microsoft365dsc.com/user-guide/get-started/export-configuration/) - Similar export concept
- [Terraform Azure AD Provider](https://registry.terraform.io/providers/hashicorp/azuread/latest/docs)
- [PowerShell Desired State Configuration](https://learn.microsoft.com/powershell/dsc/)

---

## Medium Priority

### 2. Label Analytics and Usage Reporting

**Status:** 🔴 Not Started  
**Priority:** Medium  
**Estimated Effort:** 2-3 days

**Description:**
Collect and report on label usage analytics to understand which labels are most used, by whom, and on what content types.

**Requirements:**
- Integrate with Azure Information Protection analytics
- Report on label application frequency
- Identify unused labels (candidates for retirement)
- User adoption metrics by department/group
- Content type analysis (Word, Excel, Email, SharePoint)

---

### 3. Configuration Validation and Best Practices Analyzer

**Status:** 🔴 Not Started  
**Priority:** Medium  
**Estimated Effort:** 2-3 days

**Description:**
Analyze exported configurations against Microsoft best practices and identify potential issues.

**Checks to Include:**
- Labels published but not in any policy
- Policies with no users assigned
- DLP policies in test mode for >30 days
- Labels without encryption (security gap)
- Duplicate or conflicting priorities
- Missing mandatory labeling settings
- Overly permissive encryption settings
- Orphaned sublabels (parent disabled but children enabled)

---

### 4. Multi-Tenant Configuration Comparison

**Status:** 🔴 Not Started  
**Priority:** Medium  
**Estimated Effort:** 2-3 days

**Description:**
Compare configurations across multiple tenants to identify inconsistencies or ensure standard configurations are applied.

**Use Cases:**
- Subsidiaries that should have same configuration
- Dev/Test/Prod environment consistency
- Pre/post-merger tenant comparison
- Franchise or partner tenant management

---

### 5. Incremental Export (Delta Detection)

**Status:** 🔴 Not Started  
**Priority:** Medium  
**Estimated Effort:** 1-2 days

**Description:**
Only export configurations that have changed since the last export to reduce file size and improve performance.

**Implementation:**
- Store last export metadata with timestamps
- Compare `WhenChanged` timestamps
- Generate delta reports showing what changed
- Support for "export only modified since [date]"

---

## Low Priority

### 6. Scheduled Export Automation

**Status:** 🔴 Not Started  
**Priority:** Low  
**Estimated Effort:** 1 day

**Description:**
Create scheduled task templates or automation scripts for regular exports (daily, weekly, monthly).

---

### 7. Encryption Rights Visualization

**Status:** 🔴 Not Started  
**Priority:** Low  
**Estimated Effort:** 1-2 days

**Description:**
Generate visual diagrams showing who has what access to encrypted labels. Export as PNG, SVG, or Mermaid diagram format.

---

### 8. Label Recommendation Engine

**Status:** 🔴 Not Started  
**Priority:** Low  
**Estimated Effort:** 3-4 days

**Description:**
Based on existing labels and policies, suggest new labels, conditions, or auto-labeling rules that might be beneficial.

**Suggestions might include:**
- "You have labels for Finance and HR but not Legal"
- "Consider auto-labeling based on folder location"
- "Your 'Confidential' label has no encryption - consider adding"

---

### 9. PowerShell Module Packaging

**Status:** 🔴 Not Started  
**Priority:** Low  
**Estimated Effort:** 2-3 days

**Description:**
Package the scripts as a proper PowerShell module with manifest, help documentation, and publish to PowerShell Gallery.

**Module Name:** `PurviewConfigManagement` or `PurviewDocumentation`

**Cmdlets:**
- `Export-PurviewConfiguration`
- `Import-PurviewConfiguration`
- `Compare-PurviewConfiguration`
- `Test-PurviewConfiguration`
- `Get-PurviewConfigurationDrift`

---

### 10. GUI / Web Interface

**Status:** 🔴 Not Started  
**Priority:** Low  
**Estimated Effort:** 5-7 days

**Description:**
Create a web-based interface for non-PowerShell users to export, view, and compare configurations.

**Technology Options:**
- PowerShell Universal Dashboard
- Blazor web app
- Electron app with PowerShell backend

---

## Completed

_No items completed yet._

---

## Ideas / Future Considerations

- **Microsoft Purview Compliance Manager integration** - Export assessment results and compliance scores
- **Insider Risk Management configuration export** - Policy settings and indicators
- **Communication Compliance policy export** - Supervision policies and conditions
- **Records Management export** - Retention labels and policies
- **eDiscovery case configuration export** - Case settings and search queries
- **Audit log analysis** - Correlate configuration changes with audit events
- **Sensitivity label templates** - Pre-built label configurations for common industries (healthcare, finance, legal)
- **Cost estimation** - Calculate licensing requirements based on configuration complexity
- **Mobile app configuration** - Document mobile-specific label settings and behaviors
- **Third-party integration guides** - Salesforce, ServiceNow, SAP sensitivity label integration

---

## Contributing

To add a TODO item:
1. Choose appropriate priority (High/Medium/Low)
2. Provide clear description and requirements
3. Estimate effort if possible
4. Link to relevant Microsoft documentation
5. Update status as work progresses: 🔴 Not Started → 🟡 In Progress → 🟢 Completed

