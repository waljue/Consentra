# Consentra - Export Entra OAuth2 Grants

PowerShell script that inventories Entra ID OAuth delegated consents, enriches them with service principal metadata, and produces both CSV and interactive HTML reports for security reviews.

## ✨ Features

- Queries `oauth2PermissionGrants` via Microsoft Graph and deduplicates grants per enterprise application.
- Pulls extended service principal metadata (publisher, verified publisher, service principal names, tags).
- Classifies apps as Microsoft, tenant-owned, or third party using multiple heuristics.
- Includes hidden enterprise applications by enriching the data set with the complete tenant service principal inventory.
- Highlights risky scopes (e.g., `*ReadWrite*`, `Directory.AccessAsUser.All`) and flags low-impact grants.
- Generates outputs:
  - **CSV** with raw data and metadata for spreadsheets or SIEM ingestion.
  - **HTML** dashboard with search, consent-type chips, Microsoft/third-party toggles, vendor tags, alphabetical sort toggle, and CSV export of any filtered view.

## ✅ Prerequisites

| Requirement | Notes |
| --- | --- |
| PowerShell 7.2+ (recommended) or Windows PowerShell 5.1 | Script was tested on PowerShell 7; older hosts should enforce TLS 1.2. |
| Microsoft Graph PowerShell SDK | Install with `Install-Module Microsoft.Graph.Authentication -Scope CurrentUser`. |
| Graph permissions | The signed-in identity must be able to consent to **Directory.Read.All** and **Application.Read.All** (admin consent recommended). |
| Internet access | Calls the Microsoft Graph `beta` and `v1.0` endpoints. |

> ℹ️ The script auto-connects if no existing Graph context is found, but you can connect manually (see below) to control the scope consent and profile selection.

## 🚀 Setup

```powershell
Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
Import-Module Microsoft.Graph.Authentication
# Optional: pick needed profile (beta is default when using Invoke-MgGraphRequest)
# Select-MgProfile -Name beta
Connect-MgGraph -Scopes 'Directory.Read.All','Application.Read.All'
```

If `Connect-MgGraph` prompts for admin consent, ensure you sign in with an account that can grant those scopes tenant-wide.

## ▶️ Usage

Run the script from the repo root (or provide the full path):

```powershell
# Export all grants
./consentra.ps1

# Filter to apps whose display name contains "ServiceNow"
./consentra.ps1 -AppNameContains "ServiceNow"
```

> 💡 On Windows PowerShell run `.\consentra.ps1`; on macOS/Linux under PowerShell 7 you can use `./consentra.ps1` as shown above.

Outputs are written to the current directory:

- `OAuth2_Consent_Apps_<timestamp>.csv`
- `OAuth2_Consent_Report_<timestamp>.html`

You can open the HTML file locally in any modern browser.

## 📊 HTML report overview

- **Status chips**: toggle all, critical, elevated, or low-impact apps.
- **Vendor chips**: quickly isolate Microsoft or verified third-party apps using the built-in classification.
- **Consent chips**: click on the `Admin Consent` / `User Consent` tags in a card to filter by consent type.
- **Search box**: type to filter by app display name; `⌘/` or `Ctrl+/` focuses the search.
- **Sort button**: `Sort A→Z` toggles between ascending and descending alphabetical order.
- **Metadata panel**: expand a card to review publisher, domains, service principal names, tags, and consent details.
- **Inline CSV export**: the `Export CSV` button downloads the currently filtered subset without re-running the script.

## 📁 CSV schema

The CSV contains the following key columns:

- `clientName`, `clientAppId`: display name and AppId of the enterprise application.
- `status`: `pass`, `warn`, or `fail` based on scope analysis (critical `ReadWrite` scopes drive `fail`).
- `isMicrosoft`, `isThirdParty`: vendor classification flags.
- `publisher`, `publisherDomain`, `verifiedPublisherName`, `verifiedPublisherId`, `appOwnerOrgId`.
- `servicePrincipalNames`, `servicePrincipalTags`, `servicePrincipalType`.
- `adminScopes`, `userScopes`: semicolon-separated `resource:scope` pairs.
- `userCount`: distinct user count for delegated grants.
- `tags`: the same set of label chips shown in the HTML report (admin/user consent, usage counts, vendor).

## 🛠️ Troubleshooting

- **Missing data**: ensure the account you use has sufficient permissions and that the tenant contains delegated OAuth grants.
- **Graph throttling**: the script batches service principal metadata requests in chunks of 20; if you hit throttling, re-run after a short delay.
- **Locale mismatches**: the HTML sorting uses invariant lowercase comparisons; display names with leading punctuation may group separately.
- **Consent errors**: if automatic `Connect-MgGraph` fails, run it manually before executing the script so you can resolve MFA or approval prompts.

## 🤝 Contributing

Feel free to open issues or pull requests with enhancements—ideas include additional report filters, support for application permissions, or wiring the exporter into scheduled automation (Azure Functions, GitHub Actions, etc.).
