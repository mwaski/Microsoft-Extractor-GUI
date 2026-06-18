# M.E.G. — Microsoft Extractor GUI

[![Build EXE](https://github.com/mwaski-SureFire-Cyber/M365_App/actions/workflows/build.yml/badge.svg)](https://github.com/mwaski-SureFire-Cyber/M365_App/actions/workflows/build.yml)

A Windows GUI front-end for Microsoft 365 / Entra ID forensic log collection. M.E.G. provisions a hardened, certificate-authenticated Entra app registration in a tenant, then gives you a point-and-click command builder for the [Invictus-IR Microsoft-Extractor-Suite](https://github.com/invictus-ir/Microsoft-Extractor-Suite) — so you can scope, preview, and run incident-response collections without hand-writing PowerShell.

> Built for SureFire Cyber DFIR engagements. Everything ships in a single self-contained PowerShell script: `Microsoft_Extractor_GUI.ps1`.

![M.E.G. — Microsoft Extractor GUI](docs/screenshot.png)

<!-- Add a screenshot at docs/screenshot.png (e.g. the M365 Command Builder tab). -->

---

## What it does

- **Tenant onboarding** — Creates an Entra ID app registration (`Microsoft Extractor GUI`) with a self-signed certificate, assigns the Microsoft Graph + Exchange Online application permissions needed for collection, and grants admin consent programmatically.
- **Command builder** — A categorized checklist of collection cmdlets (Exchange Online, Graph identity/devices/policies, audit logs, Unified Audit Log) that generates ready-to-run, correctly-quoted PowerShell commands.
- **Global & per-cmdlet options** — Configure output directory, encoding, date ranges, user filters, and cmdlet-specific parameters from the UI.
- **Unified Audit Log helpers** — Pull UAL via Graph (`Get-UALGraph`) by last-N-days, an explicit date range, or "pull all," with an optional **Triage mode** that filters to a curated operations list.
- **Setup / Update / Uninstall** — Install the required PowerShell modules (Microsoft-Extractor-Suite, Microsoft.Graph, ExchangeOnlineManagement), update them, or tear the app registration back down.
- **Pre-flight checks** — A "Test UAL" action confirms unified audit logging is actually enabled in the tenant before you build a collection.

## How it works

The GUI manages tenant configs under a `ConfigRoot` (default `C:\Microsoft_Extractor_GUI`), each holding the tenant's app ID, tenant ID, certificate thumbprint, primary domain, and output path. Collections authenticate to Graph and Exchange Online using the app's certificate — no interactive sign-in or stored secrets per run.

The installer logic is embedded in the same script (single source of truth) and is launched in its own elevated PowerShell window when you click **Install / Setup Tenant**.

## Requirements

- **Windows** with PowerShell.
- **PowerShell 7 (`pwsh`) is strongly recommended.** `Get-UALGraph` is known to fail with `BadRequest` under Windows PowerShell 5.1 (a JSON date-serialization bug). The GUI warns you when it falls back to `powershell.exe`.
- A tenant account with sufficient privileges to create an app registration and grant admin consent during setup (Global Administrator or equivalent).
- The following modules (the **M.E.G Setup** tab can install them for you):
  - `Microsoft-Extractor-Suite`
  - `Microsoft.Graph`
  - `ExchangeOnlineManagement`

## Usage

```powershell
# Normal launch (Commercial cloud)
.\Microsoft_Extractor_GUI.ps1

# Rotate the certificate on an existing app registration
.\Microsoft_Extractor_GUI.ps1 -NewCert

# Keep the console window visible behind the GUI (useful for debugging)
.\Microsoft_Extractor_GUI.ps1 -ShowConsole

# Use a custom config/output root
.\Microsoft_Extractor_GUI.ps1 -ConfigRoot "D:\Cases\Contoso"

# Show help
.\Microsoft_Extractor_GUI.ps1 -Help
```

### Typical workflow

1. **Install / Setup Tenant** — provision the app registration and certificate, grant consent.
2. **Test UAL** — confirm unified audit logging is enabled.
3. **M365 Command Builder** — tick the log collections you need.
4. **Global / Per-Cmdlet Options** — set the output directory, date range, and any parameters.
5. **Build Commands** — preview the generated PowerShell.
6. **Execute** — run the collection (in PowerShell 7 when available).

## Permissions requested

Setup assigns the following **application** (app-only) permissions and grants admin consent.

**Microsoft Graph**

| Permission | Purpose |
| --- | --- |
| `Application.ReadWrite.All`, `Application.Read.All` | Manage the app registration |
| `AuditLog.Read.All`, `AuditLogsQuery.Read.All` | Entra audit & sign-in logs, UAL via Graph |
| `Directory.Read.All` | Directory objects |
| `IdentityRiskEvent.Read.All`, `IdentityRiskyUser.Read.All` | Risk detections / risky users |
| `Mail.ReadBasic.All`, `Mail.ReadWrite` | Mail collection |
| `MailboxSettings.Read` | Mailbox rules / settings |
| `Policy.Read.All` | Conditional Access & policy posture |
| `UserAuthenticationMethod.Read.All` | MFA / auth methods |
| `User.Read.All`, `Group.Read.All`, `Device.Read.All` | Users, groups, devices |
| `RoleManagement.Read.Directory`, `RoleEligibilitySchedule.Read.Directory`, `RoleAssignmentSchedule.Read.Directory` | Role activity / PIM |
| `SecurityEvents.Read.All` | Security alerts / secure score |

**Office 365 Exchange Online**

| Permission | Purpose |
| --- | --- |
| `Exchange.ManageAsApp` | App-only Exchange Online connection |
| `Mail.ReadWrite` | Mail item access |

## Collections available

- **Exchange Online — Mailbox:** audit status, permissions, inbox rules (`Get-MailboxRules` / `Show-MailboxRules`)
- **Exchange Online — Audit logs:** `Get-AdminAuditLog`, `Get-MailboxAuditLog`, classic Unified Audit Log (`Get-UAL`)
- **Exchange Online — Email:** `Get-Email`, `Get-Attachment`, `Show-Email`
- **Exchange Online — Mail items accessed:** `Get-Sessions`, `Get-MessageIDs`
- **Exchange Online — Transport:** transport rules, 90-day message trace
- **Graph — Identity:** users, admins, MFA, risky detections/users, OAuth permissions
- **Graph — Devices & Groups:** devices, groups, group members, dynamic groups
- **Graph — Policies & Posture:** Conditional Access, security defaults, secure score, security alerts
- **Graph — Roles & Licenses:** role activity, PIM assignments, licenses
- **Graph — Logs:** Entra audit logs, Entra sign-in logs, mailbox rules via Graph
- **Unified Audit Log (Graph):** date-range / last-N-days / pull-all, with optional Triage mode

## Building the EXE

A GitHub Actions workflow ([.github/workflows/build.yml](.github/workflows/build.yml)) compiles the script to an EXE with [PS2EXE](https://github.com/MScholtes/PS2EXE) on push to `main` and on `v*` tags (tags also publish a GitHub Release).

> **Heads up:** PS2EXE output is frequently flagged and quarantined by antivirus/EDR. For trusted use you may need to allow-list the binary, or just run the `.ps1` directly.

## License

Internal SureFire Cyber tooling. See repository for details.
