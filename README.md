🛡️ ASR Exception Manager (Microsoft Defender for Endpoint)
📌 Overview

ASR Exception Manager is a PowerShell-based GUI tool designed to simplify the management of Attack Surface Reduction (ASR) exclusions in Microsoft Defender for Endpoint via Microsoft Graph API.

This tool allows security administrators to:

🔍 Query ASR events (Audit / Block)

📊 Visualize triggered ASR rules

🎯 Apply exclusions directly to Intune policies

⚙️ Manage both per-rule and global ASR exclusions

🔄 Automatically handle unknown ASR rules

🧠 Correlate Defender events with Intune devices (FQDN → hostname fallback)

🚀 Features

GUI-based interface (no need for manual scripting)

Direct integration with Microsoft Graph API

Supports:

ASR rule-specific exclusions

Global exclusions (Attack Surface Reduction Only Exclusions)

Automatic fallback for:

Unknown ASR Rule

Smart device resolution:

Handles FQDN vs hostname mismatch

Export events to CSV

Built-in troubleshooting (JSON dump)

🧩 How It Works

Connects to Microsoft Graph using App Registration

Queries Defender Advanced Hunting for ASR events

Maps events to Intune-managed devices

Identifies related:

Azure AD groups

Intune configuration policies

Allows applying exclusions:

Per ASR rule

Globally

Updates policy via Graph API (configurationPolicies)

🔐 Required Permissions (App Registration)

The application must be configured with the following Application permissions:

Microsoft Graph

ThreatHunting.Read.All

DeviceManagementManagedDevices.Read.All

DeviceManagementConfiguration.ReadWrite.All

Group.Read.All

Directory.Read.All

⚠️ Admin consent is required for all permissions.

⚙️ Requirements

PowerShell 5.1 or higher

Microsoft Defender for Endpoint (enabled)

Intune (Microsoft Endpoint Manager)

Azure AD App Registration (Client ID / Secret / Tenant ID)

🖥️ Usage

Launch the script:

.\ASR-Exception-Manager.ps1

Enter:

Tenant ID

Client ID

Client Secret

Click Connect

Click Search ASR Events

Select an event

Apply exclusion:

✔ Per-rule (optional)

✔ Global exclusion (optional or automatic)

🧠 Smart Logic
Unknown ASR Rule Handling

If a rule cannot be identified:

The tool automatically applies exclusion to:

Attack Surface Reduction Only Exclusions

UI toggle is automatically enabled and locked

Device Resolution Fix

Handles mismatch between:

Defender → device.domain.local

Intune → device

Fallback logic:

Full FQDN → Hostname
⚠️ Important Notes

Changes are applied directly to Intune policies

A single policy may impact multiple devices

Always validate before applying exclusions in production

📤 Export

Events can be exported to CSV for reporting or auditing

🧪 Troubleshooting

Use the built-in feature:

"Dump setting JSON"

This helps analyze policy structure when exclusions fail.

🤝 Contributing

Contributions, improvements, and suggestions are welcome!

Feel free to open issues or submit pull requests.

📄 Disclaimer

This tool is provided as-is without warranty.
Use at your own risk in production environments.

👨‍💻 Author

Created for the community ❤️
to simplify ASR management in enterprise environments.
