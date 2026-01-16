# Entra Sync Plugin for XrmToolBox

A standalone XrmToolBox plugin designed to force user synchronization between Entra ID (Azure AD) and Microsoft Dataverse.

## Overview

This tool addresses the latency in Just-In-Time (JIT) user provisioning. By impersonating a target user and executing a `WhoAmI` request, it forces the Dataverse platform to refresh the user's group memberships and profile data immediately.

## Features

*   **User Resolution**: Resolves SystemUserId from Email or UPN.
*   **Impersonation**: Uses `CallerId` impersonation to trigger sync events.
*   **Simple UI**: WinForms-based interface compatible with XrmToolBox.

## Requirements

*   **XrmToolBox**: Latest version.
*   **Permissions**: The running user must have the "Act on Behalf of Another User" (Delegate) privilege in Dataverse.

## Installation

1.  Download the latest release DLL.
2.  Copy `EntraSyncPlugin.dll` to your XrmToolBox `Plugins` folder.
3.  Restart XrmToolBox.

## Build from Source

1.  Clone the repository.
2.  Run `dotnet build`.
3.  The output DLL will be in `bin/Debug/net48`.

## License

MIT
