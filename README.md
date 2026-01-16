# Entra Sync Plugin for XrmToolBox

A focused XrmToolBox plugin that forces Dataverse to refresh a user record by impersonating the user and executing `WhoAmI`.

## What it does

When a user is newly provisioned or group membership changes lag behind, Dataverse can take time to reflect updates. This plugin:

- Finds a user by email or UPN.
- Impersonates that user with `CallerId`.
- Executes `WhoAmI` to trigger a refresh.
- Displays the user summary, personal security roles, and assigned teams.

## Why `WhoAmI` is enough

`WhoAmI` is a lightweight, safe request that forces Dataverse to resolve the caller context. When the caller is impersonated, the platform must re-evaluate the user's identity, security context, and role/teams. That makes it a reliable and low-risk way to prompt a refresh without modifying data.

## UI overview

- Left: “Entra User Sync” card with instructions, prerequisites, and status.
- Right: resizable panes for User Summary, Personal Security Roles, and Assigned Teams.
- No pop-up windows; details update in place after a successful sync.

## Example scenario

You add Jane Doe to a security group that grants Dataverse access, but her roles are not showing yet. In XrmToolBox:

1) Open Entra Sync Plugin.  
2) Enter `jane.doe@contoso.com`.  
3) Click **Force Sync**.  
4) The panel on the right shows her updated roles and teams.

Expected time: 1-5 seconds depending on org size and network latency.

## Requirements

- XrmToolBox (recent version).
- .NET Framework 4.8 (XrmToolBox baseline).
- Permissions: the running user must have **Act on Behalf of Another User** (Delegate) in Dataverse.

## Installation

1) Build or download `EntraSyncPlugin.dll`.
2) Copy it to your XrmToolBox plugins folder:
   - `%APPDATA%\MscrmTools\XrmToolBox\Plugins`
3) Restart XrmToolBox.

If you use a portable XrmToolBox, copy the DLL to its `Plugins` folder and restart.

## Build from source

```powershell
dotnet build EntraSyncPlugin.csproj -c Debug
```

Output:

```
bin\Debug\net48\EntraSyncPlugin.dll
```

## Notes

- Entra group display names are not shown in this build. That requires Microsoft Graph auth, which is intentionally not included here.
- The plugin does not alter user data. It only impersonates and runs `WhoAmI`.

## License

MIT
