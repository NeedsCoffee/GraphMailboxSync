# GraphMailboxSync

`Copy-GraphMailboxItems.ps1` copies a mailbox folder subtree from one Exchange Online mailbox to another by using Microsoft Graph application authentication with a certificate from `Cert:\CurrentUser\My`.

The script can also load default values from a local `.env` file. Command-line parameters always win over `.env` values. `TenantId`, `ClientId`, and `CertificateThumbprint` must be provided on the command line or in `.env`.

## What it does

- Authenticates as an application with a certificate thumbprint.
- Resolves each user's Exchange mailbox ID through Graph.
- Walks the source folder tree recursively.
- Creates matching folders in the target mailbox only when needed for copied content by default.
- Runs a preflight check for ambiguous duplicate target folders before copying.
- Exports mailbox items in full fidelity in batches of up to 20.
- Imports each exported item into the target mailbox.
- Shows a confirmation summary, including estimated folder and item counts and where key values came from, before proceeding unless `-Force` is specified.
- Supports a preflight-only mode that validates the copy plan without importing anything.
- Supports an overlay mode that merges the source mailbox root into the target mailbox structure without creating a container folder.
- Can load credentials, mailbox paths, and default switches from a `.env` file.
- Supports `-WhatIf` for a dry run.
- Can include only specific source folders or exclude specific source folders.
- Can copy only items created within an optional date range.
- Always skips `Journal`, `Conversation History`, and `RSS Subscriptions`.
- Treats Calendar, Tasks, Notes, and Contacts copies specially so they target the matching root in the destination mailbox.

## Why this approach

This script uses the Microsoft Graph mailbox import/export preview APIs instead of the regular `/users/{id}/messages` endpoints. That matters because these APIs preserve mailbox items in full fidelity and support restoring them into a different mailbox.

## Required Microsoft Graph application permissions

At minimum, the app registration should have:

- `MailboxFolder.ReadWrite.All`
- `MailboxItem.ImportExport.All`
- `User.Read.All`

If your tenant uses Exchange application RBAC or application access policies, the app also needs to be allowed to the specific source and target mailboxes you want to touch.

## .env configuration

The repo includes:

- `.env.example` as a template you can copy or compare against
- `.env` for local defaults on this machine

Supported `.env` keys include:

- `TENANT_ID`
- `CLIENT_ID`
- `CERTIFICATE_THUMBPRINT`
- `SOURCE_USER_PRINCIPAL_NAME`
- `TARGET_USER_PRINCIPAL_NAME`
- `SOURCE_FOLDER_PATH`
- `TARGET_FOLDER_PATH`
- `IMPORT_DIRECTLY_INTO_TARGET_FOLDER`
- `OVERLAY_MODE`
- `COPY_EMPTY_FOLDERS`
- `INCLUDE_FOLDER_PATH`
- `EXCLUDE_FOLDER_PATH`
- `OLDEST`
- `NEWEST`
- `PREFLIGHT_ONLY`
- `FORCE`
- `EXPORT_BATCH_SIZE`

Boolean values accept `true`/`false`, `yes`/`no`, `1`/`0`, and similar forms. Multi-value include or exclude paths should be separated with `;`. If `SOURCE_FOLDER_PATH` is omitted, the script defaults the source root to `\`.

## Command-line parameters

- `-SourceUserPrincipalName <string>`: Required unless provided in `.env`. Source mailbox user principal name.
- `-TargetUserPrincipalName <string>`: Required unless provided in `.env`. Target mailbox user principal name.
- `-SourceFolderPath <string>`: Optional. Source mailbox folder path to copy. Defaults to `\`.
- `-TargetFolderPath <string>`: Optional. Target mailbox folder path. Defaults to an empty string.
- `-TenantId <string>`: Required unless provided in `.env`. Microsoft Entra tenant ID.
- `-ClientId <string>`: Required unless provided in `.env`. App registration client ID.
- `-CertificateThumbprint <string>`: Required unless provided in `.env`. Certificate thumbprint looked up in `Cert:\CurrentUser\My`.
- `-ImportDirectlyIntoTargetFolder`: Optional switch. Import items directly into the selected target folder instead of creating a same-named container folder.
- `-OverlayMode`: Optional switch. Merge the source root directly into the target structure.
- `-CopyEmptyFolders`: Optional switch. Create empty folders in the destination too.
- `-IncludeFolderPath <string[]>`: Optional. Copy only the listed subfolders.
- `-ExcludeFolderPath <string[]>`: Optional. Skip the listed subfolders.
- `-Oldest <string>`: Optional. Copy only items created on or after this date or timestamp.
- `-Newest <string>`: Optional. Copy only items created on or before this date or timestamp.
- `-PreflightOnly`: Optional switch. Validate the plan without importing any items.
- `-Force`: Optional switch. Skip the confirmation prompt.
- `-EnvFile <string>`: Optional. Path to a `.env` file to load. Defaults to `.env`. Pass `''` to disable `.env` loading.
- `-ExportBatchSize <int>`: Optional. Number of item IDs to export per request. Defaults to `20` and must be between `1` and `20`.

Use a different file at runtime if needed:

```powershell
.\Copy-GraphMailboxItems.ps1 `
  -EnvFile '.env.migration-a' `
  -SourceUserPrincipalName 'source@contoso.com' `
  -TargetUserPrincipalName 'target@contoso.com' `
  -SourceFolderPath 'Inbox' `
  -PreflightOnly `
  -Force
```

To run without any `.env` file at all, pass `-EnvFile ''` and provide the required auth settings on the command line.

## Example

Preview the copy first:

```powershell
.\Copy-GraphMailboxItems.ps1 `
  -SourceUserPrincipalName 'source@contoso.com' `
  -TargetUserPrincipalName 'target@contoso.com' `
  -SourceFolderPath 'Inbox\Projects\FY26' `
  -TargetFolderPath 'Inbox' `
  -WhatIf
```

Run only the preflight validation without copying any items:

```powershell
.\Copy-GraphMailboxItems.ps1 `
  -SourceUserPrincipalName 'source@contoso.com' `
  -TargetUserPrincipalName 'target@contoso.com' `
  -SourceFolderPath 'Inbox\Projects\FY26' `
  -TargetFolderPath 'Inbox' `
  -PreflightOnly
```

Then run it for real:

```powershell
.\Copy-GraphMailboxItems.ps1 `
  -SourceUserPrincipalName 'source@contoso.com' `
  -TargetUserPrincipalName 'target@contoso.com' `
  -SourceFolderPath 'Inbox\Projects\FY26' `
  -TargetFolderPath 'Inbox'
```

To skip the confirmation prompt for an unattended run:

```powershell
.\Copy-GraphMailboxItems.ps1 `
  -SourceUserPrincipalName 'source@contoso.com' `
  -TargetUserPrincipalName 'target@contoso.com' `
  -SourceFolderPath 'Inbox\Projects\FY26' `
  -TargetFolderPath 'Inbox' `
  -Force
```

To merge a whole mailbox root directly into the target mailbox structure without creating a source-root container folder:

```powershell
.\Copy-GraphMailboxItems.ps1 `
  -SourceUserPrincipalName 'source@contoso.com' `
  -TargetUserPrincipalName 'target@contoso.com' `
  -SourceFolderPath '\' `
  -OverlayMode `
  -WhatIf
```

To import the source folder's contents directly into the selected target folder instead of creating a same-named container folder:

```powershell
.\Copy-GraphMailboxItems.ps1 `
  -SourceUserPrincipalName 'source@contoso.com' `
  -TargetUserPrincipalName 'target@contoso.com' `
  -SourceFolderPath 'Inbox\Projects\FY26' `
  -TargetFolderPath 'Inbox\Archive' `
  -ImportDirectlyIntoTargetFolder
```

To preserve empty folders too:

```powershell
.\Copy-GraphMailboxItems.ps1 `
  -SourceUserPrincipalName 'source@contoso.com' `
  -TargetUserPrincipalName 'target@contoso.com' `
  -SourceFolderPath 'Inbox\Projects\FY26' `
  -TargetFolderPath 'Inbox' `
  -CopyEmptyFolders
```

To copy only specific subfolders:

```powershell
.\Copy-GraphMailboxItems.ps1 `
  -SourceUserPrincipalName 'source@contoso.com' `
  -TargetUserPrincipalName 'target@contoso.com' `
  -SourceFolderPath 'Inbox' `
  -TargetFolderPath 'Migrated' `
  -IncludeFolderPath 'Projects\FY26','Projects\FY27'
```

To copy everything except selected subfolders:

```powershell
.\Copy-GraphMailboxItems.ps1 `
  -SourceUserPrincipalName 'source@contoso.com' `
  -TargetUserPrincipalName 'target@contoso.com' `
  -SourceFolderPath 'Inbox' `
  -TargetFolderPath 'Migrated' `
  -ExcludeFolderPath 'Newsletters','LowPriority'
```

To copy only items created on or after January 1, 2025:

```powershell
.\Copy-GraphMailboxItems.ps1 `
  -SourceUserPrincipalName 'source@contoso.com' `
  -TargetUserPrincipalName 'target@contoso.com' `
  -SourceFolderPath 'Inbox' `
  -TargetFolderPath 'Migrated' `
  -Oldest '2025-01-01'
```

To copy only items created on or before January 31, 2025:

```powershell
.\Copy-GraphMailboxItems.ps1 `
  -SourceUserPrincipalName 'source@contoso.com' `
  -TargetUserPrincipalName 'target@contoso.com' `
  -SourceFolderPath 'Inbox' `
  -TargetFolderPath 'Migrated' `
  -Newest '2025-01-31'
```

To copy only items created within a date range:

```powershell
.\Copy-GraphMailboxItems.ps1 `
  -SourceUserPrincipalName 'source@contoso.com' `
  -TargetUserPrincipalName 'target@contoso.com' `
  -SourceFolderPath 'Inbox' `
  -TargetFolderPath 'Migrated' `
  -Oldest '2025-01-01' `
  -Newest '2025-01-31'
```

To merge the source `Calendar` directly into the target mailbox's main Calendar:

```powershell
.\Copy-GraphMailboxItems.ps1 `
  -SourceUserPrincipalName 'source@contoso.com' `
  -TargetUserPrincipalName 'target@contoso.com' `
  -SourceFolderPath 'Calendar' `
  -ImportDirectlyIntoTargetFolder
```

To copy the source `Calendar` as a named sub-calendar under the target mailbox's main Calendar:

```powershell
.\Copy-GraphMailboxItems.ps1 `
  -SourceUserPrincipalName 'source@contoso.com' `
  -TargetUserPrincipalName 'target@contoso.com' `
  -SourceFolderPath 'Calendar' `
  -TargetFolderPath 'Migrated Calendar'
```

The same special handling also applies to `Tasks`, `Notes`, and `Contacts`:

- With `-ImportDirectlyIntoTargetFolder`, items are merged into the target mailbox's main `Tasks`, `Notes`, or `Contacts` folder.
- Without `-ImportDirectlyIntoTargetFolder`, `TargetFolderPath` is treated as the name of a single subfolder to create under the target mailbox's main `Tasks`, `Notes`, or `Contacts` folder.

## Notes

- `SourceFolderPath` and `TargetFolderPath` are matched by folder display name.
- Command-line parameters override `.env` values when both are present.
- The standard PowerShell `-Verbose` switch is supported when you want additional execution detail.
- The script shows a confirmation summary before proceeding unless `-Force` is used.
- `IncludeFolderPath` and `ExcludeFolderPath` are mutually exclusive.
- Included or excluded folder paths can be relative to `SourceFolderPath` or full mailbox-style paths.
- Empty folders are skipped by default unless `-CopyEmptyFolders` is used.
- A preflight check stops the run early if the target mailbox has duplicate sibling folders with the same name in a location the copy would need to use.
- `PreflightOnly` runs mailbox resolution, folder discovery, filter planning, and target ambiguity checks, then exits before any copy occurs.
- `OverlayMode` currently supports only `SourceFolderPath '\'` and merges the source mailbox root into the target mailbox root or into the folder specified by `TargetFolderPath`.
- `OverlayMode` cannot be combined with `ImportDirectlyIntoTargetFolder`.
- `Journal`, `Conversation History`, and `RSS Subscriptions` are always excluded from copy, even if explicitly included.
- `Journal` exclusion is matched by folder class when available, so it is more resilient across mailbox languages.
- `Conversation History` and `RSS Subscriptions` exclusions rely on folder name matching in this Graph mailbox API, so non-English mailboxes might still require an explicit `-ExcludeFolderPath` if Microsoft localizes those folder names differently.
- If `SourceFolderPath` points to a Calendar, Tasks, Notes, or Contacts folder, `ImportDirectlyIntoTargetFolder` means import directly into the target mailbox's matching main folder.
- If `SourceFolderPath` points to a Calendar, Tasks, Notes, or Contacts folder and `ImportDirectlyIntoTargetFolder` is not used, `TargetFolderPath` is treated as the name of a subfolder to create under the target mailbox's matching main folder.
- Calendar, Tasks, Notes, and Contacts special copies only support a single target subfolder name in `TargetFolderPath`, not a nested folder path.
- `Oldest` and `Newest` filter on the mailbox item's `createdDateTime`.
- A date-only `Oldest` value is treated as inclusive from the start of that date.
- A date-only `Newest` value is treated as inclusive through the end of that date.
- `Oldest` and `Newest` can also be full timestamps, for example `2025-01-01T12:30:00Z`.
- The Graph mailbox import/export APIs are currently `beta` preview APIs, so Microsoft can change them.
- Export is limited to 20 items per request, which is why the script batches item IDs.
