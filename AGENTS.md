# AGENTS.md

## Purpose

This repository contains a PowerShell utility, `Copy-GraphMailboxItems.ps1`, for copying Exchange Online mailbox content by using Microsoft Graph mailbox import/export APIs.

The script supports:

- same-tenant and cross-tenant copies
- `.env`-driven defaults with command-line override
- preflight validation
- confirmation with `-Force` bypass
- overlay mailbox copies
- special handling for Calendar, Contacts, Tasks, and Notes
- include/exclude folder filtering
- date-range filtering

## Primary Files

- `Copy-GraphMailboxItems.ps1`
  Main implementation. Most changes belong here.
- `README.md`
  User-facing behavior, examples, and parameter documentation.
- `SETUP.md`
  Prerequisites, permissions, tenant/app setup, and certificate requirements.
- `.env.example`
  Template for local runtime defaults.
- `.env`
  Local machine configuration. Do not commit secrets from this file.

## Editing Expectations

- Keep the script PowerShell-first. Do not replace major logic with external dependencies unless clearly necessary.
- Preserve existing operator ergonomics:
  - `-WhatIf`
  - `-Verbose`
  - `-PreflightOnly`
  - confirmation summary
  - `-Force`
- Treat the confirmation, preflight, and `.env` behaviors as core features, not optional extras.
- Prefer additive, low-risk changes over broad rewrites.
- Keep comments sparse and useful.

## Configuration Rules

- Command-line parameters must continue to override `.env` values.
- `.env.example` should stay safe to commit and should not contain live secrets.
- `.env` should remain git-ignored.
- If new settings are added to the script, update:
  - `.env.example`
  - `README.md`
  - confirmation summary if the setting materially affects behavior

## Behavioral Invariants

When modifying the script, preserve these behaviors unless explicitly changing them on purpose:

- `Journal`, `Conversation History`, and `RSS Subscriptions` are always excluded.
- `IncludeFolderPath` and `ExcludeFolderPath` are mutually exclusive.
- Empty folders are skipped by default unless `-CopyEmptyFolders` is used.
- Special roots:
  - `Calendar`
  - `Contacts`
  - `Tasks`
  - `Notes`
  must retain their current direct-import and subfolder behaviors.
- Overlay mode must not create a wrapper folder for source root copies.
- Preflight must fail early on ambiguous duplicate target folders.
- Confirmation must summarize the operation before real execution unless bypassed with `-Force`.

## Validation Checklist

After changing behavior, do as many of these as apply:

1. Run a PowerShell parse check.
2. Run a `-PreflightOnly` validation for the affected scenario.
3. Use `-WhatIf` for copy-path changes when possible.
4. Update docs if parameters, defaults, or examples changed.

Simple parse check:

```powershell
$null = [System.Management.Automation.Language.Parser]::ParseFile(
  'C:\_repos\GraphMailboxSync\Copy-GraphMailboxItems.ps1',
  [ref]$null,
  [ref]$null
)
```

## Change Guidance

- For auth or tenant changes:
  update both `README.md` and `SETUP.md`.
- For parameter changes:
  update the param block, `.env` loading, `.env.example`, and docs together.
- For target mapping or folder traversal changes:
  review overlay mode, special-folder routing, and preflight ambiguity checks together.
- For exclusion logic changes:
  document whether the match is type-based or display-name-based.

## Git Notes

- Commit only intended repo files.
- Do not commit `.env`.
- Keep commit messages short and descriptive.

