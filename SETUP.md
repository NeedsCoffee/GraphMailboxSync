# Setup Guide

This document describes what has to be in place before `Copy-GraphMailboxItems.ps1` can run successfully.

The script can be used in:

- same-tenant copies
- cross-tenant copies

For cross-tenant copies, treat the source and target as separate environments. In practice, that means you should expect to configure an application in each tenant you plan to work with.

## What You Need

At a minimum, you need:

- an Exchange Online mailbox in the source tenant
- an Exchange Online mailbox in the target tenant
- a Microsoft Entra application with certificate-based application authentication in the source tenant
- a Microsoft Entra application with certificate-based application authentication in the target tenant
- Microsoft Graph admin consent for the required application permissions in each tenant
- access for each application to the specific mailboxes it needs to touch if Exchange application RBAC or application access policies are in use
- the certificate installed on the machine that runs the script, in `Cert:\CurrentUser\My`

## Required Graph Permissions

Each application used by this script should have these Microsoft Graph application permissions:

- `MailboxFolder.ReadWrite.All`
- `MailboxItem.ImportExport.All`
- `User.Read.All`

These are the minimum permissions currently documented by this repo for the script’s folder discovery, export, import, and mailbox resolution flow.

## Tenant Model

### Same-Tenant Copy

For a same-tenant copy, one app registration is usually enough.

That single app needs:

- the required Graph application permissions
- admin consent in that tenant
- certificate-based auth configured
- access to both the source and target mailboxes

### Cross-Tenant Copy

For a cross-tenant copy, you should plan on one app per tenant:

- one source-side app in the source tenant
- one target-side app in the target tenant

The source-side app is used for:

- resolving the source mailbox
- loading the source folder tree
- reading source items
- exporting source items

The target-side app is used for:

- resolving the target mailbox
- loading the target folder tree
- creating target folders as needed
- creating the target import session
- importing items into the target mailbox

The script supports different client IDs, tenant IDs, and certificate thumbprints for each side, so the two apps can be completely separate.

## App Registration Checklist

Repeat these steps in each tenant you want to use with the script.

1. Create a new Microsoft Entra app registration.
2. Configure it for application authentication.
3. Upload or assign the certificate that the script will use.
4. Add the Graph application permissions listed above.
5. Grant admin consent for those permissions in that tenant.
6. Record the tenant ID, client ID, and certificate thumbprint for later use.

## Certificate Requirements

The script authenticates by looking up a certificate thumbprint in:

- `Cert:\CurrentUser\My`

That means the machine and user account running the script must have access to the certificate in that certificate store.

Before running the script, verify:

- the certificate exists in `Cert:\CurrentUser\My`
- the thumbprint in your command line or `.env` file matches exactly
- the certificate private key is available to the user running PowerShell

## Exchange Mailbox Scoping

If your organization restricts application access to Exchange mailboxes, make sure each app is allowed to the mailboxes it needs.

Examples:

- the source-side app must be allowed to the source mailbox
- the target-side app must be allowed to the target mailbox

If one app is used for both sides in a same-tenant copy, it must be allowed to both mailboxes.

## Values You Will Need At Runtime

For each run, gather:

- source user principal name
- target user principal name
- source folder path
- optional target folder path
- source tenant ID
- source client ID
- source certificate thumbprint
- target tenant ID
- target client ID
- target certificate thumbprint

For same-tenant runs, you can instead use the legacy shared settings:

- `TenantId`
- `ClientId`
- `CertificateThumbprint`

## Configuration Options

You can provide settings in either of these ways:

- `.env`
- command-line parameters

Command-line parameters override `.env` values.

If you want to disable `.env` loading entirely, run the script with:

```powershell
-EnvFile ''
```

## Example `.env` Shape For Cross-Tenant Use

```dotenv
SOURCE_TENANT_ID=source-tenant-guid
SOURCE_CLIENT_ID=source-app-guid
SOURCE_CERTIFICATE_THUMBPRINT=source-cert-thumbprint
TARGET_TENANT_ID=target-tenant-guid
TARGET_CLIENT_ID=target-app-guid
TARGET_CERTIFICATE_THUMBPRINT=target-cert-thumbprint
SOURCE_USER_PRINCIPAL_NAME=user@sourcetenant.com
TARGET_USER_PRINCIPAL_NAME=user@targettenant.com
SOURCE_FOLDER_PATH=Inbox\Projects
```

## Example Command-Line Shape For Cross-Tenant Use

```powershell
.\Copy-GraphMailboxItems.ps1 `
  -EnvFile '' `
  -SourceUserPrincipalName 'user@sourcetenant.com' `
  -TargetUserPrincipalName 'user@targettenant.com' `
  -SourceFolderPath 'Inbox\Projects' `
  -SourceTenantId '11111111-1111-1111-1111-111111111111' `
  -SourceClientId '22222222-2222-2222-2222-222222222222' `
  -SourceCertificateThumbprint 'SOURCECERTTHUMBPRINT' `
  -TargetTenantId '33333333-3333-3333-3333-333333333333' `
  -TargetClientId '44444444-4444-4444-4444-444444444444' `
  -TargetCertificateThumbprint 'TARGETCERTTHUMBPRINT' `
  -PreflightOnly
```

## Recommended First Test

Before a real copy, run a preflight:

```powershell
.\Copy-GraphMailboxItems.ps1 -PreflightOnly
```

That validates:

- authentication
- mailbox resolution
- source folder discovery
- target folder discovery
- target ambiguity checks

If preflight passes, then run the real copy.
