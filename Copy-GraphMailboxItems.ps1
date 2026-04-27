[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [string]$SourceUserPrincipalName,

    [string]$TargetUserPrincipalName,

    [string]$SourceFolderPath,

    [string]$TargetFolderPath = '',

    [string]$TenantId,

    [string]$ClientId,

    [string]$CertificateThumbprint,

    [switch]$ImportDirectlyIntoTargetFolder,

    [switch]$OverlayMode,

    [switch]$CopyEmptyFolders,

    [string[]]$IncludeFolderPath,

    [string[]]$ExcludeFolderPath,

    [string]$Oldest,

    [string]$Newest,

    [switch]$PreflightOnly,

    [switch]$Force,

    [string]$EnvFile = '.env',

    [int]$ExportBatchSize = 20
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$script:RootParentKey = '__root__'
$script:AlwaysExcludedFolderDisplayNames = @(
    'Conversation History',
    'Journal',
    'RSS Subscriptions'
)
$script:AlwaysExcludedFolderTypes = @(
    'IPF.Journal'
)

function ConvertFrom-DotEnvValue {
    param(
        [AllowNull()]
        [string]$Value
    )

    if ($null -eq $Value) {
        return ''
    }

    $trimmedValue = $Value.Trim()
    if ($trimmedValue.Length -ge 2) {
        $firstCharacter = $trimmedValue.Substring(0, 1)
        $lastCharacter = $trimmedValue.Substring($trimmedValue.Length - 1, 1)
        if (($firstCharacter -eq '"' -and $lastCharacter -eq '"') -or ($firstCharacter -eq "'" -and $lastCharacter -eq "'")) {
            return $trimmedValue.Substring(1, $trimmedValue.Length - 2)
        }
    }

    return $trimmedValue
}

function Read-DotEnvFile {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $resolvedPath = Resolve-Path -LiteralPath $Path -ErrorAction Stop
    $values = @{}

    foreach ($line in (Get-Content -LiteralPath $resolvedPath)) {
        $trimmedLine = $line.Trim()
        if ([string]::IsNullOrWhiteSpace($trimmedLine) -or $trimmedLine.StartsWith('#')) {
            continue
        }

        $separatorIndex = $trimmedLine.IndexOf('=')
        if ($separatorIndex -lt 1) {
            continue
        }

        $key = $trimmedLine.Substring(0, $separatorIndex).Trim()
        $value = $trimmedLine.Substring($separatorIndex + 1)
        if ($key.StartsWith('export ')) {
            $key = $key.Substring(7).Trim()
        }

        if (-not [string]::IsNullOrWhiteSpace($key)) {
            $values[$key] = ConvertFrom-DotEnvValue -Value $value
        }
    }

    return $values
}

function ConvertTo-DotEnvBoolean {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Value,

        [Parameter(Mandatory = $true)]
        [string]$Key
    )

    switch -Regex ($Value.Trim()) {
        '^(1|true|yes|y|on)$' { return $true }
        '^(0|false|no|n|off)$' { return $false }
        default { throw "Unable to parse boolean value '$Value' for $Key in the .env configuration." }
    }
}

function ConvertTo-DotEnvArray {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Value
    )

    @($Value -split ';' | ForEach-Object { $_.Trim() } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
}

function Set-SettingSource {
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$Map,

        [Parameter(Mandatory = $true)]
        [string]$Name,

        [Parameter(Mandatory = $true)]
        [string]$Source
    )

    $Map[$Name] = $Source
}

function Get-SettingSourceLabel {
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$Map,

        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    if ($Map.ContainsKey($Name)) {
        return [string]$Map[$Name]
    }

    return 'default'
}

function Get-DotEnvSettingValue {
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$DotEnvValues,

        [Parameter(Mandatory = $true)]
        [string[]]$Keys
    )

    foreach ($key in $Keys) {
        if ($DotEnvValues.ContainsKey($key)) {
            return [pscustomobject]@{
                Found = $true
                Key   = $key
                Value = $DotEnvValues[$key]
            }
        }
    }

    return [pscustomobject]@{
        Found = $false
        Key   = $null
        Value = $null
    }
}

function Set-ValueFromDotEnv {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ParameterName,

        [Parameter(Mandatory = $true)]
        [string[]]$DotEnvKeys,

        [Parameter(Mandatory = $true)]
        [scriptblock]$Assignment
    )

    if ($script:CommandLineParameterNames -contains $ParameterName) {
        return
    }

    $resolvedSetting = Get-DotEnvSettingValue -DotEnvValues $dotEnvValues -Keys $DotEnvKeys
    if (-not $resolvedSetting.Found) {
        return
    }

    & $Assignment $resolvedSetting.Value $resolvedSetting.Key
    Set-SettingSource -Map $settingSources -Name $ParameterName -Source ".env ($($resolvedSetting.Key))"
}

$dotEnvValues = @{}
$settingSources = @{}
$script:CommandLineParameterNames = @($PSBoundParameters.Keys)
$isDefaultEnvFile = [string]::Equals($EnvFile, '.env', [System.StringComparison]::OrdinalIgnoreCase)
if (-not [string]::IsNullOrWhiteSpace($EnvFile)) {
    if (Test-Path -LiteralPath $EnvFile) {
        $dotEnvValues = Read-DotEnvFile -Path $EnvFile
        Write-Verbose "Loaded configuration defaults from '$EnvFile'."
    }
    elseif (-not $isDefaultEnvFile) {
        throw "The specified EnvFile '$EnvFile' could not be found."
    }
}

foreach ($boundParameterName in $PSBoundParameters.Keys) {
    Set-SettingSource -Map $settingSources -Name $boundParameterName -Source 'command line'
}

Set-ValueFromDotEnv -ParameterName 'SourceUserPrincipalName' -DotEnvKeys @('SOURCE_USER_PRINCIPAL_NAME') -Assignment {
    param($value, $key)
    $SourceUserPrincipalName = $value
}

Set-ValueFromDotEnv -ParameterName 'TargetUserPrincipalName' -DotEnvKeys @('TARGET_USER_PRINCIPAL_NAME') -Assignment {
    param($value, $key)
    $TargetUserPrincipalName = $value
}

Set-ValueFromDotEnv -ParameterName 'SourceFolderPath' -DotEnvKeys @('SOURCE_FOLDER_PATH') -Assignment {
    param($value, $key)
    $SourceFolderPath = $value
}

Set-ValueFromDotEnv -ParameterName 'TargetFolderPath' -DotEnvKeys @('TARGET_FOLDER_PATH') -Assignment {
    param($value, $key)
    $TargetFolderPath = $value
}

Set-ValueFromDotEnv -ParameterName 'TenantId' -DotEnvKeys @('TENANT_ID') -Assignment {
    param($value, $key)
    $TenantId = $value
}

Set-ValueFromDotEnv -ParameterName 'ClientId' -DotEnvKeys @('CLIENT_ID') -Assignment {
    param($value, $key)
    $ClientId = $value
}

Set-ValueFromDotEnv -ParameterName 'CertificateThumbprint' -DotEnvKeys @('CERTIFICATE_THUMBPRINT') -Assignment {
    param($value, $key)
    $CertificateThumbprint = $value
}

Set-ValueFromDotEnv -ParameterName 'ImportDirectlyIntoTargetFolder' -DotEnvKeys @('IMPORT_DIRECTLY_INTO_TARGET_FOLDER') -Assignment {
    param($value, $key)
    $ImportDirectlyIntoTargetFolder = ConvertTo-DotEnvBoolean -Value $value -Key $key
}

Set-ValueFromDotEnv -ParameterName 'OverlayMode' -DotEnvKeys @('OVERLAY_MODE') -Assignment {
    param($value, $key)
    $OverlayMode = ConvertTo-DotEnvBoolean -Value $value -Key $key
}

Set-ValueFromDotEnv -ParameterName 'CopyEmptyFolders' -DotEnvKeys @('COPY_EMPTY_FOLDERS') -Assignment {
    param($value, $key)
    $CopyEmptyFolders = ConvertTo-DotEnvBoolean -Value $value -Key $key
}

Set-ValueFromDotEnv -ParameterName 'IncludeFolderPath' -DotEnvKeys @('INCLUDE_FOLDER_PATH') -Assignment {
    param($value, $key)
    $IncludeFolderPath = ConvertTo-DotEnvArray -Value $value
}

Set-ValueFromDotEnv -ParameterName 'ExcludeFolderPath' -DotEnvKeys @('EXCLUDE_FOLDER_PATH') -Assignment {
    param($value, $key)
    $ExcludeFolderPath = ConvertTo-DotEnvArray -Value $value
}

Set-ValueFromDotEnv -ParameterName 'Oldest' -DotEnvKeys @('OLDEST') -Assignment {
    param($value, $key)
    $Oldest = $value
}

Set-ValueFromDotEnv -ParameterName 'Newest' -DotEnvKeys @('NEWEST') -Assignment {
    param($value, $key)
    $Newest = $value
}

Set-ValueFromDotEnv -ParameterName 'PreflightOnly' -DotEnvKeys @('PREFLIGHT_ONLY') -Assignment {
    param($value, $key)
    $PreflightOnly = ConvertTo-DotEnvBoolean -Value $value -Key $key
}

Set-ValueFromDotEnv -ParameterName 'Force' -DotEnvKeys @('FORCE') -Assignment {
    param($value, $key)
    $Force = ConvertTo-DotEnvBoolean -Value $value -Key $key
}

Set-ValueFromDotEnv -ParameterName 'ExportBatchSize' -DotEnvKeys @('EXPORT_BATCH_SIZE') -Assignment {
    param($value, $key)
    $parsedExportBatchSize = 0
    if (-not [int]::TryParse($value, [ref]$parsedExportBatchSize)) {
        throw "Unable to parse integer value '$value' for $key in the .env configuration."
    }

    $ExportBatchSize = $parsedExportBatchSize
}

if ([string]::IsNullOrWhiteSpace($SourceUserPrincipalName)) {
    throw 'SourceUserPrincipalName must be provided either on the command line or in the .env configuration.'
}

if ([string]::IsNullOrWhiteSpace($TargetUserPrincipalName)) {
    throw 'TargetUserPrincipalName must be provided either on the command line or in the .env configuration.'
}

if ([string]::IsNullOrWhiteSpace($SourceFolderPath)) {
    throw 'SourceFolderPath must be provided either on the command line or in the .env configuration.'
}

if ([string]::IsNullOrWhiteSpace($TenantId)) {
    throw 'TenantId must be provided either on the command line or in the .env configuration.'
}

if ([string]::IsNullOrWhiteSpace($ClientId)) {
    throw 'ClientId must be provided either on the command line or in the .env configuration.'
}

if ([string]::IsNullOrWhiteSpace($CertificateThumbprint)) {
    throw 'CertificateThumbprint must be provided either on the command line or in the .env configuration.'
}

if ($ExportBatchSize -lt 1 -or $ExportBatchSize -gt 20) {
    throw 'ExportBatchSize must be between 1 and 20 because Graph exportItems accepts at most 20 item IDs per request.'
}

if ($IncludeFolderPath -and $ExcludeFolderPath) {
    throw 'IncludeFolderPath and ExcludeFolderPath are mutually exclusive. Specify only one of them.'
}

if ($OverlayMode -and $ImportDirectlyIntoTargetFolder) {
    throw 'OverlayMode and ImportDirectlyIntoTargetFolder cannot be used together. OverlayMode already merges directly into the target structure.'
}

function Resolve-DateFilterBoundary {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Value,

        [Parameter(Mandatory = $true)]
        [ValidateSet('Oldest', 'Newest')]
        [string]$Boundary
    )

    $trimmedValue = $Value.Trim()
    if ([string]::IsNullOrWhiteSpace($trimmedValue)) {
        throw "$Boundary cannot be empty when specified."
    }

    if ($trimmedValue -match '^\d{4}-\d{2}-\d{2}$') {
        $dateOnly = [datetime]::ParseExact($trimmedValue, 'yyyy-MM-dd', [System.Globalization.CultureInfo]::InvariantCulture)
        $comparisonOperator = if ($Boundary -eq 'Oldest') { 'ge' } else { 'lt' }
        $comparisonValue = if ($Boundary -eq 'Oldest') {
            $dateOnly.ToString('yyyy-MM-dd', [System.Globalization.CultureInfo]::InvariantCulture)
        }
        else {
            $dateOnly.AddDays(1).ToString('yyyy-MM-dd', [System.Globalization.CultureInfo]::InvariantCulture)
        }

        return [pscustomobject]@{
            RawValue            = $trimmedValue
            IsDateOnly          = $true
            InclusiveLowerBound = if ($Boundary -eq 'Oldest') { $dateOnly.Date } else { $dateOnly.Date.AddDays(1).AddTicks(-1) }
            FilterClause        = "createdDateTime $comparisonOperator $comparisonValue"
        }
    }

    try {
        $parsedDate = [datetimeoffset]::Parse($trimmedValue, [System.Globalization.CultureInfo]::InvariantCulture)
    }
    catch {
        throw "Unable to parse $Boundary value '$trimmedValue' as a date or timestamp."
    }

    $utcDate = $parsedDate.ToUniversalTime()
    $formattedUtcDate = $utcDate.ToString('yyyy-MM-ddTHH:mm:ssZ', [System.Globalization.CultureInfo]::InvariantCulture)
    $comparisonOperator = if ($Boundary -eq 'Oldest') { 'ge' } else { 'le' }

    return [pscustomobject]@{
        RawValue            = $trimmedValue
        IsDateOnly          = $false
        InclusiveLowerBound = $utcDate.UtcDateTime
        FilterClause        = "createdDateTime $comparisonOperator $formattedUtcDate"
    }
}

function New-MailboxItemDateFilter {
    param(
        [AllowNull()]
        [string]$Oldest,

        [AllowNull()]
        [string]$Newest
    )

    if ([string]::IsNullOrWhiteSpace($Oldest) -and [string]::IsNullOrWhiteSpace($Newest)) {
        return $null
    }

    $oldestBoundary = if ([string]::IsNullOrWhiteSpace($Oldest)) { $null } else { Resolve-DateFilterBoundary -Value $Oldest -Boundary Oldest }
    $newestBoundary = if ([string]::IsNullOrWhiteSpace($Newest)) { $null } else { Resolve-DateFilterBoundary -Value $Newest -Boundary Newest }

    if ($oldestBoundary -and $newestBoundary) {
        if ($oldestBoundary.InclusiveLowerBound -gt $newestBoundary.InclusiveLowerBound) {
            throw "Oldest value '$Oldest' must be earlier than or equal to Newest value '$Newest'."
        }
    }

    $filterClauses = @(
        if ($oldestBoundary) { $oldestBoundary.FilterClause }
        if ($newestBoundary) { $newestBoundary.FilterClause }
    )

    [pscustomobject]@{
        Oldest     = $oldestBoundary
        Newest     = $newestBoundary
        FilterText = ($filterClauses -join ' and ')
    }
}

function Get-ExecutionModeSummary {
    param(
        [Parameter(Mandatory = $true)]
        [switch]$OverlayMode,

        [Parameter(Mandatory = $true)]
        [switch]$ImportDirectlyIntoTargetFolder,

        [Parameter(Mandatory = $true)]
        [switch]$PreflightOnly
    )

    if ($PreflightOnly) {
        return 'PreflightOnly'
    }

    if ($OverlayMode) {
        return 'Overlay'
    }

    if ($ImportDirectlyIntoTargetFolder) {
        return 'DirectImport'
    }

    return 'StructuredCopy'
}

function Confirm-PlannedOperation {
    param(
        [Parameter(Mandatory = $true)]
        [string]$SourceUserPrincipalName,

        [Parameter(Mandatory = $true)]
        [string]$TargetUserPrincipalName,

        [Parameter(Mandatory = $true)]
        [string]$SourceFolderPath,

        [Parameter(Mandatory = $true)]
        [string]$ResolvedTargetDescription,

        [Parameter(Mandatory = $true)]
        [string]$Mode,

        [Parameter(Mandatory = $true)]
        [switch]$OverlayMode,

        [Parameter(Mandatory = $true)]
        [switch]$ImportDirectlyIntoTargetFolder,

        [AllowNull()]
        [object]$DateFilter,

        [AllowNull()]
        [string[]]$IncludeFolderPath,

        [AllowNull()]
        [string[]]$ExcludeFolderPath,

        [Parameter(Mandatory = $true)]
        [hashtable]$SettingSources,

        [Parameter(Mandatory = $true)]
        [int]$EstimatedSelectedFolderCount,

        [Parameter(Mandatory = $true)]
        [int]$EstimatedTraversedFolderCount,

        [Parameter(Mandatory = $true)]
        [int]$EstimatedItemCount,

        [Parameter(Mandatory = $true)]
        [switch]$CopyEmptyFolders,

        [Parameter(Mandatory = $true)]
        [switch]$PreflightOnly,

        [Parameter(Mandatory = $true)]
        [switch]$WhatIfMode,

        [Parameter(Mandatory = $true)]
        [switch]$Force
    )

    $sourceMailboxSource = Get-SettingSourceLabel -Map $SettingSources -Name 'SourceUserPrincipalName'
    $targetMailboxSource = Get-SettingSourceLabel -Map $SettingSources -Name 'TargetUserPrincipalName'
    $sourcePathSource = Get-SettingSourceLabel -Map $SettingSources -Name 'SourceFolderPath'
    $targetPathSource = Get-SettingSourceLabel -Map $SettingSources -Name 'TargetFolderPath'
    $copyEmptySource = Get-SettingSourceLabel -Map $SettingSources -Name 'CopyEmptyFolders'
    $modeSource = if ($PreflightOnly -and $SettingSources.ContainsKey('PreflightOnly')) {
        Get-SettingSourceLabel -Map $SettingSources -Name 'PreflightOnly'
    }
    elseif ($ImportDirectlyIntoTargetFolder -and $SettingSources.ContainsKey('ImportDirectlyIntoTargetFolder')) {
        Get-SettingSourceLabel -Map $SettingSources -Name 'ImportDirectlyIntoTargetFolder'
    }
    elseif ($OverlayMode -and $SettingSources.ContainsKey('OverlayMode')) {
        Get-SettingSourceLabel -Map $SettingSources -Name 'OverlayMode'
    }
    else {
        'default'
    }

    Write-Host 'Planned operation:'
    Write-Host ("  Source mailbox : {0} [{1}]" -f $SourceUserPrincipalName, $sourceMailboxSource)
    Write-Host ("  Target mailbox : {0} [{1}]" -f $TargetUserPrincipalName, $targetMailboxSource)
    Write-Host ("  Source path    : {0} [{1}]" -f $SourceFolderPath, $sourcePathSource)
    Write-Host ("  Target         : {0} [{1}]" -f $ResolvedTargetDescription, $targetPathSource)
    Write-Host ("  Mode           : {0} [{1}]" -f $Mode, $modeSource)
    Write-Host ("  Copy empty     : {0} [{1}]" -f ($(if ($CopyEmptyFolders) { 'Yes' } else { 'No' })), $copyEmptySource)
    Write-Host ("  Est. folders   : {0} selected, {1} traversed" -f $EstimatedSelectedFolderCount, $EstimatedTraversedFolderCount)
    Write-Host ("  Est. items     : {0}" -f $EstimatedItemCount)

    if ($DateFilter) {
        $dateFilterSource = if ($SettingSources.ContainsKey('Oldest') -or $SettingSources.ContainsKey('Newest')) {
            if ($SettingSources.ContainsKey('Oldest')) {
                Get-SettingSourceLabel -Map $SettingSources -Name 'Oldest'
            }
            else {
                Get-SettingSourceLabel -Map $SettingSources -Name 'Newest'
            }
        }
        else {
            'default'
        }

        Write-Host ("  Date filter    : {0} [{1}]" -f $DateFilter.FilterText, $dateFilterSource)
    }
    else {
        Write-Host '  Date filter    : None [default]'
    }

    if ($IncludeFolderPath) {
        Write-Host ("  Include paths  : {0} [{1}]" -f (($IncludeFolderPath | ForEach-Object { Normalize-FolderPath -Path $_ }) -join ', '), (Get-SettingSourceLabel -Map $SettingSources -Name 'IncludeFolderPath'))
    }
    elseif ($ExcludeFolderPath) {
        Write-Host ("  Exclude paths  : {0} [{1}]" -f (($ExcludeFolderPath | ForEach-Object { Normalize-FolderPath -Path $_ }) -join ', '), (Get-SettingSourceLabel -Map $SettingSources -Name 'ExcludeFolderPath'))
    }
    else {
        Write-Host '  Folder filter  : None [default]'
    }

    Write-Host ("  WhatIf         : {0} [{1}]" -f ($(if ($WhatIfMode) { 'Yes' } else { 'No' })), 'command line/session')
    Write-Host ("  Preflight only : {0} [{1}]" -f ($(if ($PreflightOnly) { 'Yes' } else { 'No' })), (Get-SettingSourceLabel -Map $SettingSources -Name 'PreflightOnly'))

    if ($Force) {
        Write-Verbose 'Skipping confirmation because Force was specified.'
        return
    }

    $caption = 'Confirm mailbox copy'
    $message = 'Proceed with this mailbox operation?'
    if (-not $PSCmdlet.ShouldContinue($message, $caption)) {
        throw 'Operation cancelled by user.'
    }
}

function ConvertTo-Base64Url {
    param(
        [Parameter(Mandatory = $true)]
        [byte[]]$Bytes
    )

    [Convert]::ToBase64String($Bytes).TrimEnd('=').Replace('+', '-').Replace('/', '_')
}

function Get-ClientCertificate {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Thumbprint
    )

    $normalizedThumbprint = $Thumbprint.Replace(' ', '').ToUpperInvariant()
    $certificate = Get-ChildItem -Path Cert:\CurrentUser\My |
        Where-Object Thumbprint -eq $normalizedThumbprint |
        Select-Object -First 1

    if (-not $certificate) {
        throw "Certificate '$normalizedThumbprint' was not found in Cert:\CurrentUser\My."
    }

    if (-not $certificate.HasPrivateKey) {
        throw "Certificate '$normalizedThumbprint' does not have an accessible private key."
    }

    return $certificate
}

function New-ClientAssertionJwt {
    param(
        [Parameter(Mandatory = $true)]
        [string]$TenantId,

        [Parameter(Mandatory = $true)]
        [string]$ClientId,

        [Parameter(Mandatory = $true)]
        [System.Security.Cryptography.X509Certificates.X509Certificate2]$Certificate
    )

    $tokenEndpoint = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
    $now = [DateTimeOffset]::UtcNow

    $headerJson = @{
        alg = 'RS256'
        typ = 'JWT'
        x5t = ConvertTo-Base64Url -Bytes $Certificate.GetCertHash()
    } | ConvertTo-Json -Compress

    $payloadJson = @{
        aud = $tokenEndpoint
        iss = $ClientId
        sub = $ClientId
        jti = [guid]::NewGuid().Guid
        nbf = $now.ToUnixTimeSeconds()
        exp = $now.AddMinutes(10).ToUnixTimeSeconds()
    } | ConvertTo-Json -Compress

    $headerEncoded = ConvertTo-Base64Url -Bytes ([Text.Encoding]::UTF8.GetBytes($headerJson))
    $payloadEncoded = ConvertTo-Base64Url -Bytes ([Text.Encoding]::UTF8.GetBytes($payloadJson))
    $unsignedToken = "$headerEncoded.$payloadEncoded"

    $rsa = [System.Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPrivateKey($Certificate)
    if (-not $rsa) {
        throw "Unable to access the private key for certificate '$($Certificate.Thumbprint)'."
    }

    try {
        $signatureBytes = $rsa.SignData(
            [Text.Encoding]::UTF8.GetBytes($unsignedToken),
            [Security.Cryptography.HashAlgorithmName]::SHA256,
            [Security.Cryptography.RSASignaturePadding]::Pkcs1
        )
    }
    finally {
        $rsa.Dispose()
    }

    $signatureEncoded = ConvertTo-Base64Url -Bytes $signatureBytes
    return "$unsignedToken.$signatureEncoded"
}

function Get-GraphAccessToken {
    param(
        [Parameter(Mandatory = $true)]
        [string]$TenantId,

        [Parameter(Mandatory = $true)]
        [string]$ClientId,

        [Parameter(Mandatory = $true)]
        [string]$CertificateThumbprint
    )

    $certificate = Get-ClientCertificate -Thumbprint $CertificateThumbprint
    $clientAssertion = New-ClientAssertionJwt -TenantId $TenantId -ClientId $ClientId -Certificate $certificate
    $tokenEndpoint = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"

    $tokenResponse = Invoke-RestMethod -Method Post -Uri $tokenEndpoint -ContentType 'application/x-www-form-urlencoded' -Body @{
        client_id = $ClientId
        scope = 'https://graph.microsoft.com/.default'
        grant_type = 'client_credentials'
        client_assertion_type = 'urn:ietf:params:oauth:client-assertion-type:jwt-bearer'
        client_assertion = $clientAssertion
    }

    if (-not $tokenResponse.access_token) {
        throw 'Access token request did not return an access_token value.'
    }

    return $tokenResponse.access_token
}

function Invoke-GraphApiRequest {
    param(
        [Parameter(Mandatory = $true)]
        [string]$AccessToken,

        [Parameter(Mandatory = $true)]
        [string]$Uri,

        [ValidateSet('GET', 'POST', 'PATCH', 'DELETE')]
        [string]$Method = 'GET',

        $Body
    )

    $resolvedUri = if ($Uri -match '^https://') {
        $Uri
    }
    else {
        "https://graph.microsoft.com$Uri"
    }

    $invokeParams = @{
        Method  = $Method
        Uri     = $resolvedUri
        Headers = @{
            Authorization = "Bearer $AccessToken"
        }
    }

    if ($null -ne $Body) {
        $invokeParams.ContentType = 'application/json'
        $invokeParams.Body = $Body | ConvertTo-Json -Depth 10 -Compress
    }

    Invoke-RestMethod @invokeParams
}

function Invoke-GraphCollectionRequest {
    param(
        [Parameter(Mandatory = $true)]
        [string]$AccessToken,

        [Parameter(Mandatory = $true)]
        [string]$Uri
    )

    $items = [System.Collections.Generic.List[object]]::new()
    $nextUri = $Uri

    while ($nextUri) {
        $response = Invoke-GraphApiRequest -AccessToken $AccessToken -Uri $nextUri
        if ($response.value) {
            foreach ($item in $response.value) {
                $items.Add($item)
            }
        }
        elseif ($response) {
            $items.Add($response)
        }

        $nextLinkProperty = $response.PSObject.Properties['@odata.nextLink']
        $nextUri = if ($nextLinkProperty) { [string]$nextLinkProperty.Value } else { $null }
    }

    return $items
}

function Get-ExchangeSettings {
    param(
        [Parameter(Mandatory = $true)]
        [string]$AccessToken,

        [Parameter(Mandatory = $true)]
        [string]$UserPrincipalName
    )

    $encodedUser = [Uri]::EscapeDataString($UserPrincipalName)
    Invoke-GraphApiRequest -AccessToken $AccessToken -Uri "/beta/users/$encodedUser/settings/exchange"
}

function Resolve-MailboxId {
    param(
        [Parameter(Mandatory = $true)]
        [string]$AccessToken,

        [Parameter(Mandatory = $true)]
        [string]$UserPrincipalName
    )

    $exchangeSettings = Get-ExchangeSettings -AccessToken $AccessToken -UserPrincipalName $UserPrincipalName

    if (-not $exchangeSettings.primaryMailboxId) {
        throw "No primaryMailboxId was returned for '$UserPrincipalName'."
    }

    return $exchangeSettings.primaryMailboxId
}

function Get-MailboxFolderTree {
    param(
        [Parameter(Mandatory = $true)]
        [string]$AccessToken,

        [Parameter(Mandatory = $true)]
        [string]$MailboxId
    )

    $encodedMailboxId = [Uri]::EscapeDataString($MailboxId)
    $allFolders = [System.Collections.Generic.List[object]]::new()
    $queue = [System.Collections.Generic.Queue[object]]::new()
    $topLevelFolders = [System.Collections.Generic.List[object]]::new()

    foreach ($folder in (Invoke-GraphCollectionRequest -AccessToken $AccessToken -Uri "/beta/admin/exchange/mailboxes/$encodedMailboxId/folders")) {
        if ($folder.PSObject.Properties['id']) {
            $topLevelFolders.Add($folder)
            $queue.Enqueue($folder)
        }
    }

    while ($queue.Count -gt 0) {
        $folder = $queue.Dequeue()
        if (-not $folder.PSObject.Properties['id']) {
            continue
        }

        $allFolders.Add($folder)

        foreach ($child in (Invoke-GraphCollectionRequest -AccessToken $AccessToken -Uri "/beta/admin/exchange/mailboxes/$encodedMailboxId/folders/$([Uri]::EscapeDataString($folder.id))/childFolders")) {
            if ($child.PSObject.Properties['id']) {
                $queue.Enqueue($child)
            }
        }
    }

    $byId = @{}
    $childrenByParentId = @{}
    foreach ($folder in $allFolders) {
        $byId[$folder.id] = $folder
        $parentKey = if ($null -eq $folder.parentFolderId) { $script:RootParentKey } else { $folder.parentFolderId }

        if (-not $childrenByParentId.ContainsKey($parentKey)) {
            $childrenByParentId[$parentKey] = [System.Collections.Generic.List[object]]::new()
        }

        $childrenByParentId[$parentKey].Add($folder)
    }

    return @{
        AllFolders        = $allFolders
        ById              = $byId
        ChildrenByParentId = $childrenByParentId
        TopLevelFolders   = $topLevelFolders
    }
}

function Resolve-MailboxFolderPath {
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$FolderTree,

        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $segments = @($Path -split '[\\/]' | Where-Object { $_ -and $_.Trim() })
    if ($segments.Count -eq 0) {
        return [pscustomobject]@{
            id              = '$root'
            displayName     = ''
            parentFolderId  = $null
            childFolderCount = @($FolderTree.TopLevelFolders).Count
            totalItemCount  = 0
            type            = 'IPF.Root'
        }
    }

    $currentFolder = $null
    foreach ($segment in $segments) {
        $candidateChildren = @(if ($null -eq $currentFolder) {
            $FolderTree.AllFolders | Where-Object { $_.displayName -eq $segment }
        }
        else {
            @($FolderTree.ChildrenByParentId[$currentFolder.id]) | Where-Object { $_.displayName -eq $segment }
        })

        if (-not $candidateChildren -or $candidateChildren.Count -eq 0) {
            throw "Folder path '$Path' could not be resolved at segment '$segment'."
        }

        if ($candidateChildren.Count -gt 1) {
            throw "Folder path '$Path' is ambiguous at segment '$segment'."
        }

        $currentFolder = $candidateChildren[0]
    }

    return $currentFolder
}

function Get-MailboxFolderDisplayPath {
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$FolderTree,

        [Parameter(Mandatory = $true)]
        [string]$FolderId
    )

    if ($FolderId -eq '$root') {
        return '\'
    }

    $segments = [System.Collections.Generic.List[string]]::new()
    $currentId = $FolderId

    while ($currentId -and $FolderTree.ById.ContainsKey($currentId)) {
        $folder = $FolderTree.ById[$currentId]
        $segments.Insert(0, $folder.displayName)
        $currentId = $folder.parentFolderId
    }

    return ($segments -join '\')
}

function Normalize-FolderPath {
    param(
        [AllowNull()]
        [string]$Path
    )

    $segments = @($Path -split '[\\/]' | Where-Object { $_ -and $_.Trim() })
    if ($segments.Count -eq 0) {
        return ''
    }

    return ($segments -join '\')
}

function Get-RelativeFolderPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$FullPath,

        [Parameter(Mandatory = $true)]
        [string]$RootPath
    )

    $normalizedFullPath = Normalize-FolderPath -Path $FullPath
    $normalizedRootPath = Normalize-FolderPath -Path $RootPath

    if ([string]::IsNullOrWhiteSpace($normalizedRootPath)) {
        return $normalizedFullPath
    }

    if ($normalizedFullPath -eq $normalizedRootPath) {
        return ''
    }

    if ($normalizedFullPath.StartsWith("$normalizedRootPath\", [System.StringComparison]::OrdinalIgnoreCase)) {
        return $normalizedFullPath.Substring($normalizedRootPath.Length + 1)
    }

    return $normalizedFullPath
}

function Test-FolderPathMatch {
    param(
        [Parameter(Mandatory = $true)]
        [string]$FolderRelativePath,

        [Parameter(Mandatory = $true)]
        [string]$FolderFullPath,

        [Parameter(Mandatory = $true)]
        [string[]]$CandidatePaths
    )

    $normalizedRelativePath = Normalize-FolderPath -Path $FolderRelativePath
    $normalizedFullPath = Normalize-FolderPath -Path $FolderFullPath

    foreach ($candidatePath in $CandidatePaths) {
        $normalizedCandidate = Normalize-FolderPath -Path $candidatePath
        if ($normalizedCandidate -eq '') {
            return $true
        }

        foreach ($comparisonPath in @($normalizedRelativePath, $normalizedFullPath)) {
            if ([string]::IsNullOrWhiteSpace($comparisonPath)) {
                continue
            }

            if ($comparisonPath.Equals($normalizedCandidate, [System.StringComparison]::OrdinalIgnoreCase)) {
                return $true
            }

            if ($comparisonPath.StartsWith("$normalizedCandidate\", [System.StringComparison]::OrdinalIgnoreCase)) {
                return $true
            }
        }
    }

    return $false
}

function Test-IsAlwaysExcludedFolder {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Folder
    )

    $folderType = if ($Folder.PSObject.Properties['type']) { [string]$Folder.type } else { '' }
    foreach ($excludedFolderType in $script:AlwaysExcludedFolderTypes) {
        if ($folderType.Equals($excludedFolderType, [System.StringComparison]::OrdinalIgnoreCase)) {
            return $true
        }
    }

    $displayName = if ($Folder.PSObject.Properties['displayName']) { [string]$Folder.displayName } else { '' }
    foreach ($excludedDisplayName in $script:AlwaysExcludedFolderDisplayNames) {
        if ($displayName.Equals($excludedDisplayName, [System.StringComparison]::OrdinalIgnoreCase)) {
            return $true
        }
    }

    return $false
}

function Test-IsCalendarFolder {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Folder
    )

    if (-not $Folder.PSObject.Properties['type']) {
        return $false
    }

    return ([string]$Folder.type).Equals('IPF.Appointment', [System.StringComparison]::OrdinalIgnoreCase)
}

function Get-SpecialFolderRootDescriptor {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Folder
    )

    if (-not $Folder.PSObject.Properties['type']) {
        return $null
    }

    $folderType = [string]$Folder.type
    switch -Regex ($folderType) {
        '^IPF\.Appointment$' {
            return [pscustomobject]@{
                FolderType            = 'IPF.Appointment'
                CanonicalDisplayNames = @('Calendar')
                FriendlyName          = 'Calendar'
            }
        }
        '^IPF\.Task$' {
            return [pscustomobject]@{
                FolderType            = 'IPF.Task'
                CanonicalDisplayNames = @('Tasks')
                FriendlyName          = 'Tasks'
            }
        }
        '^IPF\.StickyNote$' {
            return [pscustomobject]@{
                FolderType            = 'IPF.StickyNote'
                CanonicalDisplayNames = @('Notes')
                FriendlyName          = 'Notes'
            }
        }
        '^IPF\.Contact$' {
            return [pscustomobject]@{
                FolderType            = 'IPF.Contact'
                CanonicalDisplayNames = @('Contacts')
                FriendlyName          = 'Contacts'
            }
        }
    }

    return $null
}

function Resolve-DefaultSpecialFolder {
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$FolderTree,

        [Parameter(Mandatory = $true)]
        [string]$FolderType,

        [Parameter(Mandatory = $true)]
        [string[]]$CanonicalDisplayNames,

        [Parameter(Mandatory = $true)]
        [string]$FriendlyName
    )

    $topLevelMatchingFolders = @(
        $FolderTree.TopLevelFolders |
        Where-Object {
            $_ -and
            $_.PSObject.Properties['type'] -and
            ([string]$_.type).Equals($FolderType, [System.StringComparison]::OrdinalIgnoreCase)
        }
    )

    if ($topLevelMatchingFolders.Count -eq 0) {
        throw "Unable to locate the target mailbox primary $FriendlyName folder."
    }

    if ($topLevelMatchingFolders.Count -eq 1) {
        return $topLevelMatchingFolders[0]
    }

    $canonicalNamedFolder = @(
        $topLevelMatchingFolders |
        Where-Object {
            $currentFolder = $_
            $currentFolder.PSObject.Properties['displayName'] -and
            @(
                $CanonicalDisplayNames |
                Where-Object {
                    $candidateName = [string]$_
                    -not [string]::IsNullOrWhiteSpace($candidateName) -and
                    $candidateName.Equals([string]$currentFolder.displayName, [System.StringComparison]::OrdinalIgnoreCase)
                }
            ).Count -gt 0
        }
    )

    if ($canonicalNamedFolder.Count -eq 1) {
        return $canonicalNamedFolder[0]
    }

    throw "Multiple top-level $FriendlyName folders were found, so the target primary $FriendlyName folder could not be determined automatically."
}

function Get-TreeChildFolders {
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$FolderTree,

        [AllowNull()]
        [object]$ParentFolder
    )

    if ($null -eq $ParentFolder) {
        return @($FolderTree.TopLevelFolders)
    }

    return @($FolderTree.ChildrenByParentId[$ParentFolder.id])
}

function Find-UniqueTargetChildFolder {
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$FolderTree,

        [AllowNull()]
        [object]$ParentFolder,

        [Parameter(Mandatory = $true)]
        [string]$DisplayName,

        [Parameter(Mandatory = $true)]
        [string]$ContextDescription
    )

    $matchingFolders = @(
        Get-TreeChildFolders -FolderTree $FolderTree -ParentFolder $ParentFolder |
        Where-Object {
            $_ -and
            $_.PSObject.Properties['displayName'] -and
            ([string]$_.displayName).Equals($DisplayName, [System.StringComparison]::OrdinalIgnoreCase)
        }
    )

    if ($matchingFolders.Count -gt 1) {
        $parentLabel = if ($null -eq $ParentFolder) { '\' } else { $ParentFolder.displayName }
        throw "Preflight failed: multiple target folders named '$DisplayName' already exist under '$parentLabel'. Unable to determine which folder to use for $ContextDescription."
    }

    if ($matchingFolders.Count -eq 1) {
        return $matchingFolders[0]
    }

    return $null
}

function Assert-TargetFolderPathIsUnambiguous {
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$FolderTree,

        [AllowNull()]
        [object]$ParentFolder,

        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [string]$ContextDescription
    )

    $segments = @($Path -split '[\\/]' | Where-Object { $_ -and $_.Trim() })
    $currentFolder = $ParentFolder

    foreach ($segment in $segments) {
        $currentFolder = Find-UniqueTargetChildFolder `
            -FolderTree $FolderTree `
            -ParentFolder $currentFolder `
            -DisplayName $segment `
            -ContextDescription $ContextDescription

        if (-not $currentFolder) {
            return
        }
    }
}

function Assert-CopyTargetPreflight {
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$SourceFolderTree,

        [Parameter(Mandatory = $true)]
        [object]$SourceRootFolder,

        [Parameter(Mandatory = $true)]
        [hashtable]$TargetFolderTree,

        [AllowNull()]
        [object]$TargetParentFolder,

        [Parameter(Mandatory = $true)]
        [switch]$ImportDirectlyIntoTargetFolder,

        [Parameter(Mandatory = $true)]
        [hashtable]$FolderSelectionPlan,

        [AllowNull()]
        [string]$RootTargetDisplayNameOverride
    )

    $simulatedFolderMap = @{}
    $rootTargetFolder = if ($SourceRootFolder.id -eq '$root') {
        $TargetParentFolder
    }
    elseif ($ImportDirectlyIntoTargetFolder) {
        $TargetParentFolder
    }
    else {
        $rootDisplayName = if ([string]::IsNullOrWhiteSpace($RootTargetDisplayNameOverride)) { $SourceRootFolder.displayName } else { $RootTargetDisplayNameOverride }
        $rootTargetFolder = Find-UniqueTargetChildFolder `
            -FolderTree $TargetFolderTree `
            -ParentFolder $TargetParentFolder `
            -DisplayName $rootDisplayName `
            -ContextDescription "the destination root folder for source '$($SourceRootFolder.displayName)'"

        if (-not $rootTargetFolder) {
            $rootTargetFolder = [pscustomobject]@{
                id          = '$preflight-root'
                displayName = $rootDisplayName
            }
        }

        $simulatedFolderMap[$SourceRootFolder.id] = $rootTargetFolder
    }

    if ($SourceRootFolder.id -ne '$root' -and -not $simulatedFolderMap.ContainsKey($SourceRootFolder.id)) {
        $simulatedFolderMap[$SourceRootFolder.id] = $rootTargetFolder
    }

    $pendingFolders = [System.Collections.Generic.Queue[object]]::new()
    $pendingFolders.Enqueue($SourceRootFolder)

    while ($pendingFolders.Count -gt 0) {
        $sourceFolder = $pendingFolders.Dequeue()

        $childFolders = if ($sourceFolder.id -eq '$root') {
            @($SourceFolderTree.TopLevelFolders | Where-Object { $_.PSObject.Properties['id'] -and $FolderSelectionPlan.TraversalFolderIds.ContainsKey($_.id) })
        }
        else {
            @(@($SourceFolderTree.ChildrenByParentId[$sourceFolder.id]) | Where-Object { $_ -and $_.PSObject.Properties['id'] -and $FolderSelectionPlan.TraversalFolderIds.ContainsKey($_.id) })
        }

        foreach ($childFolder in $childFolders) {
            $parentTargetFolder = if ($simulatedFolderMap.ContainsKey($sourceFolder.id)) { $simulatedFolderMap[$sourceFolder.id] } else { $rootTargetFolder }
            $resolvedTargetFolder = Find-UniqueTargetChildFolder `
                -FolderTree $TargetFolderTree `
                -ParentFolder $parentTargetFolder `
                -DisplayName $childFolder.displayName `
                -ContextDescription "source folder '$((Get-MailboxFolderDisplayPath -FolderTree $SourceFolderTree -FolderId $childFolder.id))'"

            if (-not $resolvedTargetFolder) {
                $resolvedTargetFolder = [pscustomobject]@{
                    id          = '$preflight-child'
                    displayName = $childFolder.displayName
                }
            }

            $simulatedFolderMap[$childFolder.id] = $resolvedTargetFolder
            $pendingFolders.Enqueue($childFolder)
        }
    }
}

function New-FolderSelectionPlan {
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$FolderTree,

        [Parameter(Mandatory = $true)]
        [object]$SourceRootFolder,

        [string[]]$IncludeFolderPath,

        [string[]]$ExcludeFolderPath
    )

    $sourceRootFullPath = Get-MailboxFolderDisplayPath -FolderTree $FolderTree -FolderId $SourceRootFolder.id
    $relativePathById = @{}
    $fullPathById = @{}
    $selectedFolderIds = @{}
    $traversalFolderIds = @{}
    $alwaysExcludedFolderIds = @{}
    $folderQueue = [System.Collections.Generic.Queue[object]]::new()
    $folderQueue.Enqueue($SourceRootFolder)

    while ($folderQueue.Count -gt 0) {
        $folder = $folderQueue.Dequeue()
        $fullPath = Get-MailboxFolderDisplayPath -FolderTree $FolderTree -FolderId $folder.id
        $relativePath = Get-RelativeFolderPath -FullPath $fullPath -RootPath $sourceRootFullPath
        $relativePathById[$folder.id] = $relativePath
        $fullPathById[$folder.id] = $fullPath

        if ($folder.id -ne '$root' -and (Test-IsAlwaysExcludedFolder -Folder $folder)) {
            $alwaysExcludedFolderIds[$folder.id] = $true
        }

        $childFolders = if ($folder.id -eq '$root') {
            @($FolderTree.TopLevelFolders)
        }
        else {
            @($FolderTree.ChildrenByParentId[$folder.id])
        }

        foreach ($childFolder in $childFolders) {
            $folderQueue.Enqueue($childFolder)
        }
    }

    foreach ($folderId in $relativePathById.Keys) {
        if ($folderId -eq '$root') {
            continue
        }

        $isAlwaysExcluded = $false
        $ancestorFolderId = $folderId
        while ($ancestorFolderId) {
            if ($alwaysExcludedFolderIds.ContainsKey($ancestorFolderId)) {
                $isAlwaysExcluded = $true
                break
            }

            if ($ancestorFolderId -eq '$root') {
                break
            }

            $ancestorFolder = $FolderTree.ById[$ancestorFolderId]
            if (-not $ancestorFolder) {
                break
            }

            $ancestorFolderId = if ($ancestorFolder.parentFolderId) { $ancestorFolder.parentFolderId } else { '$root' }
        }

        if ($isAlwaysExcluded) {
            continue
        }

        $isSelected = $true
        if ($IncludeFolderPath) {
            $isSelected = Test-FolderPathMatch -FolderRelativePath $relativePathById[$folderId] -FolderFullPath $fullPathById[$folderId] -CandidatePaths $IncludeFolderPath
        }
        elseif ($ExcludeFolderPath) {
            $isSelected = -not (Test-FolderPathMatch -FolderRelativePath $relativePathById[$folderId] -FolderFullPath $fullPathById[$folderId] -CandidatePaths $ExcludeFolderPath)
        }

        if (-not $isSelected) {
            continue
        }

        $selectedFolderIds[$folderId] = $true
        $currentFolderId = $folderId
        while ($currentFolderId) {
            $traversalFolderIds[$currentFolderId] = $true
            if ($currentFolderId -eq '$root') {
                break
            }

            $currentFolder = $FolderTree.ById[$currentFolderId]
            if (-not $currentFolder) {
                break
            }

            $currentFolderId = if ($currentFolder.parentFolderId) { $currentFolder.parentFolderId } else { '$root' }
        }
    }

    if (-not $traversalFolderIds.ContainsKey('$root')) {
        $traversalFolderIds['$root'] = $true
    }

    return @{
        RelativePathById        = $relativePathById
        FullPathById            = $fullPathById
        SelectedFolderIds       = $selectedFolderIds
        TraversalFolderIds      = $traversalFolderIds
        AlwaysExcludedFolderIds = $alwaysExcludedFolderIds
    }
}

function Ensure-MailboxFolderPath {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter(Mandatory = $true)]
        [string]$AccessToken,

        [Parameter(Mandatory = $true)]
        [string]$MailboxId,

        [AllowNull()]
        [object]$ParentFolder,

        [Parameter(Mandatory = $true)]
        [string]$Path,

        [switch]$CreateIfMissing = $true,

        [string]$DefaultFolderType = 'IPF.Note'
    )

    $segments = @($Path -split '[\\/]' | Where-Object { $_ -and $_.Trim() })
    if ($segments.Count -eq 0) {
        return $ParentFolder
    }

    $currentFolder = $ParentFolder
    foreach ($segment in $segments) {
        $siblingFolders = if ($null -eq $currentFolder) {
            Invoke-GraphCollectionRequest -AccessToken $AccessToken -Uri "/beta/admin/exchange/mailboxes/$([Uri]::EscapeDataString($MailboxId))/folders"
        }
        else {
            Invoke-GraphCollectionRequest -AccessToken $AccessToken -Uri "/beta/admin/exchange/mailboxes/$([Uri]::EscapeDataString($MailboxId))/folders/$([Uri]::EscapeDataString($currentFolder.id))/childFolders"
        }

        $existingFolder = @(
            $siblingFolders |
            Where-Object { $_.PSObject.Properties['displayName'] -and $_.displayName -eq $segment }
        )

        if ($existingFolder.Count -gt 1) {
            throw "Multiple target folders named '$segment' already exist under the selected parent while resolving path '$Path'."
        }

        if ($existingFolder.Count -eq 1) {
            $currentFolder = $existingFolder[0]
            continue
        }

        if (-not $CreateIfMissing) {
            if ($PSCmdlet.ShouldProcess($Path, "Create missing folder path segment '$segment'")) {
                $currentFolder = Ensure-MailboxFolder -AccessToken $AccessToken -MailboxId $MailboxId -ParentFolder $currentFolder -DisplayName $segment -FolderType $DefaultFolderType
            }
            else {
                $currentFolder = [pscustomobject]@{
                    id          = '$whatif'
                    displayName = $segment
                    type        = $DefaultFolderType
                }
            }

            continue
        }

        $currentFolder = Ensure-MailboxFolder -AccessToken $AccessToken -MailboxId $MailboxId -ParentFolder $currentFolder -DisplayName $segment -FolderType $DefaultFolderType
    }

    return $currentFolder
}

function Ensure-MailboxFolder {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter(Mandatory = $true)]
        [string]$AccessToken,

        [Parameter(Mandatory = $true)]
        [string]$MailboxId,

        [AllowNull()]
        [object]$ParentFolder,

        [Parameter(Mandatory = $true)]
        [string]$DisplayName,

        [Parameter(Mandatory = $true)]
        [string]$FolderType
    )

    $encodedMailboxId = [Uri]::EscapeDataString($MailboxId)
    $resolvedFolderType = if ([string]::IsNullOrWhiteSpace($FolderType)) { 'IPF.Note' } else { $FolderType }

    if ($ParentFolder -and $ParentFolder.PSObject.Properties['id'] -and $ParentFolder.id -eq '$whatif') {
        return [pscustomobject]@{
            id          = '$whatif'
            displayName = $DisplayName
            type        = $resolvedFolderType
        }
    }

    $siblingFolders = if ($null -eq $ParentFolder) {
        Invoke-GraphCollectionRequest -AccessToken $AccessToken -Uri "/beta/admin/exchange/mailboxes/$encodedMailboxId/folders"
    }
    else {
        Invoke-GraphCollectionRequest -AccessToken $AccessToken -Uri "/beta/admin/exchange/mailboxes/$encodedMailboxId/folders/$([Uri]::EscapeDataString($ParentFolder.id))/childFolders"
    }

    $existing = @(
        $siblingFolders |
        Where-Object { $_.PSObject.Properties['displayName'] -and $_.displayName -eq $DisplayName }
    )
    if ($existing.Count -gt 1) {
        throw "Multiple target folders named '$DisplayName' already exist under the selected parent."
    }

    if ($existing.Count -eq 1) {
        return $existing[0]
    }

    $createUri = if ($null -eq $ParentFolder) {
        "/beta/admin/exchange/mailboxes/$encodedMailboxId/folders"
    }
    else {
        "/beta/admin/exchange/mailboxes/$encodedMailboxId/folders/$([Uri]::EscapeDataString($ParentFolder.id))/childFolders"
    }

    Write-Verbose "Creating target folder '$DisplayName' with type '$resolvedFolderType'."
    if ($PSCmdlet.ShouldProcess($MailboxId, "Create target folder '$DisplayName'")) {
        Invoke-GraphApiRequest -AccessToken $AccessToken -Uri $createUri -Method POST -Body @{
            displayName = $DisplayName
            type        = $resolvedFolderType
        }
    }
    else {
        [pscustomobject]@{
            id          = '$whatif'
            displayName = $DisplayName
            type        = $resolvedFolderType
        }
    }
}

function Get-MailboxItems {
    param(
        [Parameter(Mandatory = $true)]
        [string]$AccessToken,

        [Parameter(Mandatory = $true)]
        [string]$MailboxId,

        [Parameter(Mandatory = $true)]
        [string]$FolderId
        ,

        [AllowNull()]
        [object]$DateFilter
    )

    $encodedMailboxId = [Uri]::EscapeDataString($MailboxId)
    $encodedFolderId = [Uri]::EscapeDataString($FolderId)
    $queryParameters = @(
        '$select=id,createdDateTime'
    )

    if ($DateFilter -and $DateFilter.FilterText) {
        $queryParameters += '$filter={0}' -f [Uri]::EscapeDataString($DateFilter.FilterText)
    }

    $queryString = $queryParameters -join '&'
    @(
        Invoke-GraphCollectionRequest -AccessToken $AccessToken -Uri "/beta/admin/exchange/mailboxes/$encodedMailboxId/folders/$encodedFolderId/items?$queryString" |
        Where-Object { $_.PSObject.Properties['id'] }
    )
}

function Export-MailboxItems {
    param(
        [Parameter(Mandatory = $true)]
        [string]$AccessToken,

        [Parameter(Mandatory = $true)]
        [string]$MailboxId,

        [Parameter(Mandatory = $true)]
        [string[]]$ItemIds
    )

    $encodedMailboxId = [Uri]::EscapeDataString($MailboxId)
    $response = Invoke-GraphApiRequest -AccessToken $AccessToken -Uri "/beta/admin/exchange/mailboxes/$encodedMailboxId/exportItems" -Method POST -Body @{
        itemIds = $ItemIds
    }

    return @($response.value)
}

function Get-ImportSession {
    param(
        [Parameter(Mandatory = $true)]
        [string]$AccessToken,

        [Parameter(Mandatory = $true)]
        [string]$MailboxId,

        [AllowNull()]
        [hashtable]$ExistingSession
    )

    if ($ExistingSession) {
        $expires = [DateTimeOffset]::Parse($ExistingSession.ExpirationDateTime)
        if ($expires -gt [DateTimeOffset]::UtcNow.AddMinutes(5)) {
            return $ExistingSession
        }
    }

    $encodedMailboxId = [Uri]::EscapeDataString($MailboxId)
    $response = Invoke-GraphApiRequest -AccessToken $AccessToken -Uri "/beta/admin/exchange/mailboxes/$encodedMailboxId/createImportSession" -Method POST

    return @{
        ImportUrl          = $response.importUrl
        ExpirationDateTime = $response.expirationDateTime
    }
}

function Import-MailboxItem {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ImportUrl,

        [Parameter(Mandatory = $true)]
        [string]$FolderId,

        [Parameter(Mandatory = $true)]
        [string]$Data
    )

    Invoke-RestMethod -Method Post -Uri $ImportUrl -ContentType 'application/json' -Body (@{
        FolderId = $FolderId
        Mode     = 'create'
        Data     = $Data
    } | ConvertTo-Json -Compress)
}

function Copy-MailboxFolderItems {
    param(
        [Parameter(Mandatory = $true)]
        [string]$AccessToken,

        [Parameter(Mandatory = $true)]
        [string]$SourceMailboxId,

        [Parameter(Mandatory = $true)]
        [string]$TargetMailboxId,

        [Parameter(Mandatory = $true)]
        [hashtable]$SourceFolderTree,

        [Parameter(Mandatory = $true)]
        [object]$SourceRootFolder,

        [AllowNull()]
        [object]$TargetParentFolder,

        [Parameter(Mandatory = $true)]
        [switch]$OverlayMode,

        [Parameter(Mandatory = $true)]
        [switch]$ImportDirectlyIntoTargetFolder,

        [Parameter(Mandatory = $true)]
        [switch]$CopyEmptyFolders,

        [Parameter(Mandatory = $true)]
        [hashtable]$FolderSelectionPlan,

        [AllowNull()]
        [string]$RootTargetDisplayNameOverride,

        [AllowNull()]
        [object]$DateFilter,

        [Parameter(Mandatory = $true)]
        [int]$ExportBatchSize
    )

    $importSession = $null
    $folderMap = @{}
    $foldersToProcess = [System.Collections.Generic.List[object]]::new()
    $folderCountQueue = [System.Collections.Generic.Queue[object]]::new()
    $folderCountQueue.Enqueue($SourceRootFolder)

    while ($folderCountQueue.Count -gt 0) {
        $queuedFolder = $folderCountQueue.Dequeue()
        if (-not $FolderSelectionPlan.TraversalFolderIds.ContainsKey($queuedFolder.id)) {
            continue
        }

        $foldersToProcess.Add($queuedFolder)

        $queuedChildren = if ($queuedFolder.id -eq '$root') {
            @($SourceFolderTree.TopLevelFolders | Where-Object { $_.PSObject.Properties['id'] -and $FolderSelectionPlan.TraversalFolderIds.ContainsKey($_.id) })
        }
        else {
            @(@($SourceFolderTree.ChildrenByParentId[$queuedFolder.id]) | Where-Object { $_ -and $_.PSObject.Properties['id'] -and $FolderSelectionPlan.TraversalFolderIds.ContainsKey($_.id) })
        }

        foreach ($queuedChildFolder in $queuedChildren) {
            $folderCountQueue.Enqueue($queuedChildFolder)
        }
    }

    $totalFolders = $foldersToProcess.Count
    $totalEstimatedItems = (@($foldersToProcess) | Measure-Object -Property totalItemCount -Sum).Sum
    if ($null -eq $totalEstimatedItems) {
        $totalEstimatedItems = 0
    }
    $processedFolderCount = 0
    $copiedItemCount = 0

    $rootTargetFolder = if ($SourceRootFolder.id -eq '$root') {
        $TargetParentFolder
    }
    elseif ($ImportDirectlyIntoTargetFolder) {
        $TargetParentFolder
    }
    else {
        $sourceRootFolderType = if ([string]::IsNullOrWhiteSpace($SourceRootFolder.type)) { 'IPF.Note' } else { [string]$SourceRootFolder.type }
        $rootDisplayName = if ([string]::IsNullOrWhiteSpace($RootTargetDisplayNameOverride)) { $SourceRootFolder.displayName } else { $RootTargetDisplayNameOverride }
        Ensure-MailboxFolder -AccessToken $AccessToken -MailboxId $TargetMailboxId -ParentFolder $TargetParentFolder -DisplayName $rootDisplayName -FolderType $sourceRootFolderType
    }

    if ($null -eq $rootTargetFolder -and -not ($OverlayMode -and $SourceRootFolder.id -eq '$root')) {
        throw 'The destination root folder could not be resolved.'
    }

    $folderMap[$SourceRootFolder.id] = $rootTargetFolder
    $pendingFolders = [System.Collections.Generic.Queue[object]]::new()
    $pendingFolders.Enqueue($SourceRootFolder)

    while ($pendingFolders.Count -gt 0) {
        $sourceFolder = $pendingFolders.Dequeue()
        $targetFolder = if ($folderMap.ContainsKey($sourceFolder.id)) { $folderMap[$sourceFolder.id] } else { $null }
        $processedFolderCount++
        $sourceFolderPath = Get-MailboxFolderDisplayPath -FolderTree $SourceFolderTree -FolderId $sourceFolder.id
        $estimatedItemsInFolder = if ($null -ne $sourceFolder.totalItemCount) { [int]$sourceFolder.totalItemCount } else { 0 }
        $isSelectedFolder = $FolderSelectionPlan.SelectedFolderIds.ContainsKey($sourceFolder.id)

        Write-Host ("[{0}/{1}] Processing folder: {2}" -f $processedFolderCount, $totalFolders, $sourceFolderPath)
        Write-Progress -Id 1 -Activity 'Copying mailbox folders' -Status $sourceFolderPath -PercentComplete (($processedFolderCount / [Math]::Max($totalFolders, 1)) * 100)

        if ($sourceFolder.id -eq '$root') {
            $items = @()
            Write-Host '  Root mailbox container selected; processing child folders only.'
        }
        elseif (-not $isSelectedFolder) {
            $items = @()
            Write-Host '  Folder skipped for item copy; evaluating descendants only.'
        }
        else {
            Write-Verbose "Enumerating items in source folder '$($sourceFolder.displayName)' ($($sourceFolder.id))."
            $items = @(Get-MailboxItems -AccessToken $AccessToken -MailboxId $SourceMailboxId -FolderId $sourceFolder.id -DateFilter $DateFilter)
        }

        $itemCount = @($items).Count

        if ($itemCount -gt 0) {
            Write-Verbose "Found $itemCount item(s) in '$($sourceFolder.displayName)'."
        }
        elseif ($sourceFolder.id -ne '$root') {
            Write-Host "  No items found in this folder."
        }

        $childFolders = if ($sourceFolder.id -eq '$root') {
            @($SourceFolderTree.TopLevelFolders | Where-Object { $_.PSObject.Properties['id'] -and $FolderSelectionPlan.TraversalFolderIds.ContainsKey($_.id) })
        }
        else {
            @(@($SourceFolderTree.ChildrenByParentId[$sourceFolder.id]) | Where-Object { $_ -and $_.PSObject.Properties['id'] -and $FolderSelectionPlan.TraversalFolderIds.ContainsKey($_.id) })
        }

        $childFolderCount = @($childFolders).Count
        $hasTraversedChildren = $childFolderCount -gt 0

        if ($sourceFolder.id -ne '$root' -and -not $targetFolder) {
            $shouldCreateCurrentFolder = $CopyEmptyFolders -or $itemCount -gt 0 -or $hasTraversedChildren
            if ($shouldCreateCurrentFolder) {
                $parentSourceFolderId = if ($sourceFolder.parentFolderId) { $sourceFolder.parentFolderId } else { '$root' }
                $targetParentForCurrent = if ($folderMap.ContainsKey($parentSourceFolderId)) { $folderMap[$parentSourceFolderId] } else { $rootTargetFolder }
                $currentFolderType = if ([string]::IsNullOrWhiteSpace($sourceFolder.type)) { 'IPF.Note' } else { [string]$sourceFolder.type }
                $targetFolder = Ensure-MailboxFolder -AccessToken $AccessToken -MailboxId $TargetMailboxId -ParentFolder $targetParentForCurrent -DisplayName $sourceFolder.displayName -FolderType $currentFolderType
                $folderMap[$sourceFolder.id] = $targetFolder
            }
        }

        $folderImportedCount = 0
        for ($chunkStart = 0; $chunkStart -lt $itemCount; $chunkStart += $ExportBatchSize) {
            $chunk = @($items[$chunkStart..([Math]::Min($chunkStart + $ExportBatchSize - 1, $itemCount - 1))])
            $itemIds = @($chunk | ForEach-Object { $_.id })
            $chunkEnd = [Math]::Min($chunkStart + $ExportBatchSize, $itemCount)

            if ($itemIds.Count -eq 0) {
                continue
            }

            Write-Host ("  Batch {0}-{1} of {2} item(s) in folder" -f ($chunkStart + 1), $chunkEnd, $itemCount)
            Write-Progress -Id 2 -ParentId 1 -Activity 'Copying items in current folder' -Status "$sourceFolderPath ($chunkEnd of $itemCount)" -PercentComplete (($chunkEnd / [Math]::Max($itemCount, 1)) * 100)

            if (-not $targetFolder) {
                throw "A target folder was not available for '$sourceFolderPath' even though items were selected for copy."
            }

            if ($PSCmdlet.ShouldProcess("$($TargetMailboxId):$($targetFolder.displayName)", "Import $($itemIds.Count) item(s) from '$($SourceMailboxId):$($sourceFolder.displayName)'")) {
                $exportedItems = Export-MailboxItems -AccessToken $AccessToken -MailboxId $SourceMailboxId -ItemIds $itemIds
                $importSession = Get-ImportSession -AccessToken $AccessToken -MailboxId $TargetMailboxId -ExistingSession $importSession

                foreach ($exportedItem in $exportedItems) {
                    $errorProperty = $exportedItem.PSObject.Properties['error']
                    if ($errorProperty -and $null -ne $errorProperty.Value) {
                        throw "Graph exportItems returned an error for source item '$($exportedItem.itemId)'."
                    }

                    Import-MailboxItem -ImportUrl $importSession.ImportUrl -FolderId $targetFolder.id -Data $exportedItem.data | Out-Null
                    $folderImportedCount++
                    $copiedItemCount++
                }

                Write-Host ("    Imported {0} item(s) this batch. Running total: {1}" -f $itemIds.Count, $copiedItemCount)
                if ($totalEstimatedItems -gt 0) {
                    Write-Progress -Id 3 -Activity 'Overall item progress' -Status "$copiedItemCount of estimated $totalEstimatedItems item(s) imported" -PercentComplete (($copiedItemCount / $totalEstimatedItems) * 100)
                }
            }
        }

        if ($itemCount -gt 0) {
            Write-Host ("  Completed folder: imported {0} of {1} discovered item(s)." -f $folderImportedCount, $itemCount)
        }

        Write-Progress -Id 2 -ParentId 1 -Activity 'Copying items in current folder' -Completed
        if ($processedFolderCount -eq $totalFolders) {
            Write-Progress -Id 1 -Activity 'Copying mailbox folders' -Completed
            Write-Progress -Id 3 -Activity 'Overall item progress' -Completed
        }

        foreach ($childFolder in $childFolders) {
            $pendingFolders.Enqueue($childFolder)
        }
    }

    Write-Host ("Finished. Processed {0} folder(s) and imported {1} item(s)." -f $processedFolderCount, $copiedItemCount)
}

$accessToken = Get-GraphAccessToken -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint

Write-Verbose "Resolving source mailbox ID for '$SourceUserPrincipalName'."
$sourceMailboxId = Resolve-MailboxId -AccessToken $accessToken -UserPrincipalName $SourceUserPrincipalName

Write-Verbose "Resolving target mailbox ID for '$TargetUserPrincipalName'."
$targetMailboxId = Resolve-MailboxId -AccessToken $accessToken -UserPrincipalName $TargetUserPrincipalName

Write-Verbose "Loading source mailbox folder tree."
$sourceFolderTree = Get-MailboxFolderTree -AccessToken $accessToken -MailboxId $sourceMailboxId

Write-Verbose "Loading target mailbox folder tree."
$targetFolderTree = Get-MailboxFolderTree -AccessToken $accessToken -MailboxId $targetMailboxId

$sourceRootFolder = Resolve-MailboxFolderPath -FolderTree $sourceFolderTree -Path $SourceFolderPath

if ($OverlayMode -and $sourceRootFolder.id -ne '$root') {
    throw "OverlayMode currently supports only SourceFolderPath '\' so the source mailbox root can be merged into the target structure."
}

$folderSelectionPlan = New-FolderSelectionPlan -FolderTree $sourceFolderTree -SourceRootFolder $sourceRootFolder -IncludeFolderPath $IncludeFolderPath -ExcludeFolderPath $ExcludeFolderPath
$itemDateFilter = New-MailboxItemDateFilter -Oldest $Oldest -Newest $Newest
if ($itemDateFilter) {
    Write-Verbose "Applying item createdDateTime filter: $($itemDateFilter.FilterText)"
}
$estimatedSelectedFolders = @(
    $sourceFolderTree.AllFolders |
    Where-Object { $_.PSObject.Properties['id'] -and $folderSelectionPlan.SelectedFolderIds.ContainsKey($_.id) }
).Count
$estimatedTraversedFolders = @(
    if ($folderSelectionPlan.TraversalFolderIds.ContainsKey('$root')) {
        [pscustomobject]@{ id = '$root'; totalItemCount = 0 }
    }
    $sourceFolderTree.AllFolders |
    Where-Object { $_.PSObject.Properties['id'] -and $folderSelectionPlan.TraversalFolderIds.ContainsKey($_.id) }
).Count
$estimatedItems = (@(
    $sourceFolderTree.AllFolders |
    Where-Object { $_.PSObject.Properties['id'] -and $folderSelectionPlan.SelectedFolderIds.ContainsKey($_.id) }
) | Measure-Object -Property totalItemCount -Sum).Sum
if ($null -eq $estimatedItems) {
    $estimatedItems = 0
}
$sourceSpecialFolderDescriptor = Get-SpecialFolderRootDescriptor -Folder $sourceRootFolder
$rootTargetDisplayNameOverride = $null

if ($sourceSpecialFolderDescriptor) {
    $targetPrimarySpecialFolder = Resolve-DefaultSpecialFolder `
        -FolderTree $targetFolderTree `
        -FolderType $sourceSpecialFolderDescriptor.FolderType `
        -CanonicalDisplayNames $sourceSpecialFolderDescriptor.CanonicalDisplayNames `
        -FriendlyName $sourceSpecialFolderDescriptor.FriendlyName

    if ($ImportDirectlyIntoTargetFolder) {
        Write-Verbose "$($sourceSpecialFolderDescriptor.FriendlyName) source detected. ImportDirectlyIntoTargetFolder will merge items directly into target $($sourceSpecialFolderDescriptor.FriendlyName) folder '$($targetPrimarySpecialFolder.displayName)'."
        if (-not [string]::IsNullOrWhiteSpace($TargetFolderPath)) {
            Write-Verbose "Ignoring TargetFolderPath '$TargetFolderPath' because special-folder direct import always uses the target primary $($sourceSpecialFolderDescriptor.FriendlyName) folder."
        }

        $targetParentFolder = $targetPrimarySpecialFolder
    }
    else {
        $specialTargetName = Normalize-FolderPath -Path $TargetFolderPath
        if ([string]::IsNullOrWhiteSpace($specialTargetName)) {
            $specialTargetName = $sourceRootFolder.displayName
        }

        if ($specialTargetName.Contains('\')) {
            throw "$($sourceSpecialFolderDescriptor.FriendlyName) target name '$specialTargetName' is invalid. Special-folder copies can only create a single named subfolder under the target primary $($sourceSpecialFolderDescriptor.FriendlyName) folder."
        }

        Write-Verbose "$($sourceSpecialFolderDescriptor.FriendlyName) source detected. A target subfolder named '$specialTargetName' will be created under '$($targetPrimarySpecialFolder.displayName)'."
        $targetParentFolder = $targetPrimarySpecialFolder
        $rootTargetDisplayNameOverride = $specialTargetName
    }
}
else {
    $targetParentFolder = if ([string]::IsNullOrWhiteSpace($TargetFolderPath)) {
        $null
    }
    else {
        Assert-TargetFolderPathIsUnambiguous `
            -FolderTree $targetFolderTree `
            -ParentFolder $null `
            -Path $TargetFolderPath `
            -ContextDescription "the requested target folder path '$TargetFolderPath'"
        Ensure-MailboxFolderPath -AccessToken $accessToken -MailboxId $targetMailboxId -Path $TargetFolderPath -CreateIfMissing:$false
    }
}

if ($OverlayMode) {
    if (-not [string]::IsNullOrWhiteSpace($TargetFolderPath)) {
        Write-Verbose "OverlayMode will merge the source mailbox root into the target path '$TargetFolderPath' without creating a source-root container folder."
    }
    else {
        Write-Verbose 'OverlayMode will merge the source mailbox root directly into the target mailbox root without creating a source-root container folder.'
    }
}

$resolvedTargetDescription = if ($OverlayMode) {
    if ([string]::IsNullOrWhiteSpace($TargetFolderPath)) {
        'Target mailbox root'
    }
    else {
        "Overlay into target path '$TargetFolderPath'"
    }
}
elseif ($sourceSpecialFolderDescriptor -and $ImportDirectlyIntoTargetFolder) {
    "Primary target $($sourceSpecialFolderDescriptor.FriendlyName) folder"
}
elseif ($sourceSpecialFolderDescriptor -and -not [string]::IsNullOrWhiteSpace($rootTargetDisplayNameOverride)) {
    "$($sourceSpecialFolderDescriptor.FriendlyName) subfolder '$rootTargetDisplayNameOverride'"
}
elseif ([string]::IsNullOrWhiteSpace($TargetFolderPath)) {
    'Mailbox root or matching existing folders'
}
else {
    $TargetFolderPath
}

Confirm-PlannedOperation `
    -SourceUserPrincipalName $SourceUserPrincipalName `
    -TargetUserPrincipalName $TargetUserPrincipalName `
    -SourceFolderPath $SourceFolderPath `
    -ResolvedTargetDescription $resolvedTargetDescription `
    -Mode (Get-ExecutionModeSummary -OverlayMode:$OverlayMode -ImportDirectlyIntoTargetFolder:$ImportDirectlyIntoTargetFolder -PreflightOnly:$PreflightOnly) `
    -OverlayMode:$OverlayMode `
    -ImportDirectlyIntoTargetFolder:$ImportDirectlyIntoTargetFolder `
    -DateFilter $itemDateFilter `
    -IncludeFolderPath $IncludeFolderPath `
    -ExcludeFolderPath $ExcludeFolderPath `
    -SettingSources $settingSources `
    -EstimatedSelectedFolderCount $estimatedSelectedFolders `
    -EstimatedTraversedFolderCount $estimatedTraversedFolders `
    -EstimatedItemCount $estimatedItems `
    -CopyEmptyFolders:$CopyEmptyFolders `
    -PreflightOnly:$PreflightOnly `
    -WhatIfMode:$WhatIfPreference `
    -Force:$Force

Write-Verbose 'Running target folder preflight checks.'
Assert-CopyTargetPreflight `
    -SourceFolderTree $sourceFolderTree `
    -SourceRootFolder $sourceRootFolder `
    -TargetFolderTree $targetFolderTree `
    -TargetParentFolder $targetParentFolder `
    -ImportDirectlyIntoTargetFolder:$ImportDirectlyIntoTargetFolder `
    -FolderSelectionPlan $folderSelectionPlan `
    -RootTargetDisplayNameOverride $rootTargetDisplayNameOverride

if ($PreflightOnly) {
    Write-Host 'Preflight completed successfully. No copy was performed because PreflightOnly was specified.'
    return
}

Copy-MailboxFolderItems `
    -AccessToken $accessToken `
    -SourceMailboxId $sourceMailboxId `
    -TargetMailboxId $targetMailboxId `
    -SourceFolderTree $sourceFolderTree `
    -SourceRootFolder $sourceRootFolder `
    -TargetParentFolder $targetParentFolder `
    -OverlayMode:$OverlayMode `
    -ImportDirectlyIntoTargetFolder:$ImportDirectlyIntoTargetFolder `
    -CopyEmptyFolders:$CopyEmptyFolders `
    -FolderSelectionPlan $folderSelectionPlan `
    -RootTargetDisplayNameOverride $rootTargetDisplayNameOverride `
    -DateFilter $itemDateFilter `
    -ExportBatchSize $ExportBatchSize
