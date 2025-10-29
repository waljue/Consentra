<#
.SYNOPSIS
  Export OAuth delegated consents to CSV and HTML.
.DESCRIPTION
  Retrieves oauth2PermissionGrant data from Microsoft Graph, enriches it with
  service principal metadata, classifies each app by vendor, and generates both
  CSV and interactive HTML reports.
#>

[CmdletBinding()]
param(
    [string]$AppNameContains
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

#region Constants
$script:MicrosoftDomainSuffixes = @(
    'microsoft.com',
    'microsoftonline.com',
    'windows.net',
    'windows.com',
    'azure.com',
    'azure.net',
    'office.com',
    'office365.com',
    'live.com',
    'xboxlive.com',
    'microsoft.us'
)

$script:KnownMicrosoftTenantIds = @(
    'f8cdef31-a31e-4b4a-93e4-5f571e91255a',
    '9188040d-6c67-4c5b-b112-36a304b66dad',
    '00000000-0000-0000-0000-000000000000',
    '72f988bf-86f1-41af-91ab-2d7cd011db47'
)

$script:KnownMicrosoftAppIds = @(
    '00000003-0000-0000-c000-000000000000',
    '00000002-0000-0000-c000-000000000000',
    'd3590ed6-52b3-4102-aeff-aad2292ab01c',
    '6a3a5d73-3c83-4e2e-8c6f-ef807d52c6f3',
    '88cfea84-7919-4c1a-9c29-200b48e0b8fd',
    '00000004-0000-0ff1-ce00-000000000000',
    'fa4345a4-a730-4230-84a8-7d9651b86739',
    'b8456c59-1230-44c7-a4a2-99b085333e84',
    '6ba09155-cb24-475b-b24f-b4e28fc74365',
    '3e622cf1-17df-4a0d-b6f1-c7baabc1e37e',
    '5a0aa725-4958-4b0c-80a9-34562e23f3b7',
    'de8bc8b5-d9f9-48b1-a8ad-b748da725064',
    'd1ddf0e4-d672-4dae-b554-9d5bdfd93547',
    '7aa75825-1ea2-4cd1-a38c-90c595a91411',
    '92f79662-d42d-488f-9366-478da237826a',
    '14d82eec-204b-4c2f-b7e8-296a70dab67e',
    '08e18876-6177-487e-b8b5-cf950c1e598c',
    '5d5e5296-3121-49bf-bdbe-2cd00b9d8295',
    'de0853a1-ab20-47bd-990b-71ad5077ac7b',
    'ea66284f-8898-44f3-b935-d182ea20816c'
)
#endregion

function Get-EffectiveString {
    param([string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return $null }
    if ($Value -eq 'SkipToTheEndpoint') { return $null }
    return $Value
}

function Get-DomainFromServicePrincipalNames {
    param([string[]]$Names)
    if (-not $Names) { return $null }
    foreach ($name in $Names) {
        if ([string]::IsNullOrWhiteSpace($name)) { continue }
        $candidate = $name.Trim()
        try {
            if ($candidate -match '^[a-f0-9-]{36}$') { continue }
            if ($candidate -match '://') {
                $uri = [System.Uri]$candidate
                if ($uri.Host) { return $uri.Host.ToLowerInvariant() }
            } elseif ($candidate -like '*.*') {
                $parts = $candidate.Split('/')[-1]
                if ($parts -match '\.') {
                    return $parts.ToLowerInvariant()
                }
            }
        } catch {
            # ignore malformed values
        }
    }
    return $null
}

function Initialize-GraphContext {
    try { Get-MgContext | Out-Null }
    catch { Connect-MgGraph -Scopes 'Directory.Read.All','Application.Read.All' | Out-Null }
}

function Get-TenantMetadata {
    $uri = 'https://graph.microsoft.com/v1.0/organization?$select=id,displayName,verifiedDomains'
    $response = Invoke-MgGraphRequest -Method GET -Uri $uri

    $domains = @()
    $displayName = $null

    if ($response.value) {
        foreach ($org in $response.value) {
            if (-not $displayName -and $org.displayName) {
                $displayName = $org.displayName
            }
            if ($org.verifiedDomains) {
                foreach ($domain in $org.verifiedDomains) {
                    if ($domain.name) {
                        $domains += $domain.name.ToLowerInvariant()
                    }
                }
            }
        }
    }

    return [pscustomobject]@{
        DisplayName = $displayName
        Domains     = ($domains | Sort-Object -Unique)
    }
}

function Get-GraphResponseProperty {
    param(
        [Parameter(Mandatory)][object]$Response,
        [Parameter(Mandatory)][string]$PropertyName
    )

    $prop = $Response.PSObject.Properties[$PropertyName]
    if ($prop) {
        return $prop.Value
    }

    $ciProp = $Response.PSObject.Properties |
        Where-Object { $_.Name -ieq $PropertyName } |
        Select-Object -First 1
    if ($ciProp) {
        return $ciProp.Value
    }

    if ($Response -is [System.Collections.IDictionary]) {
        foreach ($key in $Response.Keys) {
            if ($key -ieq $PropertyName) {
                return $Response[$key]
            }
        }
    }

    $additionalProp = $Response.PSObject.Properties['AdditionalProperties']
    if ($additionalProp) {
        $additional = $additionalProp.Value
        if ($additional -is [System.Collections.IDictionary]) {
            foreach ($key in $additional.Keys) {
                if ($key -ieq $PropertyName) {
                    return $additional[$key]
                }
            }
        }
    }

    return $null
}

function Invoke-GraphPaged {
    param([Parameter(Mandatory)][string]$Uri)
    $headers = @{ 'ConsistencyLevel' = 'eventual' }
    $results = @()
    $nextLink = $Uri

    while ($nextLink) {
        $response = Invoke-MgGraphRequest -Method GET -Uri $nextLink -Headers $headers
        if ($response.value) { $results += $response.value }
        $nextLink = if ($response) { Get-GraphResponseProperty -Response $response -PropertyName '@odata.nextLink' } else { $null }
    }

    return $results
}

function Get-DirectoryObjectsByIds {
    param([string[]]$Ids)
    $idList = @($Ids | Where-Object { $_ })
    if ($idList.Count -eq 0) { return @() }

    $results = @()
    for ($index = 0; $index -lt $idList.Count; $index += 1000) {
        $chunk = @($idList[$index..([Math]::Min($index + 999, $idList.Count - 1))])
        $body = @{ ids = $chunk } | ConvertTo-Json
        $response = Invoke-MgGraphRequest -Method POST -Uri 'https://graph.microsoft.com/beta/directoryObjects/getByIds' -Body $body -ContentType 'application/json'
        if ($response.value) { $results += $response.value }
    }
    return $results
}

function Get-ServicePrincipalMetadata {
    param([string[]]$Ids)
    $idList = @($Ids | Where-Object { $_ })
    if ($idList.Count -eq 0) { return @{} }

    $results = @{}
    $batchEndpoint = 'https://graph.microsoft.com/beta/$batch'
    $select = 'id,appId,displayName,appOwnerOrganizationId,publisherName,tags,servicePrincipalType,servicePrincipalNames,verifiedPublisher'
    $batchSize = 20

    for ($index = 0; $index -lt $idList.Count; $index += $batchSize) {
        $chunk = @($idList[$index..([Math]::Min($index + $batchSize - 1, $idList.Count - 1))])
        $requests = @()
        foreach ($id in $chunk) {
            if ([string]::IsNullOrWhiteSpace($id)) { continue }
            $url = [string]::Format('servicePrincipals/{0}?$select={1}', $id, $select)
            $requests += @{
                id     = [Guid]::NewGuid().ToString()
                method = 'GET'
                url    = $url
            }
        }
        if (-not $requests) { continue }
        $body = @{ requests = $requests } | ConvertTo-Json -Depth 6
        $response = Invoke-MgGraphRequest -Method POST -Uri $batchEndpoint -Body $body -ContentType 'application/json'
        if (-not $response.responses) { continue }
        foreach ($entry in $response.responses) {
            if ($entry.status -ne 200 -or -not $entry.body -or -not $entry.body.id) { continue }
            $results[$entry.body.id] = $entry.body
        }
    }

    return $results
}

function Test-ServicePrincipalHidden {
    param([string[]]$Tags)
    if (-not $Tags) { return $false }
    foreach ($tag in $Tags) {
        if ([string]::IsNullOrWhiteSpace($tag)) { continue }
        switch -Regex ($tag) {
            '^HideApp$' { return $true }
            '^HideAppForManagedApps$' { return $true }
            '^NonVisibleToUsers$' { return $true }
        }
    }
    return $false
}

function Get-AllServicePrincipals {
    $select = 'id,appId,displayName,appOwnerOrganizationId,publisherName,publisherDomain,tags,servicePrincipalNames,servicePrincipalType,verifiedPublisher'
    $baseUri = "https://graph.microsoft.com/beta/servicePrincipals?\$count=true&\$top=999&includeHidden=true&\$select=$select"
    try {
        return Invoke-GraphPaged -Uri $baseUri
    } catch {
        Write-Warning ("Unable to include hidden service principals via includeHidden=true: {0}" -f $_.Exception.Message)
        $fallbackUri = "https://graph.microsoft.com/beta/servicePrincipals?\$count=true&\$top=999&\$select=$select"
        return Invoke-GraphPaged -Uri $fallbackUri
    }
}

function New-ServicePrincipalProfile {
    param([pscustomobject]$Source)
    if (-not $Source) { return $null }

    $tagsRaw = Get-GraphResponseProperty -Response $Source -PropertyName 'tags'
    $tags = @()
    if ($tagsRaw) { $tags = @($tagsRaw | ForEach-Object { $_.ToString() }) }

    $spNamesRaw = Get-GraphResponseProperty -Response $Source -PropertyName 'servicePrincipalNames'
    $spNames = @()
    if ($spNamesRaw) { $spNames = @($spNamesRaw | ForEach-Object { $_.ToString() }) }

    $verifiedPublisherName = $null
    $verifiedPublisherId = $null
    $verifiedPublisher = Get-GraphResponseProperty -Response $Source -PropertyName 'verifiedPublisher'
    if ($verifiedPublisher) {
        $verifiedPublisherName = Get-EffectiveString (Get-GraphResponseProperty -Response $verifiedPublisher -PropertyName 'displayName')
        $verifiedPublisherId   = Get-EffectiveString (Get-GraphResponseProperty -Response $verifiedPublisher -PropertyName 'verifiedPublisherId')
        if (-not $verifiedPublisherName -and $verifiedPublisher -is [System.Collections.IDictionary]) {
            foreach ($key in $verifiedPublisher.Keys) {
                if ($key -ieq 'displayName') {
                    $verifiedPublisherName = Get-EffectiveString $verifiedPublisher[$key]
                } elseif ($key -ieq 'verifiedPublisherId') {
                    $verifiedPublisherId = Get-EffectiveString $verifiedPublisher[$key]
                }
            }
        }
    }
    if (-not $verifiedPublisherName) {
        $verifiedPublisherName = Get-EffectiveString (Get-GraphResponseProperty -Response $Source -PropertyName 'verifiedPublisherName')
    }
    if (-not $verifiedPublisherName) {
        $verifiedPublisherName = Get-EffectiveString (Get-GraphResponseProperty -Response $Source -PropertyName 'VerifiedPublisherName')
    }
    if (-not $verifiedPublisherId) {
        $verifiedPublisherId = Get-EffectiveString (Get-GraphResponseProperty -Response $Source -PropertyName 'verifiedPublisherId')
    }
    if (-not $verifiedPublisherId) {
        $verifiedPublisherId = Get-EffectiveString (Get-GraphResponseProperty -Response $Source -PropertyName 'VerifiedPublisherId')
    }

    $profile = [ordered]@{
        Id                    = Get-GraphResponseProperty -Response $Source -PropertyName 'id'
        AppId                 = Get-GraphResponseProperty -Response $Source -PropertyName 'appId'
        Name                  = Get-GraphResponseProperty -Response $Source -PropertyName 'displayName'
        Tags                  = $tags
        Publisher             = Get-EffectiveString (Get-GraphResponseProperty -Response $Source -PropertyName 'publisherName')
        PublisherDomain       = Get-EffectiveString (Get-GraphResponseProperty -Response $Source -PropertyName 'publisherDomain')
        AppOwnerOrgId         = Get-EffectiveString (Get-GraphResponseProperty -Response $Source -PropertyName 'appOwnerOrganizationId')
        VerifiedPublisherName = $verifiedPublisherName
        VerifiedPublisherId   = $verifiedPublisherId
        ServicePrincipalNames = $spNames
        ServicePrincipalType  = Get-EffectiveString (Get-GraphResponseProperty -Response $Source -PropertyName 'servicePrincipalType')
    }

    $derivedDomain = Get-DomainFromServicePrincipalNames -Names $profile.ServicePrincipalNames
    if ($derivedDomain) { $profile.PublisherDomain = $derivedDomain }

    return $profile
}

function Merge-ServicePrincipalProfile {
    param(
        [hashtable]$Target,
        [pscustomobject]$Source
    )

    $profile = New-ServicePrincipalProfile -Source $Source
    if (-not $profile) { return }

    foreach ($key in 'Name','AppId','AppOwnerOrgId','ServicePrincipalType') {
        if ($profile[$key]) { $Target[$key] = $profile[$key] }
    }

    if ($profile.Tags -and @($profile.Tags).Count -gt 0) { $Target.Tags = $profile.Tags }
    if ($profile.ServicePrincipalNames -and @($profile.ServicePrincipalNames).Count -gt 0) { $Target.ServicePrincipalNames = $profile.ServicePrincipalNames }
    if ($profile.Publisher) { $Target.Publisher = $profile.Publisher }
    if ($profile.PublisherDomain) { $Target.PublisherDomain = $profile.PublisherDomain }
    if ($profile.VerifiedPublisherName) { $Target.VerifiedPublisherName = $profile.VerifiedPublisherName }
    if ($profile.VerifiedPublisherId) { $Target.VerifiedPublisherId = $profile.VerifiedPublisherId }
}

function Get-MicrosoftVendorInfo {
    param(
        [hashtable]$Profile,
        [string[]]$TenantDomains,
        [string]$TenantDisplayName
    )

    $vendorInfo = [ordered]@{
        IsMicrosoft                          = $false
        IsThirdParty                         = $false
        OwnerOrg                             = $Profile.AppOwnerOrgId
        OwnerOrgMatch                        = $false
        Publisher                            = $Profile.Publisher
        PublisherMatched                     = $false
        PublisherDomain                      = $Profile.PublisherDomain
        PublisherDomainMatchesMicrosoft      = $false
        PublisherDomainMatchesTenant         = $false
        HasIntegratedTag                     = $false
        HasGalleryTag                        = $false
        KnownAppIdMatch                      = $null
        Tags                                 = $Profile.Tags
        ServicePrincipalNames                = $Profile.ServicePrincipalNames
        ServicePrincipalNameMatchesMicrosoft = $false
        ServicePrincipalNameMatchesTenant    = $false
        ServicePrincipalType                 = $Profile.ServicePrincipalType
        VerifiedPublisherName                = $Profile.VerifiedPublisherName
        VerifiedPublisherId                  = $Profile.VerifiedPublisherId
        VerifiedPublisherMatchesMicrosoft    = $false
    }

    if ($Profile.AppOwnerOrgId) {
        $vendorInfo.OwnerOrgMatch = ($script:KnownMicrosoftTenantIds -contains $Profile.AppOwnerOrgId.ToLowerInvariant())
    }

    if ($Profile.Publisher) {
        $vendorInfo.PublisherMatched = ($Profile.Publisher.ToLowerInvariant() -match '\bmicrosoft\b')
    }

    if ($Profile.PublisherDomain) {
        $domain = $Profile.PublisherDomain.ToLowerInvariant()
        if ($TenantDomains) {
            foreach ($tenantDomain in $TenantDomains) {
                if ($domain -eq $tenantDomain -or $domain -like "*.$tenantDomain") {
                    $vendorInfo.PublisherDomainMatchesTenant = $true
                    break
                }
            }
        }
        foreach ($suffix in $script:MicrosoftDomainSuffixes) {
            if ($domain -eq $suffix -or $domain -like "*.$suffix") {
                $vendorInfo.PublisherDomainMatchesMicrosoft = $true
                break
            }
        }
    }

    if ($Profile.VerifiedPublisherName) {
        $vendorInfo.VerifiedPublisherMatchesMicrosoft = ($Profile.VerifiedPublisherName.ToLowerInvariant() -match '\bmicrosoft\b')
    }
    if (-not $vendorInfo.VerifiedPublisherMatchesMicrosoft -and $Profile.VerifiedPublisherId) {
        $vendorInfo.VerifiedPublisherMatchesMicrosoft = ($script:KnownMicrosoftTenantIds -contains $Profile.VerifiedPublisherId.ToLowerInvariant())
    }

    if ($Profile.ServicePrincipalNames -and @($Profile.ServicePrincipalNames).Count -gt 0) {
        foreach ($name in $Profile.ServicePrincipalNames) {
            if (-not $name) { continue }
            $lower = $name.ToLowerInvariant()
            if ($TenantDomains) {
                foreach ($tenantDomain in $TenantDomains) {
                    if ($lower -like "*$tenantDomain*") {
                        $vendorInfo.ServicePrincipalNameMatchesTenant = $true
                        break
                    }
                }
            }
            if ($lower -match 'microsoft' -or $lower -match 'windows\.net' -or $lower -match 'azure\.com' -or $lower -match 'office\.com') {
                $vendorInfo.ServicePrincipalNameMatchesMicrosoft = $true
            } elseif ($script:MicrosoftDomainSuffixes) {
                foreach ($suffix in $script:MicrosoftDomainSuffixes) {
                    $pattern = [regex]::Escape($suffix)
                    if ($lower -match $pattern) {
                        $vendorInfo.ServicePrincipalNameMatchesMicrosoft = $true
                        break
                    }
                }
            }
            if ($vendorInfo.ServicePrincipalNameMatchesMicrosoft -and $vendorInfo.ServicePrincipalNameMatchesTenant) { break }
        }
    }

    $tags = $Profile.Tags
    if ($tags) {
        $vendorInfo.HasIntegratedTag = $tags -contains 'WindowsAzureActiveDirectoryIntegratedApp'
        $vendorInfo.HasGalleryTag = ($tags -contains 'WindowsAzureActiveDirectoryGalleryApplicationPrimaryV1' -or $tags -contains 'WindowsAzureActiveDirectoryGalleryApplicationNonPrimaryV1')
    }

    if ($Profile.AppId -and ($script:KnownMicrosoftAppIds -contains $Profile.AppId)) {
        $vendorInfo.KnownAppIdMatch = $Profile.AppId
    }

    $tagIndicatesMicrosoft = $vendorInfo.HasIntegratedTag -or $vendorInfo.HasGalleryTag
    $publisherAvailable = [bool]$Profile.Publisher

    $vendorInfo.IsMicrosoft = $vendorInfo.OwnerOrgMatch -or
                              $vendorInfo.VerifiedPublisherMatchesMicrosoft -or
                              $vendorInfo.PublisherDomainMatchesMicrosoft -or
                              $vendorInfo.ServicePrincipalNameMatchesMicrosoft -or
                              ($vendorInfo.PublisherMatched -and ($tagIndicatesMicrosoft -or -not $tagIndicatesMicrosoft -and -not $publisherAvailable)) -or
                              (-not $publisherAvailable -and $tagIndicatesMicrosoft) -or
                              ($vendorInfo.KnownAppIdMatch -ne $null)

    if (-not $vendorInfo.IsMicrosoft) {
        $isTenantPublisher = $vendorInfo.PublisherDomainMatchesTenant -or
                              $vendorInfo.ServicePrincipalNameMatchesTenant -or
                              ($Profile.Publisher -and $TenantDisplayName -and ($Profile.Publisher -eq $TenantDisplayName))

        if ($Profile.VerifiedPublisherName) {
            $vendorInfo.IsThirdParty = (-not $vendorInfo.VerifiedPublisherMatchesMicrosoft -and -not $isTenantPublisher)
        } elseif ($Profile.PublisherDomain) {
            $vendorInfo.IsThirdParty = (-not $vendorInfo.PublisherDomainMatchesMicrosoft -and -not $vendorInfo.PublisherDomainMatchesTenant -and -not $isTenantPublisher)
        } elseif ($Profile.Publisher) {
            $vendorInfo.IsThirdParty = (-not $vendorInfo.PublisherMatched -and -not $isTenantPublisher)
        }
    }

    return [pscustomobject]$vendorInfo
}

$criticalPatterns = @(
    '.*ReadWrite.*',
    '^Directory\.AccessAsUser\.All$',
    '^Mail\.ReadWrite(\.All)?$',
    '^Calendars\.ReadWrite(\.All)?$',
    '^Files\.ReadWrite(\.All)?$',
    '^Contacts\.ReadWrite(\.All)?$',
    '^MailboxSettings\.ReadWrite$',
    '^SecurityEvents\.ReadWrite\.All$',
    '^PrivilegedAccess.*$',
    '^Sites\.FullControl\.All$'
)
$lowImpactExact = @('User.Read','openid','profile','email','offline_access')

function Test-CriticalScope {
    param([string]$Scope)
    foreach ($pattern in $criticalPatterns) {
        if ($Scope -match $pattern) { return $true }
    }
    return $false
}

function Is-LowImpactScopeSet {
    param([string[]]$Scopes)
    if (-not $Scopes -or @($Scopes).Count -eq 0) { return $false }
    foreach ($scope in $Scopes) {
        if ($lowImpactExact -notcontains $scope) { return $false }
    }
    return $true
}

Initialize-GraphContext
$tenantMetadata = Get-TenantMetadata
$tenantDomains = $tenantMetadata.Domains
$tenantDisplayName = $tenantMetadata.DisplayName

$grantUri = 'https://graph.microsoft.com/beta/oauth2PermissionGrants?$count=true&$top=999'
$grants = Invoke-GraphPaged -Uri $grantUri
if (-not $grants) {
    Write-Host 'No oauth2PermissionGrants found.'
    return
}

$clientIds = [System.Collections.Generic.HashSet[string]]::new()
$resourceIds = [System.Collections.Generic.HashSet[string]]::new()
$userIds = [System.Collections.Generic.HashSet[string]]::new()
foreach ($grant in $grants) {
    if ($grant.clientId) { [void]$clientIds.Add($grant.clientId) }
    if ($grant.resourceId) { [void]$resourceIds.Add($grant.resourceId) }
    if ($grant.consentType -eq 'Principal' -and $grant.principalId) { [void]$userIds.Add($grant.principalId) }
}

$directoryObjects = Get-DirectoryObjectsByIds -Ids @($clientIds + $resourceIds + $userIds)
$servicePrincipals = @{}
$users = @{}

foreach ($obj in $directoryObjects) {
    switch ($obj.'@odata.type') {
        '#microsoft.graph.servicePrincipal' {
            $profile = New-ServicePrincipalProfile -Source $obj
            if ($profile) { $servicePrincipals[$obj.id] = $profile }
        }
        '#microsoft.graph.user' {
            $users[$obj.id] = [pscustomobject]@{
                Id   = $obj.id
                UPN  = $obj.userPrincipalName
                Name = $obj.displayName
            }
        }
    }
}

$allServicePrincipalIds = (@($clientIds) + @($resourceIds)) | Where-Object { $_ } | Sort-Object -Unique
$metadataById = @{}
if ($allServicePrincipalIds) {
    try {
        $metadataById = Get-ServicePrincipalMetadata -Ids $allServicePrincipalIds
    } catch {
        Write-Warning ("Unable to fetch extended service principal metadata: {0}" -f $_.Exception.Message)
        $metadataById = @{}
    }
}

foreach ($id in $metadataById.Keys) {
    if (-not $servicePrincipals.ContainsKey($id)) {
        $profile = New-ServicePrincipalProfile -Source $metadataById[$id]
        if ($profile) { $servicePrincipals[$id] = $profile }
        continue
    }
    Merge-ServicePrincipalProfile -Target $servicePrincipals[$id] -Source $metadataById[$id]
}

if ($AppNameContains) {
    $grants = $grants | Where-Object {
        $client = $servicePrincipals[$_.clientId]
        $client -and $client.Name -and ($client.Name -like "*$AppNameContains*")
    }
}

if (-not $grants) {
    Write-Host 'No grants found after filtering.'
    return
}

$expandedGrants = foreach ($grant in $grants) {
    $client = $servicePrincipals[$grant.clientId]
    $resource = $servicePrincipals[$grant.resourceId]
    $scopes = ($grant.scope -split '\s+') | Where-Object { $_ }
    foreach ($scope in $scopes) {
        [pscustomobject]@{
            GrantId       = $grant.id
            ConsentType   = $grant.consentType
            ClientId      = $grant.clientId
            ClientName    = $client.Name
            ClientAppId   = $client.AppId
            ResourceId    = $grant.resourceId
            ResourceName  = $resource.Name
            ResourceAppId = $resource.AppId
            Scope         = $scope
            PrincipalId   = if ($grant.consentType -eq 'Principal') { $grant.principalId }
        }
    }
}

$recordsByClient = $expandedGrants | Group-Object ClientId
$appSummaries = @()
foreach ($group in $recordsByClient) {
    $recordsForClient = $group.Group
    $clientProfile = $servicePrincipals[$group.Name]
    $allScopes = @(($recordsForClient | Select-Object -Expand Scope) | Sort-Object -Unique)

    $isCritical = $false
    foreach ($scope in $allScopes) {
        if (Test-CriticalScope -Scope $scope) { $isCritical = $true; break }
    }
    $status = if ($isCritical) { 'fail' } elseif (Is-LowImpactScopeSet -Scopes $allScopes) { 'pass' } else { 'warn' }

    $administratorGrants = @()
    $userGrants = @()

    foreach ($adminGroup in (($recordsForClient | Where-Object { $_.ConsentType -eq 'AllPrincipals' }) | Group-Object ResourceName, Scope)) {
        $sample = $adminGroup.Group[0]
        $administratorGrants += [pscustomobject]@{
            resource = $sample.ResourceName
            scope    = $sample.Scope
            type     = 'Admin'
        }
    }

    foreach ($userGroup in (($recordsForClient | Where-Object { $_.ConsentType -eq 'Principal' }) | Group-Object ResourceName, Scope)) {
        $sample = $userGroup.Group[0]
        $userUpns = $userGroup.Group | Where-Object { $_.PrincipalId } | ForEach-Object {
            if ($users.ContainsKey($_.PrincipalId)) { $users[$_.PrincipalId].UPN }
        } | Where-Object { $_ } | Sort-Object -Unique
        $userUpns = @($userUpns)
        $userGrants += [pscustomobject]@{
            resource = $sample.ResourceName
            scope    = $sample.Scope
            type     = 'User'
            users    = @($userUpns)
            userCount= @($userUpns).Count
        }
    }

    $vendorInfo = Get-MicrosoftVendorInfo -Profile $clientProfile -TenantDomains $tenantDomains -TenantDisplayName $tenantDisplayName
    $isHiddenApp = Test-ServicePrincipalHidden -Tags $clientProfile.Tags
    $userScopeCount = @($userGrants | Select-Object -ExpandProperty scope -Unique).Count
    $userCount = @($userGrants | ForEach-Object { $_.users } | Where-Object { $_ } | Select-Object -Unique).Count

    $tags = @()
    if ($vendorInfo.IsMicrosoft) { $tags += 'Microsoft' }
    elseif ($vendorInfo.IsThirdParty) { $tags += '3rd Party' }
    if ($isHiddenApp) { $tags += 'Hidden' }
    if ($administratorGrants) { $tags += 'Admin Consent' }
    if ($userGrants) { $tags += 'User Consent' }
    $tags += "user-perm: $userScopeCount"
    $tags += "user-count: $userCount"

    $appSummaries += [pscustomobject]@{
        clientId      = $group.Name
        clientName    = $clientProfile.Name
        clientAppId   = $clientProfile.AppId
        status        = $status
        tags          = $tags
        isMicrosoft   = $vendorInfo.IsMicrosoft
        isThirdParty  = $vendorInfo.IsThirdParty
        isHidden      = $isHiddenApp
        vendorInfo    = $vendorInfo
        adminGrants   = $administratorGrants
        userGrants    = $userGrants
        allScopes     = $allScopes
    }
}

$summariesByClientId = @{}
foreach ($summary in $appSummaries) {
    if ($summary.clientId) { $summariesByClientId[$summary.clientId] = $summary }
}

$allServicePrincipals = @()
try {
    $allServicePrincipals = Get-AllServicePrincipals
} catch {
    Write-Warning ("Unable to retrieve complete service principal list: {0}" -f $_.Exception.Message)
    $allServicePrincipals = @()
}

foreach ($spRaw in $allServicePrincipals) {
    $spId = Get-GraphResponseProperty -Response $spRaw -PropertyName 'id'
    if (-not $spId) { continue }

    $profile = $null
    if ($servicePrincipals.ContainsKey($spId)) {
        $profile = $servicePrincipals[$spId]
        Merge-ServicePrincipalProfile -Target $profile -Source $spRaw
    } else {
        $profile = New-ServicePrincipalProfile -Source $spRaw
        if ($profile) { $servicePrincipals[$spId] = $profile }
    }
    if (-not $profile) { continue }

    $vendorInfo = Get-MicrosoftVendorInfo -Profile $profile -TenantDomains $tenantDomains -TenantDisplayName $tenantDisplayName
    $isHiddenSp = Test-ServicePrincipalHidden -Tags $profile.Tags

    if ($summariesByClientId.ContainsKey($spId)) {
        $summary = $summariesByClientId[$spId]
        $summary.vendorInfo = $vendorInfo
        $summary.isMicrosoft = $vendorInfo.IsMicrosoft
        $summary.isThirdParty = $vendorInfo.IsThirdParty
        $summary.isHidden = $isHiddenSp

        if ($vendorInfo.IsMicrosoft) {
            if (-not ($summary.tags -contains 'Microsoft')) { $summary.tags += 'Microsoft' }
            $summary.tags = @($summary.tags | Where-Object { $_ -ne '3rd Party' })
        } elseif ($vendorInfo.IsThirdParty) {
            if (-not ($summary.tags -contains '3rd Party')) { $summary.tags += '3rd Party' }
            $summary.tags = @($summary.tags | Where-Object { $_ -ne 'Microsoft' })
        } else {
            $summary.tags = @($summary.tags | Where-Object { ($_ -ne 'Microsoft') -and ($_ -ne '3rd Party') })
        }

        if ($isHiddenSp) {
            if (-not ($summary.tags -contains 'Hidden')) { $summary.tags += 'Hidden' }
        } else {
            $summary.tags = @($summary.tags | Where-Object { $_ -ne 'Hidden' })
        }

        $hasGrants = (@($summary.adminGrants).Count -gt 0) -or (@($summary.userGrants).Count -gt 0)
        if ($hasGrants) {
            $summary.tags = @($summary.tags | Where-Object { $_ -ne 'No OAuth grants' })
        } elseif (-not ($summary.tags -contains 'No OAuth grants')) {
            $summary.tags += 'No OAuth grants'
        }
    } else {
        $nameFallback = if ($profile.Name) { $profile.Name } elseif ($profile.AppId) { $profile.AppId } else { $spId }
        $tags = @()
        if ($vendorInfo.IsMicrosoft) { $tags += 'Microsoft' }
        elseif ($vendorInfo.IsThirdParty) { $tags += '3rd Party' }
        if ($isHiddenSp) { $tags += 'Hidden' }
        $tags += 'No OAuth grants'

        $newSummary = [pscustomobject]@{
            clientId     = $spId
            clientName   = $nameFallback
            clientAppId  = $profile.AppId
            status       = 'pass'
            tags         = $tags
            isMicrosoft  = $vendorInfo.IsMicrosoft
            isThirdParty = $vendorInfo.IsThirdParty
            isHidden     = $isHiddenSp
            vendorInfo   = $vendorInfo
            adminGrants  = @()
            userGrants   = @()
            allScopes    = @()
        }
        $appSummaries += $newSummary
        $summariesByClientId[$spId] = $newSummary
    }
}

$totalApps = @($appSummaries).Count
$failCount = @($appSummaries | Where-Object status -eq 'fail').Count
$passCount = @($appSummaries | Where-Object status -eq 'pass').Count
$warnCount = @($appSummaries | Where-Object status -eq 'warn').Count
$microsoftCount = @($appSummaries | Where-Object isMicrosoft).Count
$thirdPartyCount = @($appSummaries | Where-Object isThirdParty).Count

$outDir = (Get-Location).Path
$timestamp = Get-Date -Format 'yyyyMMdd_HHmm'
$csvPath = Join-Path $outDir ("OAuth2_Consent_Apps_$timestamp.csv")

$appSummaries |
    Select-Object clientName,
                  clientAppId,
                  status,
                  @{n='isMicrosoft';e={$_.isMicrosoft}},
                  @{n='isThirdParty';e={$_.isThirdParty}},
                  @{n='isHidden';e={$_.isHidden}},
                  @{n='publisher';e={$_.vendorInfo.Publisher}},
                  @{n='publisherDomain';e={$_.vendorInfo.PublisherDomain}},
                  @{n='publisherDomainMatchesMicrosoft';e={$_.vendorInfo.PublisherDomainMatchesMicrosoft}},
                  @{n='publisherDomainMatchesTenant';e={$_.vendorInfo.PublisherDomainMatchesTenant}},
                  @{n='verifiedPublisherName';e={$_.vendorInfo.VerifiedPublisherName}},
                  @{n='verifiedPublisherId';e={$_.vendorInfo.VerifiedPublisherId}},
                  @{n='appOwnerOrgId';e={$_.vendorInfo.OwnerOrg}},
                  @{n='msOwnerOrgMatch';e={$_.vendorInfo.OwnerOrgMatch}},
                  @{n='msPublisherMatch';e={$_.vendorInfo.PublisherMatched}},
                  @{n='verifiedPublisherMatchesMicrosoft';e={$_.vendorInfo.VerifiedPublisherMatchesMicrosoft}},
                  @{n='msHasIntegratedTag';e={$_.vendorInfo.HasIntegratedTag}},
                  @{n='msHasGalleryTag';e={$_.vendorInfo.HasGalleryTag}},
                  @{n='msKnownAppId';e={$_.vendorInfo.KnownAppIdMatch}},
                  @{n='servicePrincipalType';e={$_.vendorInfo.ServicePrincipalType}},
                  @{n='servicePrincipalNames';e={($_.vendorInfo.ServicePrincipalNames) -join ';'}},
                  @{n='servicePrincipalNameMatchesMicrosoft';e={$_.vendorInfo.ServicePrincipalNameMatchesMicrosoft}},
                  @{n='servicePrincipalNameMatchesTenant';e={$_.vendorInfo.ServicePrincipalNameMatchesTenant}},
                  @{n='servicePrincipalTags';e={($_.vendorInfo.Tags) -join ';'}},
                  @{n='tags';e={$_.tags -join ';'}},
                  @{n='adminScopes';e={($_.adminGrants | ForEach-Object { "{0}:{1}" -f $_.resource, $_.scope }) -join ';'}},
                  @{n='userScopes'; e={($_.userGrants  | ForEach-Object { "{0}:{1}" -f $_.resource, $_.scope }) -join ';'}},
                  @{n='userCount';  e={@($_.userGrants  | ForEach-Object { $_.users } | Where-Object { $_ } | Select-Object -Unique).Count}} |
    Export-Csv -NoTypeInformation -Encoding UTF8 -Path $csvPath

$report = [pscustomobject]@{
    generatedAt = (Get-Date).ToString('s')
    totals      = @{ total=$totalApps; pass=$passCount; warn=$warnCount; fail=$failCount; microsoft=$microsoftCount; thirdParty=$thirdPartyCount }
    items       = $appSummaries
}

$json = $report | ConvertTo-Json -Depth 12
$htmlTemplate = @'
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>OAuth2 Consent Report</title>
<style>
  :root{
    --brand:#7ad12b;
    --brand-2:#00c2a8;
    --ok:#7ad12b;
    --warn:#ffd166;
    --fail:#ff5a5f;
    --muted:#9aa3b2;
    --bg:#0a0f14;
    --card:#0f151c;
    --text:#e9f0f3;
    --chip:#111b22;
    --line:#1e2a33;
    --focus-ring:0 0 0 3px rgba(0,194,168,0.35);
    --microsoft:#4da3ff;
    --thirdparty:#ffb147;
  }
  html,body{margin:0;height:100%;background:var(--bg);color:var(--text);font:14px/1.45 system-ui,Segoe UI,Roboto,Arial,sans-serif}
  a{color:var(--brand)}
  header{position:sticky;top:0;background:linear-gradient(180deg,#0a0f14 60%,#0a0f14cc 100%);backdrop-filter:blur(6px);z-index:10;border-bottom:1px solid var(--line)}
  .wrap{max-width:1200px;margin:0 auto;padding:14px 16px}
  .row{display:flex;gap:8px;flex-wrap:wrap;align-items:center}
  input[type="search"],button{background:#0c1319;border:1px solid #1a2a33;color:var(--text);padding:8px 10px;border-radius:8px}
  input[type="search"]::placeholder{color:#7b8794}
  button{cursor:pointer}
  .chip{display:inline-flex;gap:8px;align-items:center;background:var(--chip);padding:6px 10px;border-radius:20px;border:1px solid #1a2a33}
  .chip b{font-weight:600}
  .chip.clickable{cursor:pointer}
  .chip.active{outline:2px solid var(--brand-2);box-shadow:0 0 0 2px #081118 inset}
  .status{width:10px;height:10px;border-radius:999px}
  .status.pass{background:var(--ok)} .status.warn{background:var(--warn)} .status.fail{background:var(--fail)}
    /* Center the list */
  #list{max-width:1100px;margin:14px auto;display:grid;grid-template-columns:1fr;gap:10px}
  .card{background:var(--card);border:1px solid var(--line);border-radius:12px;overflow:hidden}
  .item{display:grid;grid-template-columns:28px 1fr auto;gap:10px;align-items:center;padding:12px;border-bottom:1px solid var(--line)}
  h4{margin:0;font-size:15px}
  .meta{color:#7b8794;font-size:12px}
  .tags{display:flex;gap:6px;flex-wrap:wrap;margin-top:6px}
  .tag{background:#0c1319;border:1px solid #1a2a33;padding:2px 8px;border-radius:999px;font-size:12px;transition:background-color .18s ease,border-color .18s ease,color .18s ease,box-shadow .18s ease,transform .18s ease}
  .tag-admin{background:rgba(122,209,43,0.12);border-color:#2f4;color:#bdf8a7}
  .tag-user{background:rgba(0,194,168,0.12);border-color:#0bd;color:#a6fff3}
    .tag-microsoft{background:rgba(77,163,255,0.14);border-color:rgba(77,163,255,0.7);color:#bcdcff}
    .tag-thirdparty{background:rgba(255,177,71,0.12);border-color:rgba(255,177,71,0.7);color:#ffd7a2}
    .tag-hidden{background:rgba(148,163,184,0.12);border-color:rgba(148,163,184,0.6);color:#d8dee9;font-style:italic}
  .tag-clickable{cursor:pointer}
  .tag-clickable:focus-visible{outline:none;box-shadow:var(--focus-ring)}
  .tag-clickable:hover{transform:translateY(-1px)}
  .tag-admin.tag-clickable:hover{background:rgba(122,209,43,0.22);border-color:#4dff6b;color:#e3ffd4}
  .tag-user.tag-clickable:hover{background:rgba(0,194,168,0.22);border-color:#34f1ff;color:#e2fffb}
  .tag-microsoft.tag-clickable:hover{background:rgba(77,163,255,0.25);border-color:#86c5ff;color:#e8f3ff}
    .tag-thirdparty.tag-clickable:hover{background:rgba(255,177,71,0.22);border-color:rgba(255,214,153,0.9);color:#ffe7c4}
  .tag.tag-active{box-shadow:0 0 0 2px rgba(0,194,168,0.35);border-color:#34f1ff}
  .tag-microsoft.tag-active{box-shadow:0 0 0 2px rgba(77,163,255,0.45);border-color:#bcdcff}
  .meta-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(220px,1fr));gap:8px 18px;margin:0 0 14px}
  .meta-label{display:block;font-size:11px;letter-spacing:.06em;text-transform:uppercase;color:#8093a6;margin-bottom:2px}
  .meta-value{font-size:13px;color:#d5dee5}
  .actions{display:flex;gap:6px}
  details{background:#0b1218}
  details[open] summary{border-bottom:1px dashed var(--line)}
  summary{list-style:none;padding:10px 12px;cursor:pointer;color:#cbd5e1}
  summary::-webkit-details-marker{display:none}
  .sec-title{font-size:12px;text-transform:uppercase;letter-spacing:.08em;color:#9aa3b2;margin:8px 0}
  .divider{height:1px;background:linear-gradient(90deg,transparent,#24424f 30%,#24424f 70%,transparent);margin:10px 0}
  .consent-row{padding:6px 0}
  .pill{border-radius:999px;padding:2px 8px;border:1px solid #1a2a33;background:#0c1319;font-size:12px}
  .footer{padding:20px;color:#9aa3b2;text-align:center;border-top:1px solid var(--line)}
    .status.microsoft{background:var(--microsoft)}
    .status.thirdparty{background:var(--thirdparty)}
  .list-counter{max-width:1100px;margin:12px auto 0;padding:0 16px;color:#9aa3b2;font-size:13px}
  .list-counter b{color:var(--text)}
  .footer .sig{display:inline-flex;gap:8px;align-items:center}
  .icon{width:16px;height:16px;vertical-align:-3px;fill:#0e76a8}
</style>
</head>
<body>
<header>
  <div class="wrap row">
    <strong style="font-size:16px">Entra ID – OAuth2 Consent Report</strong>

        <!-- Filter chips -->
    <span id="chipTotal" class="chip clickable" role="button" tabindex="0" aria-pressed="false"><b id="tTotal">0</b> Enterprise apps</span>
  <span id="chipFail"  class="chip clickable" role="button" tabindex="0" aria-pressed="false"><span class="status fail"></span><b id="tFail">0</b> Critical scopes</span>
  <span id="chipWarn"  class="chip clickable" role="button" tabindex="0" aria-pressed="false"><span class="status warn"></span><b id="tWarn">0</b> Elevated scopes</span>
    <span id="chipPass"  class="chip clickable" role="button" tabindex="0" aria-pressed="false"><span class="status pass"></span><b id="tPass">0</b> Low impact only</span>
    <span id="chipMicrosoft" class="chip clickable" role="button" tabindex="0" aria-pressed="false"><span class="status microsoft"></span><b id="tMicrosoft">0</b> Microsoft apps</span>
    <span id="chipThirdParty" class="chip clickable" role="button" tabindex="0" aria-pressed="false"><span class="status thirdparty"></span><b id="tThirdParty">0</b> 3rd party apps</span>

    <div style="flex:1"></div>
    <input id="q" type="search" placeholder="Search enterprise app" />
    <button id="sortToggle" type="button" aria-pressed="false" title="Sort enterprise apps alphabetically">Sort A→Z</button>
    <button id="exportCsv" title="Export filtered apps to CSV">Export CSV</button>
  </div>
</header>
<div class="wrap list-counter" aria-live="polite">Currently visible: <b id="visibleCount">0</b> enterprise apps</div>
<div class="wrap" id="list"></div>
<div class="footer">
  <span class="sig">
    created by <strong>Jürgen Waldl</strong>
    <a href="https://www.linkedin.com/in/jürgen-waldl-6592837b/" target="_blank" rel="noopener" title="LinkedIn – Jürgen Waldl">
      <svg class="icon" viewBox="0 0 24 24" aria-hidden="true">
        <path d="M4.98 3.5C4.98 4.88 3.86 6 2.5 6S0 4.88 0 3.5 1.12 1 2.5 1 4.98 2.12 4.98 3.5zM0 8h5v16H0zM8 8h4.8v2.2h.07c.67-1.2 2.3-2.46 4.73-2.46 5.06 0 6 3.33 6 7.66V24h-5V16.4c0-1.8-.03-4.1-2.5-4.1-2.5 0-2.88 1.95-2.88 3.97V24H8z"/>
      </svg>
    </a>
  </span>
  <div>generated @ <span id="genAt"></span></div>
</div>
<script>
const report = __JSON__;

// --- State & Elements ---
const state = { q:"", status:"all", consent:"any", vendor:"any", sort:"asc" };
const $q = document.getElementById("q");
const $list = document.getElementById("list");

// Chips
const $chipTotal = document.getElementById("chipTotal");
const $chipFail  = document.getElementById("chipFail");
const $chipWarn  = document.getElementById("chipWarn");
const $chipPass  = document.getElementById("chipPass");
const $chipMicrosoft = document.getElementById("chipMicrosoft");
const $chipThirdParty = document.getElementById("chipThirdParty");
const $visibleCount = document.getElementById("visibleCount");
const $sortToggle = document.getElementById("sortToggle");

// Header totals
document.getElementById("tTotal").textContent = report.totals.total;
document.getElementById("tPass").textContent  = report.totals.pass;
document.getElementById("tWarn").textContent  = report.totals.warn;
document.getElementById("tFail").textContent  = report.totals.fail;
document.getElementById("tMicrosoft").textContent = report.totals.microsoft;
document.getElementById("tThirdParty").textContent = report.totals.thirdParty;
document.getElementById("genAt").textContent  = report.generatedAt;

// Search input
$q.addEventListener("input", ()=>{ state.q = $q.value.trim().toLowerCase(); render(); });
document.addEventListener("keydown", e=>{ if(e.key==="/" && (e.ctrlKey||e.metaKey)){ e.preventDefault(); $q.focus(); }});

function updateSortButton(){
    const asc = state.sort === "asc";
    $sortToggle.textContent = asc ? "Sort A→Z" : "Sort Z→A";
    $sortToggle.setAttribute("aria-pressed", String(!asc));
    $sortToggle.title = asc ? "Sort descending (Z→A)" : "Sort ascending (A→Z)";
}
$sortToggle.addEventListener("click", ()=>{
    state.sort = state.sort === "asc" ? "desc" : "asc";
    updateSortButton();
    render();
});
updateSortButton();

// Filter chip UI
function setStatusFilter(filter){
  state.status = filter;
  updateStatusChips();
  render();
}
function updateStatusChips(){
  const mapping = [
    [$chipTotal, "all"],
    [$chipFail, "fail"],
    [$chipWarn, "warn"],
    [$chipPass, "pass"]
  ];
  mapping.forEach(([chip, value])=>{
    const active = state.status === value;
    chip.classList.toggle("active", active);
    chip.setAttribute("aria-pressed", String(active));
  });
}
function toggleVendorFilter(vendor){
    state.vendor = (state.vendor === vendor) ? "any" : vendor;
    updateVendorChips();
    render();
}
function updateVendorChips(){
    const vendorMapping = [
        [$chipMicrosoft, "microsoft"],
        [$chipThirdParty, "thirdparty"]
    ];
    vendorMapping.forEach(([chip, value])=>{
        const active = state.vendor === value;
        chip.classList.toggle("active", active);
        chip.setAttribute("aria-pressed", String(active));
    });
}
$chipTotal.addEventListener("click", ()=> setStatusFilter("all"));
$chipFail .addEventListener("click", ()=> setStatusFilter("fail"));
$chipWarn .addEventListener("click", ()=> setStatusFilter("warn"));
$chipPass .addEventListener("click", ()=> setStatusFilter("pass"));
$chipMicrosoft.addEventListener("click", ()=> toggleVendorFilter("microsoft"));
$chipThirdParty.addEventListener("click", ()=> toggleVendorFilter("thirdparty"));
[$chipTotal,$chipFail,$chipWarn,$chipPass,$chipMicrosoft,$chipThirdParty].forEach(chip=>{
  chip.addEventListener("keydown", event=>{
    if(event.key === "Enter" || event.key === " "){
      event.preventDefault();
      chip.click();
    }
  });
});
updateStatusChips();
updateVendorChips();

function toggleConsentFilter(consent){
  state.consent = (state.consent === consent) ? "any" : consent;
  render();
}

$list.addEventListener("click", (event)=>{
  const tag = event.target.closest('.tag[data-consent], .tag[data-vendor]');
  if(!tag) return;
  const consent = tag.getAttribute('data-consent');
  const vendor = tag.getAttribute('data-vendor');
  if(consent){
    toggleConsentFilter(consent);
  } else if(vendor){
    toggleVendorFilter(vendor);
  }
});

$list.addEventListener("keydown", (event)=>{
  const tag = event.target.closest('.tag[data-consent], .tag[data-vendor]');
  if(!tag) return;
  if(event.key === "Enter" || event.key === " "){
    event.preventDefault();
    const consent = tag.getAttribute('data-consent');
    const vendor = tag.getAttribute('data-vendor');
    if(consent){
      toggleConsentFilter(consent);
    } else if(vendor){
      toggleVendorFilter(vendor);
    }
  }
});

// Utils
function escapeHtml(s){ return String(s).replace(/[&<>"']/g,m=>({ "&":"&amp;","<":"&lt;",">":"&gt;","\"":"&quot;","'":"&#039;"}[m])) }
function asArray(x){ return Array.isArray(x) ? x : (x ? [x] : []); }
function tagHtml(t){
  const label = escapeHtml(t);
  const lower = String(t).toLowerCase();
  const isAdmin = lower === "admin consent";
  const isUser  = lower === "user consent";
  const isMicrosoft = lower === "microsoft";
  const isThirdParty = lower === "3rd party";
  const isHidden = lower === "hidden";
  if (isAdmin || isUser){
    const consent = isAdmin ? "admin" : "user";
    const classes = ["tag", isAdmin ? "tag-admin" : "tag-user", "tag-clickable"];
    const active = state.consent === consent;
    if(active){ classes.push("tag-active"); }
    return `<span class="${classes.join(" ")}" data-consent="${consent}" role="button" tabindex="0" aria-pressed="${active}">${label}</span>`;
  }
  if (isMicrosoft){
    const active = state.vendor === "microsoft";
    const classes = ["tag", "tag-microsoft", "tag-clickable"];
    if(active){ classes.push("tag-active"); }
    return `<span class="${classes.join(" ")}" data-vendor="microsoft" role="button" tabindex="0" aria-pressed="${active}">${label}</span>`;
  }
  if (isThirdParty){
    const active = state.vendor === "thirdparty";
    const classes = ["tag", "tag-thirdparty", "tag-clickable"];
    if(active){ classes.push("tag-active"); }
    return `<span class="${classes.join(" ")}" data-vendor="thirdparty" role="button" tabindex="0" aria-pressed="${active}">${label}</span>`;
  }
  if (isHidden){
    return `<span class="tag tag-hidden">${label}</span>`;
  }
  return `<span class="tag">${label}</span>`;
}
function filterItems(items){
  let out = items;
  if(state.status !== "all"){ out = out.filter(x => (x.status||"") === state.status); }
  if(state.consent === "admin"){ out = out.filter(x => (x.adminGrants||[]).length > 0); }
  if(state.consent === "user"){ out = out.filter(x => (x.userGrants||[]).length > 0); }
  if(state.vendor === "microsoft"){ out = out.filter(x => x.isMicrosoft); }
  if(state.vendor === "thirdparty"){ out = out.filter(x => x.isThirdParty); }
  if(state.q){
    const q = state.q;
    out = out.filter(x => (x.clientName||"").toLowerCase().includes(q));
  }
  return out;
}

function sortItems(items){
    return [...items].sort((a, b) => {
        const nameA = (a.clientName || "").toLocaleLowerCase();
        const nameB = (b.clientName || "").toLocaleLowerCase();
        if (nameA !== nameB) {
            const order = nameA.localeCompare(nameB);
            return state.sort === "asc" ? order : -order;
        }
        const idA = (a.clientAppId || "").toLocaleLowerCase();
        const idB = (b.clientAppId || "").toLocaleLowerCase();
        const order = idA.localeCompare(idB);
        return state.sort === "asc" ? order : -order;
    });
}

// Render
function render(){
    const items = sortItems(filterItems(report.items));
  $visibleCount.textContent = items.length;
  $list.innerHTML = items.map(item => {
    const tags = (item.tags||[]).map(tag => tagHtml(tag)).join("");
    const adminRows = (item.adminGrants||[]).map(a=>`<div class="consent-row"><span class="pill">Admin</span> ${escapeHtml(a.resource)} • <b>${escapeHtml(a.scope)}</b></div>`).join("");
    const userRows  = (item.userGrants ||[]).map(u=>{
      const arr = asArray(u.users);
      return `<div class="consent-row"><span class="pill">User</span> ${escapeHtml(u.resource)} • <b>${escapeHtml(u.scope)}</b> • users: ${escapeHtml(arr.join(", "))}</div>`;
    }).join("");
    const hasAdmin = !!adminRows, hasUser = !!userRows;
    const sections = [
      hasAdmin ? `<div class="sec-title">Admin Consents</div>${adminRows}` : "",
      (hasAdmin && hasUser) ? `<div class="divider"></div>` : "",
      hasUser  ? `<div class="sec-title">User Consents</div>${userRows}`   : ""
    ].join("");
    const none = (!hasAdmin && !hasUser) ? "<div class='consent-row' style='color:#7b8794'>No delegated grants found.</div>" : "";
    const scopeArr = asArray(item.allScopes);
    const info = item.vendorInfo || {};
    const vendorClass = info.IsMicrosoft ? "Microsoft" : (info.IsThirdParty ? "3rd Party" : "Custom / unknown");
    const vendorClassEsc = escapeHtml(vendorClass);
    const verifiedPublisherRaw = info.VerifiedPublisherName || info.Publisher || "—";
    const verifiedPublisherEsc = escapeHtml(verifiedPublisherRaw);
    const ownerOrgEsc = info.OwnerOrg ? escapeHtml(info.OwnerOrg) : "—";
    const spTagsEsc = (info.Tags && info.Tags.length) ? escapeHtml(info.Tags.join(", ")) : "—";
    const publisherDomainEsc = info.PublisherDomain ? escapeHtml(info.PublisherDomain) : "—";
    const spNames = Array.isArray(info.ServicePrincipalNames) ? info.ServicePrincipalNames : [];
    const spNameDisplay = spNames.length ? escapeHtml(spNames.slice(0,3).join(", ")) + (spNames.length > 3 ? " …" : "") : "—";
    const metadataBlock = `
        <div class="sec-title">App metadata</div>
        <div class="meta-grid">
          <div><span class="meta-label">Vendor type</span><span class="meta-value">${vendorClassEsc}</span></div>
          <div><span class="meta-label">Verified publisher</span><span class="meta-value">${verifiedPublisherEsc}</span></div>
          <div><span class="meta-label">Publisher domain</span><span class="meta-value">${publisherDomainEsc}</span></div>
          <div><span class="meta-label">App owner org</span><span class="meta-value">${ownerOrgEsc}</span></div>
          <div><span class="meta-label">Service principal names</span><span class="meta-value">${spNameDisplay}</span></div>
          <div><span class="meta-label">SP tags</span><span class="meta-value">${spTagsEsc}</span></div>
        </div>
        <div class="divider"></div>`;
    return `
      <div class="card">
        <div class="item">
          <span class="status ${item.status}"></span>
          <div>
            <h4>${escapeHtml(item.clientName||"(no name)")}</h4>
            <div class="meta">${escapeHtml(item.clientAppId||"")} • ${scopeArr.length} scope(s)</div>
            <div class="tags">${tags}</div>
          </div>
          <div class="actions">
            <button onclick='copyJson(${JSON.stringify(item).replace(/</g,"\\u003c")})' title="Copy JSON">Copy</button>
          </div>
        </div>
        <details>
          <summary>Details</summary>
          <div style="padding:12px 14px">
            ${metadataBlock}${sections}${none}
          </div>
        </details>
      </div>`;
  }).join("");
}

function copyJson(obj){ navigator.clipboard.writeText(JSON.stringify(obj,null,2)); }

// Export CSV (current filter selection)
document.getElementById("exportCsv").addEventListener("click", () => {
  const items = filterItems(report.items);
  const headers = ["clientName","clientAppId","status","isMicrosoft","isThirdParty","isHidden","tags","adminScopes","userScopes","userCount"];
  const rows = items.map(x=>{
    const adminScopes = (x.adminGrants||[]).map(a=>`${a.resource}:${a.scope}`).join(";")
    const userScopes  = (x.userGrants ||[]).map(u=>`${u.resource}:${u.scope}`).join(";")
    const userCount   = (x.userGrants ||[]).flatMap(u => asArray(u.users)).filter((v,i,arr)=>arr.indexOf(v)===i).length;
    const tags = (x.tags||[]).join(";")
    return {
      clientName: x.clientName||"",
      clientAppId: x.clientAppId||"",
      status: x.status||"",
      isMicrosoft: x.isMicrosoft ? "true" : "false",
      isThirdParty: x.isThirdParty ? "true" : "false",
      isHidden: x.isHidden ? "true" : "false",
      tags,
      adminScopes,
      userScopes,
      userCount
    };
  });
  const csv = [headers.join(",")].concat(rows.map(r=>headers.map(h=>`"${String(r[h]).replace(/"/g,'""')}"`).join(","))).join("\n");
  const blob = new Blob([csv], {type:"text/csv;charset=utf-8;"});
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a"); a.href = url; a.download = "oauth2-consents.csv"; a.click();
  URL.revokeObjectURL(url);
});

render();
</script>
</body>
</html>
'@

$escapedJson = $json.Replace('</script>','<\/script>')
$html = $htmlTemplate
$html = $html -replace '__JSON__', $escapedJson

$ts = Get-Date -Format "yyyyMMdd_HHmm"
$htmlPath = Join-Path $outDir ("OAuth2_Consent_Report_$ts.html")
Set-Content -Path $htmlPath -Encoding UTF8 -Value $html

Write-Host "`nDone:" -ForegroundColor Cyan
Write-Host "  CSV : $csvPath"
Write-Host "  HTML: $htmlPath"
