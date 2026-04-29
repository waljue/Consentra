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
    [string]$AppNameContains,
    [switch]$VerboseLogging
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

$script:LogPath = Join-Path -Path $PSScriptRoot -ChildPath 'consentra.log'
$script:DebugEnabled = $false
$script:ShowLogs = $false
$script:ProgressStep = 0
$script:ProgressTotal = 10

function Initialize-Progress {
    param([int]$TotalSteps = 10)
    $script:ProgressStep = 0
    $script:ProgressTotal = $TotalSteps
    Write-Progress -Activity 'Consentra report' -Status 'Starting...' -PercentComplete 0
}

function Update-Progress {
    param([Parameter(Mandatory)][string]$Message)
    $script:ProgressStep = [Math]::Min($script:ProgressStep + 1, $script:ProgressTotal)
    $percent = if ($script:ProgressTotal -gt 0) { [Math]::Round(($script:ProgressStep / $script:ProgressTotal) * 100) } else { 0 }
    Write-Host ("[{0}/{1}] {2}" -f $script:ProgressStep, $script:ProgressTotal, $Message)
    Write-Progress -Activity 'Consentra report' -Status $Message -PercentComplete $percent
}

function Complete-Progress {
    Write-Progress -Activity 'Consentra report' -Completed
}

function Write-Log {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO','WARN','ERROR','DEBUG')][string]$Level = 'INFO',
        [object]$Data
    )

    if ($Level -eq 'DEBUG' -and -not $script:DebugEnabled) { return }

    $timestamp = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss.fff')
    $line = "[$timestamp][$Level] $Message"
    if ($Data) {
        try {
            $json = $Data | ConvertTo-Json -Depth 6 -Compress
            $line = "$line | data=$json"
        } catch {
            $line = "$line | data=$($Data.ToString())"
        }
    }

    if ($script:ShowLogs) { Write-Host $line }
    try { Add-Content -Path $script:LogPath -Value $line } catch { }
}

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
    $tenantId = $null

    if ($response.value) {
        foreach ($org in $response.value) {
            if (-not $tenantId -and $org.id) { $tenantId = $org.id }
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
        TenantId    = $tenantId
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
    $select = 'id,appId,displayName,appOwnerOrganizationId,publisherName,tags,servicePrincipalType,servicePrincipalNames,verifiedPublisher,notes,description,preferredSingleSignOnMode,replyUrls'
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

function Get-SsoState {
    param(
        [string[]]$ReplyUrls,
        [string]$PreferredSingleSignOnMode
    )
    if (-not [string]::IsNullOrWhiteSpace($PreferredSingleSignOnMode)) {
        $mode = $PreferredSingleSignOnMode.Trim().ToLowerInvariant()
        if ($mode -like '*saml*') { return 'SAML' }
        if ($mode -like '*oidc*' -or $mode -like '*openidconnect*') { return 'OpenID Connect' }
        switch ($mode) {
            'password' { return 'Password (legacy)' }
            'notsupported' { return 'Disabled' }
            'unknownfuturevalue' { return 'Unknown' }
            default { return ($PreferredSingleSignOnMode) }
        }
    }

    $replyUris = @($ReplyUrls | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    if ($replyUris.Count -gt 0) { return 'OpenID Connect' }

    return 'Disabled'
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
    $select = 'id,appId,displayName,appOwnerOrganizationId,publisherName,publisherDomain,tags,servicePrincipalNames,servicePrincipalType,verifiedPublisher,notes,description,preferredSingleSignOnMode,replyUrls'
    $baseUri = "https://graph.microsoft.com/beta/servicePrincipals?`$count=true&`$top=999&includeHidden=true&`$select=$select"
    try {
        return Invoke-GraphPaged -Uri $baseUri
    } catch {
        Write-Warning ("Unable to include hidden service principals via includeHidden=true: {0}" -f $_.Exception.Message)
    $fallbackUri = "https://graph.microsoft.com/beta/servicePrincipals?`$count=true&`$top=999&`$select=$select"
        return Invoke-GraphPaged -Uri $fallbackUri
    }
}

function Get-AllApplications {
    $select = 'id,appId,displayName,description,web'
    $baseUri = "https://graph.microsoft.com/beta/applications?`$count=true&`$top=999&`$select=$select"
    return Invoke-GraphPaged -Uri $baseUri
}

function Get-StaleAppIndicators {
    $appIds = [System.Collections.Generic.HashSet[string]]::new()
    $spIds = [System.Collections.Generic.HashSet[string]]::new()
    $debugLogged = 0
    $sampleLogged = 0
    $endpoints = @(
        'https://graph.microsoft.com/beta/directory/recommendations/a65f6b72-abef-4f0b-88d6-3256be9a72d1_staleApps/impactedResources'
    )

    foreach ($endpoint in $endpoints) {
        try {
            $results = Invoke-GraphPaged -Uri $endpoint
        } catch {
            continue
        }

        Write-Log -Level 'DEBUG' -Message 'Stale app endpoint returned items.' -Data @{ endpoint = $endpoint; count = @($results).Count }

        foreach ($item in $results) {
            $impactType = Get-GraphResponseProperty -Response $item -PropertyName 'impactType'
            $impactCategory = Get-GraphResponseProperty -Response $item -PropertyName 'impactCategory'
            $category = Get-GraphResponseProperty -Response $item -PropertyName 'category'
            $indicator = @($impactType, $impactCategory, $category) | Where-Object { $_ } | Select-Object -First 1

            $shouldProcess = $endpoint -like '*recommendations/*_staleApps/impactedResources' -or
                             $endpoint -like '*staleApps' -or
                             ($indicator -and ($indicator.ToString() -match 'stale'))
            if ($shouldProcess) {
                $appId = Get-GraphResponseProperty -Response $item -PropertyName 'appId'
                $resourceId = Get-GraphResponseProperty -Response $item -PropertyName 'resourceId'
                $resourceType = Get-GraphResponseProperty -Response $item -PropertyName 'resourceType'
                $subjectId = Get-GraphResponseProperty -Response $item -PropertyName 'subjectId'
                $id = Get-GraphResponseProperty -Response $item -PropertyName 'id'

                $nested = @(
                    (Get-GraphResponseProperty -Response $item -PropertyName 'resource'),
                    (Get-GraphResponseProperty -Response $item -PropertyName 'resourceInfo'),
                    (Get-GraphResponseProperty -Response $item -PropertyName 'resourceMetadata'),
                    (Get-GraphResponseProperty -Response $item -PropertyName 'resourceDetails'),
                    (Get-GraphResponseProperty -Response $item -PropertyName 'recommendationResource'),
                    (Get-GraphResponseProperty -Response $item -PropertyName 'impactedResource'),
                    (Get-GraphResponseProperty -Response $item -PropertyName 'targetResource')
                ) | Where-Object { $_ }

                foreach ($node in $nested) {
                    if (-not $appId) { $appId = Get-GraphResponseProperty -Response $node -PropertyName 'appId' }
                    if (-not $resourceId) { $resourceId = Get-GraphResponseProperty -Response $node -PropertyName 'resourceId' }
                    if (-not $resourceId) { $resourceId = Get-GraphResponseProperty -Response $node -PropertyName 'id' }
                    if (-not $resourceType) { $resourceType = Get-GraphResponseProperty -Response $node -PropertyName 'resourceType' }
                    if (-not $resourceType) { $resourceType = Get-GraphResponseProperty -Response $node -PropertyName 'type' }
                    if ($appId -or $resourceId) { break }
                }

                if (-not $appId -and $resourceType -and $resourceType.ToString().ToLowerInvariant() -eq 'app') {
                    if ($subjectId) { $appId = $subjectId }
                    elseif ($id) { $appId = $id }
                }

                if (-not $appId) {
                    $portalUrl = Get-GraphResponseProperty -Response $item -PropertyName 'portalUrl'
                    if ($portalUrl -and $portalUrl -match 'appId/([0-9a-fA-F-]{36})') {
                        $appId = $Matches[1]
                    }
                }

                if ($appId) { [void]$appIds.Add($appId.ToString()) }

                if ($VerboseLogging -and $sampleLogged -lt 5) {
                    $sampleLogged++
                    Write-Log -Level 'DEBUG' -Message 'Stale app sample item.' -Data @{ resourceType = $resourceType; subjectId = $subjectId; appId = $appId; resourceId = $resourceId; id = $id }
                }

                $resourceTypeText = if ($resourceType) { $resourceType.ToString().ToLowerInvariant() } else { '' }
                if (-not $appId -and $resourceId) {
                    if ($resourceTypeText -match 'application') {
                        [void]$appIds.Add($resourceId.ToString())
                    } else {
                        [void]$spIds.Add($resourceId.ToString())
                    }
                }

                if ($id -and -not $spIds.Contains($id.ToString())) { [void]$spIds.Add($id.ToString()) }

                if (-not $appId -and -not $resourceId -and $debugLogged -lt 3) {
                    $debugLogged++
                    Write-Log -Level 'DEBUG' -Message 'Stale app item missing IDs.' -Data @{ keys = @($item.PSObject.Properties.Name); item = $item }
                }
            }
        }

        if ($appIds.Count -gt 0 -or $spIds.Count -gt 0) { break }
    }

    return [pscustomobject]@{
        AppIds = $appIds
        ServicePrincipalIds = $spIds
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

    $notesValue = Get-EffectiveString (Get-GraphResponseProperty -Response $Source -PropertyName 'notes')
    if (-not $notesValue) {
        $notesValue = Get-EffectiveString (Get-GraphResponseProperty -Response $Source -PropertyName 'description')
    }

    $replyUrlsRaw = Get-GraphResponseProperty -Response $Source -PropertyName 'replyUrls'
    $replyUrls = @()
    if ($replyUrlsRaw) { $replyUrls = @($replyUrlsRaw | ForEach-Object { $_.ToString() }) }

    $spProfile = [ordered]@{
        Id                    = Get-GraphResponseProperty -Response $Source -PropertyName 'id'
        AppId                 = Get-GraphResponseProperty -Response $Source -PropertyName 'appId'
        Name                  = Get-GraphResponseProperty -Response $Source -PropertyName 'displayName'
        Tags                  = $tags
        Notes                 = $notesValue
        PreferredSSOMode      = Get-EffectiveString (Get-GraphResponseProperty -Response $Source -PropertyName 'preferredSingleSignOnMode')
        ReplyUrls             = $replyUrls
        Publisher             = Get-EffectiveString (Get-GraphResponseProperty -Response $Source -PropertyName 'publisherName')
        PublisherDomain       = Get-EffectiveString (Get-GraphResponseProperty -Response $Source -PropertyName 'publisherDomain')
        AppOwnerOrgId         = Get-EffectiveString (Get-GraphResponseProperty -Response $Source -PropertyName 'appOwnerOrganizationId')
        VerifiedPublisherName = $verifiedPublisherName
        VerifiedPublisherId   = $verifiedPublisherId
        ServicePrincipalNames = $spNames
        ServicePrincipalType  = Get-EffectiveString (Get-GraphResponseProperty -Response $Source -PropertyName 'servicePrincipalType')
    }

    $derivedDomain = Get-DomainFromServicePrincipalNames -Names $spProfile.ServicePrincipalNames
    if ($derivedDomain) { $spProfile.PublisherDomain = $derivedDomain }

    return $spProfile
}

function Merge-ServicePrincipalProfile {
    param(
        [hashtable]$Target,
        [pscustomobject]$Source
    )

    $spProfile = New-ServicePrincipalProfile -Source $Source
    if (-not $spProfile) { return }

    foreach ($key in 'Name','AppId','AppOwnerOrgId','ServicePrincipalType','Notes','PreferredSSOMode','ReplyUrls') {
        if ($spProfile[$key]) { $Target[$key] = $spProfile[$key] }
    }

    if ($spProfile.Tags -and @($spProfile.Tags).Count -gt 0) { $Target.Tags = $spProfile.Tags }
    if ($spProfile.ServicePrincipalNames -and @($spProfile.ServicePrincipalNames).Count -gt 0) { $Target.ServicePrincipalNames = $spProfile.ServicePrincipalNames }
    if ($spProfile.Publisher) { $Target.Publisher = $spProfile.Publisher }
    if ($spProfile.PublisherDomain) { $Target.PublisherDomain = $spProfile.PublisherDomain }
    if ($spProfile.VerifiedPublisherName) { $Target.VerifiedPublisherName = $spProfile.VerifiedPublisherName }
    if ($spProfile.VerifiedPublisherId) { $Target.VerifiedPublisherId = $spProfile.VerifiedPublisherId }
}

function Get-MicrosoftVendorInfo {
    param(
        [object]$SpProfile,
        [string[]]$TenantDomains,
        [string]$TenantDisplayName
    )

    if (-not $SpProfile) {
        $SpProfile = [ordered]@{}
    }

    $vendorInfo = [ordered]@{
        IsMicrosoft                          = $false
        IsThirdParty                         = $false
        IsHomeTenant                         = $false
        OwnerOrg                             = $SpProfile.AppOwnerOrgId
        OwnerOrgMatch                        = $false
        Publisher                            = $SpProfile.Publisher
        PublisherMatched                     = $false
        PublisherDomain                      = $SpProfile.PublisherDomain
        PublisherDomainMatchesMicrosoft      = $false
        PublisherDomainMatchesTenant         = $false
        HasIntegratedTag                     = $false
        HasGalleryTag                        = $false
        KnownAppIdMatch                      = $null
        Tags                                 = $SpProfile.Tags
        ServicePrincipalNames                = $SpProfile.ServicePrincipalNames
        ServicePrincipalNameMatchesMicrosoft = $false
        ServicePrincipalNameMatchesTenant    = $false
        ServicePrincipalType                 = $SpProfile.ServicePrincipalType
        VerifiedPublisherName                = $SpProfile.VerifiedPublisherName
        VerifiedPublisherId                  = $SpProfile.VerifiedPublisherId
        VerifiedPublisherMatchesMicrosoft    = $false
    }

    if ($SpProfile.AppOwnerOrgId) {
        $vendorInfo.OwnerOrgMatch = ($script:KnownMicrosoftTenantIds -contains $SpProfile.AppOwnerOrgId.ToLowerInvariant())
    }

    if ($SpProfile.Publisher) {
        $vendorInfo.PublisherMatched = ($SpProfile.Publisher.ToLowerInvariant() -match '\bmicrosoft\b')
    }

    if ($SpProfile.PublisherDomain) {
        $domain = $SpProfile.PublisherDomain.ToLowerInvariant()
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

    if ($SpProfile.VerifiedPublisherName) {
        $vendorInfo.VerifiedPublisherMatchesMicrosoft = ($SpProfile.VerifiedPublisherName.ToLowerInvariant() -match '\bmicrosoft\b')
    }
    if (-not $vendorInfo.VerifiedPublisherMatchesMicrosoft -and $SpProfile.VerifiedPublisherId) {
        $vendorInfo.VerifiedPublisherMatchesMicrosoft = ($script:KnownMicrosoftTenantIds -contains $SpProfile.VerifiedPublisherId.ToLowerInvariant())
    }

    if ($SpProfile.ServicePrincipalNames -and @($SpProfile.ServicePrincipalNames).Count -gt 0) {
        foreach ($name in $SpProfile.ServicePrincipalNames) {
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

    $tags = $SpProfile.Tags
    if ($tags) {
        $vendorInfo.HasIntegratedTag = $tags -contains 'WindowsAzureActiveDirectoryIntegratedApp'
        $vendorInfo.HasGalleryTag = ($tags -contains 'WindowsAzureActiveDirectoryGalleryApplicationPrimaryV1' -or $tags -contains 'WindowsAzureActiveDirectoryGalleryApplicationNonPrimaryV1')
    }

    if ($SpProfile.AppId -and ($script:KnownMicrosoftAppIds -contains $SpProfile.AppId)) {
        $vendorInfo.KnownAppIdMatch = $SpProfile.AppId
    }

    $tagIndicatesMicrosoft = $vendorInfo.HasIntegratedTag -or $vendorInfo.HasGalleryTag
    $publisherAvailable = [bool]$SpProfile.Publisher

    $vendorInfo.IsMicrosoft = $vendorInfo.OwnerOrgMatch -or
                              $vendorInfo.VerifiedPublisherMatchesMicrosoft -or
                              $vendorInfo.PublisherDomainMatchesMicrosoft -or
                              $vendorInfo.ServicePrincipalNameMatchesMicrosoft -or
                              ($vendorInfo.PublisherMatched -and ($tagIndicatesMicrosoft -or -not $tagIndicatesMicrosoft -and -not $publisherAvailable)) -or
                              (-not $publisherAvailable -and $tagIndicatesMicrosoft) -or
                              ($null -ne $vendorInfo.KnownAppIdMatch)

    if (-not $vendorInfo.IsMicrosoft) {
        $isTenantPublisher = $vendorInfo.PublisherDomainMatchesTenant -or
                              $vendorInfo.ServicePrincipalNameMatchesTenant -or
                              ($SpProfile.Publisher -and $TenantDisplayName -and ($SpProfile.Publisher -eq $TenantDisplayName))
        if ($isTenantPublisher) { $vendorInfo.IsHomeTenant = $true }
        if ($SpProfile.VerifiedPublisherName) {
            $vendorInfo.IsThirdParty = (-not $vendorInfo.VerifiedPublisherMatchesMicrosoft -and -not $isTenantPublisher)
        } elseif ($SpProfile.PublisherDomain) {
            $vendorInfo.IsThirdParty = (-not $vendorInfo.PublisherDomainMatchesMicrosoft -and -not $vendorInfo.PublisherDomainMatchesTenant -and -not $isTenantPublisher)
        } elseif ($SpProfile.Publisher) {
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

function Test-LowImpactScopeSet {
    param([string[]]$Scopes)
    if (-not $Scopes -or @($Scopes).Count -eq 0) { return $false }
    foreach ($scope in $Scopes) {
        if ($lowImpactExact -notcontains $scope) { return $false }
    }
    return $true
}

function Get-ScopeImpactLevel {
    param([string]$Scope)
    if ([string]::IsNullOrWhiteSpace($Scope)) { return 'unknown' }
    if (Test-CriticalScope -Scope $Scope) { return 'critical' }
    if ($lowImpactExact -contains $Scope) { return 'low' }
    return 'elevated'
}

Initialize-GraphContext
$script:DebugEnabled = [bool]$VerboseLogging -or ($VerbosePreference -ne 'SilentlyContinue')
$script:ShowLogs = $script:DebugEnabled
Initialize-Progress -TotalSteps 11
Update-Progress -Message 'Initialize Graph connection'
Write-Log -Message 'Graph context initialized.'
$tenantMetadata = Get-TenantMetadata
$tenantDomains = $tenantMetadata.Domains
$tenantDisplayName = $tenantMetadata.DisplayName
$tenantId = $tenantMetadata.TenantId
Update-Progress -Message 'Load tenant metadata'
Write-Log -Message 'Tenant metadata loaded.' -Data @{ displayName = $tenantDisplayName; domains = $tenantDomains }

$grantUri = 'https://graph.microsoft.com/beta/oauth2PermissionGrants?$count=true&$top=999'
$grants = Invoke-GraphPaged -Uri $grantUri
Update-Progress -Message 'Load OAuth2 grants'
Write-Log -Message 'OAuth grants retrieved.' -Data @{ count = @($grants).Count }
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
Update-Progress -Message 'Load directory objects'
Write-Log -Message 'Directory objects retrieved.' -Data @{ count = @($directoryObjects).Count }
$servicePrincipals = @{}
$users = @{}

foreach ($obj in $directoryObjects) {
    switch ($obj.'@odata.type') {
        '#microsoft.graph.servicePrincipal' {
            $spProfile = New-ServicePrincipalProfile -Source $obj
            if ($spProfile) { $servicePrincipals[$obj.id] = $spProfile }
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
    Update-Progress -Message 'Load service principal metadata'
        Write-Log -Message 'Service principal metadata retrieved.' -Data @{ count = @($metadataById.Keys).Count }
    } catch {
        Write-Warning ("Unable to fetch extended service principal metadata: {0}" -f $_.Exception.Message)
        Write-Log -Level 'WARN' -Message 'Unable to fetch extended service principal metadata.' -Data $_.Exception.Message
        $metadataById = @{}
    }
}

foreach ($id in $metadataById.Keys) {
    if (-not $servicePrincipals.ContainsKey($id)) {
        $spProfile = New-ServicePrincipalProfile -Source $metadataById[$id]
        if ($spProfile) { $servicePrincipals[$id] = $spProfile }
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

$applicationsByAppId = @{}
try {
    $allApplications = Get-AllApplications
    Update-Progress -Message 'Load app registrations'
    Write-Log -Message 'Applications list retrieved.' -Data @{ count = @($allApplications).Count }
    foreach ($app in $allApplications) {
        if ($app.appId) { $applicationsByAppId[$app.appId] = $app }
    }
} catch {
    Write-Warning ("Unable to retrieve applications list: {0}" -f $_.Exception.Message)
    Write-Log -Level 'WARN' -Message 'Unable to retrieve applications list.' -Data $_.Exception.Message
    $applicationsByAppId = @{}
}

$staleIndicators = $null
try {
    $staleIndicators = Get-StaleAppIndicators
    Update-Progress -Message 'Fetch stale apps (recommendation)'
    Write-Log -Message 'Stale app indicators retrieved.' -Data @{ appIds = $staleIndicators.AppIds.Count; spIds = $staleIndicators.ServicePrincipalIds.Count }
} catch {
    Write-Log -Level 'WARN' -Message 'Unable to retrieve stale app indicators.' -Data $_.Exception.Message
}
$staleAppIds = if ($staleIndicators) { $staleIndicators.AppIds } else { [System.Collections.Generic.HashSet[string]]::new() }
$staleSpIds = if ($staleIndicators) { $staleIndicators.ServicePrincipalIds } else { [System.Collections.Generic.HashSet[string]]::new() }
if (-not $staleAppIds) { $staleAppIds = [System.Collections.Generic.HashSet[string]]::new() }
if (-not $staleSpIds) { $staleSpIds = [System.Collections.Generic.HashSet[string]]::new() }

$appRoleAssignmentLookup = @{}
$userConsentClientIds = [System.Collections.Generic.HashSet[string]]::new()
foreach ($grant in $grants) {
    if ($grant.consentType -eq 'Principal' -and $grant.clientId) {
        [void]$userConsentClientIds.Add($grant.clientId)
    }
}
if ($userConsentClientIds.Count -gt 0) {
    Update-Progress -Message 'Fetch app role assignments for user consent links'
    foreach ($cId in $userConsentClientIds) {
        try {
            $roleAssignments = Invoke-GraphPaged -Uri "https://graph.microsoft.com/v1.0/servicePrincipals/$cId/appRoleAssignedTo?`$select=id,principalId"
            foreach ($ra in $roleAssignments) {
                if ($ra.id -and $ra.principalId) {
                    $appRoleAssignmentLookup["$cId|$($ra.principalId)"] = $ra.id
                }
            }
        } catch {
            Write-Log -Level 'WARN' -Message "Unable to fetch appRoleAssignedTo for SP $cId" -Data $_.Exception.Message
        }
    }
}

$recordsByClient = $expandedGrants | Group-Object ClientId
$appSummaries = @()
Update-Progress -Message 'Build app summaries'
foreach ($group in $recordsByClient) {
    try {
        $recordsForClient = $group.Group
        $clientProfile = $servicePrincipals[$group.Name]
        if (-not $clientProfile) {
            Write-Log -Level 'WARN' -Message 'Client profile missing, using fallback.' -Data @{ clientId = $group.Name }
            $clientProfile = [ordered]@{
                Id                = $group.Name
                Name              = $group.Name
                AppId             = $null
                Tags              = @()
                Notes             = $null
                ReplyUrls         = @()
                PreferredSSOMode  = $null
                ServicePrincipalNames = @()
            }
        }
        $appInfo = $null
        if ($clientProfile.AppId -and $applicationsByAppId.ContainsKey($clientProfile.AppId)) {
            $appInfo = $applicationsByAppId[$clientProfile.AppId]
        }
        $appDescription = if ($appInfo) { Get-EffectiveString (Get-GraphResponseProperty -Response $appInfo -PropertyName 'description') }
        if (-not $appDescription) { $appDescription = $clientProfile.Notes }
        $redirectUris = @()
        if ($appInfo) {
            $webInfo = Get-GraphResponseProperty -Response $appInfo -PropertyName 'web'
            $redirectUrisRaw = if ($webInfo) { Get-GraphResponseProperty -Response $webInfo -PropertyName 'redirectUris' }
            if ($redirectUrisRaw) { $redirectUris = @($redirectUrisRaw) }
        }
        $replyUrls = if ($clientProfile.ReplyUrls) { $clientProfile.ReplyUrls } else { $redirectUris }
        $allScopes = @(($recordsForClient | Select-Object -Expand Scope) | Sort-Object -Unique)

        $isCritical = $false
        foreach ($scope in $allScopes) {
            if (Test-CriticalScope -Scope $scope) { $isCritical = $true; break }
        }
        $status = if ($isCritical) { 'fail' } elseif (Test-LowImpactScopeSet -Scopes $allScopes) { 'pass' } else { 'warn' }

        $administratorGrants = @()
        $userGrants = @()

        foreach ($adminGroup in (($recordsForClient | Where-Object { $_.ConsentType -eq 'AllPrincipals' }) | Group-Object ResourceName, Scope)) {
            $sample = $adminGroup.Group[0]
            $administratorGrants += [pscustomobject]@{
                resource = $sample.ResourceName
                scope    = $sample.Scope
                type     = 'Admin'
                impact   = Get-ScopeImpactLevel -Scope $sample.Scope
                grantId  = $sample.GrantId
            }
        }

        foreach ($userGroup in (($recordsForClient | Where-Object { $_.ConsentType -eq 'Principal' }) | Group-Object ResourceName, Scope)) {
            $sample = $userGroup.Group[0]
            $userDetails = @($userGroup.Group | Where-Object { $_.PrincipalId } | ForEach-Object {
                $principalId = $_.PrincipalId
                $upn = if ($users.ContainsKey($principalId)) { $users[$principalId].UPN } else { $null }
                $assignmentId = $appRoleAssignmentLookup["$($group.Name)|$principalId"]
                [pscustomobject]@{
                    principalId  = $principalId
                    upn          = $upn
                    grantId      = $_.GrantId
                    assignmentId = $assignmentId
                }
            } | Where-Object { $_.upn } | Group-Object principalId | ForEach-Object { $_.Group[0] } | Sort-Object upn)
            $userUpns = @($userDetails | ForEach-Object { $_.upn })
            $userGrants += [pscustomobject]@{
                resource    = $sample.ResourceName
                scope       = $sample.Scope
                type        = 'User'
                users       = $userUpns
                userDetails = $userDetails
                userCount   = $userDetails.Count
                impact      = Get-ScopeImpactLevel -Scope $sample.Scope
                grantId     = $sample.GrantId
            }
        }

        $vendorInfo = Get-MicrosoftVendorInfo -SpProfile $clientProfile -TenantDomains $tenantDomains -TenantDisplayName $tenantDisplayName
        $isHiddenApp = Test-ServicePrincipalHidden -Tags $clientProfile.Tags
        $isStaleApp = $false
        if ($clientProfile.AppId -and $staleAppIds.Contains($clientProfile.AppId)) { $isStaleApp = $true }
        elseif ($group.Name -and $staleSpIds.Contains($group.Name)) { $isStaleApp = $true }
        $userScopeCount = @($userGrants | Select-Object -ExpandProperty scope -Unique).Count
        $userCount = @($userGrants | ForEach-Object { $_.users } | Where-Object { $_ } | Select-Object -Unique).Count

        $tags = @()
        if ($vendorInfo.IsMicrosoft) { $tags += 'Microsoft' }
        elseif ($vendorInfo.IsThirdParty) { $tags += '3rd Party' }
        elseif ($vendorInfo.IsHomeTenant) { $tags += 'Home Tenant' }
        if ($isHiddenApp) { $tags += 'Hidden' }
        if ($isStaleApp) { $tags += 'staleApps' }
        if ($administratorGrants) { $tags += 'Admin Consent' }
        if ($userGrants) { $tags += 'User Consent' }
        $tags += "user-perm: $userScopeCount"
        $tags += "user-count: $userCount"

        $appSummaries += [pscustomobject]@{
            clientId      = $group.Name
            clientName    = $clientProfile.Name
            clientAppId   = $clientProfile.AppId
            notes         = $appDescription
            ssoState      = Get-SsoState -ReplyUrls $replyUrls -PreferredSingleSignOnMode $clientProfile.PreferredSSOMode
            staleApp      = $isStaleApp
            status        = $status
            tags          = $tags
            isMicrosoft   = $vendorInfo.IsMicrosoft
            isThirdParty  = $vendorInfo.IsThirdParty
            isHomeTenant  = $vendorInfo.IsHomeTenant
            isHidden      = $isHiddenApp
            vendorInfo    = $vendorInfo
            adminGrants   = $administratorGrants
            userGrants    = $userGrants
            allScopes     = $allScopes
        }
    } catch {
        Write-Log -Level 'ERROR' -Message 'Failed to build summary for client.' -Data @{ clientId = $group.Name; error = $_.Exception.Message; line = $_.InvocationInfo.ScriptLineNumber }
        throw
    }
}

$summariesByClientId = @{}
foreach ($summary in $appSummaries) {
    if ($summary.clientId) { $summariesByClientId[$summary.clientId] = $summary }
}

$allServicePrincipals = @()
try {
    $allServicePrincipals = Get-AllServicePrincipals
    Update-Progress -Message 'Enrich with all enterprise apps'
} catch {
    Write-Warning ("Unable to retrieve complete service principal list: {0}" -f $_.Exception.Message)
    $allServicePrincipals = @()
}

foreach ($spRaw in $allServicePrincipals) {
    $spId = Get-GraphResponseProperty -Response $spRaw -PropertyName 'id'
    if (-not $spId) { continue }

    $spProfile = $null
    if ($servicePrincipals.ContainsKey($spId)) {
        $spProfile = $servicePrincipals[$spId]
        Merge-ServicePrincipalProfile -Target $spProfile -Source $spRaw
    } else {
        $spProfile = New-ServicePrincipalProfile -Source $spRaw
        if ($spProfile) { $servicePrincipals[$spId] = $spProfile }
    }
    if (-not $spProfile) { continue }

    $vendorInfo = Get-MicrosoftVendorInfo -SpProfile $spProfile -TenantDomains $tenantDomains -TenantDisplayName $tenantDisplayName
    $isHiddenSp = Test-ServicePrincipalHidden -Tags $spProfile.Tags
    $isStaleApp = $false
    if ($spProfile.AppId -and $staleAppIds.Contains($spProfile.AppId)) { $isStaleApp = $true }
    elseif ($spId -and $staleSpIds.Contains($spId)) { $isStaleApp = $true }

    if ($summariesByClientId.ContainsKey($spId)) {
        $summary = $summariesByClientId[$spId]
        $summary.vendorInfo = $vendorInfo
        $summary.isMicrosoft = $vendorInfo.IsMicrosoft
        $summary.isThirdParty = $vendorInfo.IsThirdParty
        $summary.isHomeTenant = $vendorInfo.IsHomeTenant
        $summary.isHidden = $isHiddenSp
        $summary.tags = @($summary.tags | Where-Object { ($_ -ne 'Microsoft') -and ($_ -ne '3rd Party') -and ($_ -ne 'Home Tenant') })
        if ($vendorInfo.IsMicrosoft) { $summary.tags += 'Microsoft' }
        elseif ($vendorInfo.IsThirdParty) { $summary.tags += '3rd Party' }
        elseif ($vendorInfo.IsHomeTenant) { $summary.tags += 'Home Tenant' }
        if ($isStaleApp -and ($summary.tags -notcontains 'staleApps')) { $summary.tags += 'staleApps' }
        $summary.staleApp = $isStaleApp
    } else {
        $nameFallback = if ($spProfile.Name) { $spProfile.Name } elseif ($spProfile.AppId) { $spProfile.AppId } else { $spId }
        $tags = @()
        if ($vendorInfo.IsMicrosoft) { $tags += 'Microsoft' }
        elseif ($vendorInfo.IsThirdParty) { $tags += '3rd Party' }
        elseif ($vendorInfo.IsHomeTenant) { $tags += 'Home Tenant' }
        if ($isHiddenSp) { $tags += 'Hidden' }
        if ($isStaleApp) { $tags += 'staleApps' }
        $tags += 'No OAuth grants'

        $appInfo = $null
        if ($spProfile.AppId -and $applicationsByAppId.ContainsKey($spProfile.AppId)) {
            $appInfo = $applicationsByAppId[$spProfile.AppId]
        }
        $appDescription = if ($appInfo) { Get-EffectiveString (Get-GraphResponseProperty -Response $appInfo -PropertyName 'description') }
        if (-not $appDescription) { $appDescription = $spProfile.Notes }
        $redirectUris = @()
        if ($appInfo) {
            $webInfo = Get-GraphResponseProperty -Response $appInfo -PropertyName 'web'
            $redirectUrisRaw = if ($webInfo) { Get-GraphResponseProperty -Response $webInfo -PropertyName 'redirectUris' }
            if ($redirectUrisRaw) { $redirectUris = @($redirectUrisRaw) }
        }
        $replyUrls = if ($spProfile.ReplyUrls) { $spProfile.ReplyUrls } else { $redirectUris }

        $newSummary = [pscustomobject]@{
            clientId     = $spId
            clientName   = $nameFallback
            clientAppId  = $spProfile.AppId
            notes        = $appDescription
            ssoState     = Get-SsoState -ReplyUrls $replyUrls -PreferredSingleSignOnMode $spProfile.PreferredSSOMode
            staleApp     = $isStaleApp
            status       = 'pass'
            tags         = $tags
            isMicrosoft  = $vendorInfo.IsMicrosoft
            isThirdParty = $vendorInfo.IsThirdParty
            isHomeTenant = $vendorInfo.IsHomeTenant
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

$servicePrincipalAppIds = [System.Collections.Generic.HashSet[string]]::new()
foreach ($spProfile in $servicePrincipals.Values) {
    if ($spProfile.AppId) { [void]$servicePrincipalAppIds.Add($spProfile.AppId) }
}

foreach ($appId in $applicationsByAppId.Keys) {
    if ($servicePrincipalAppIds.Contains($appId)) { continue }

    $appInfo = $applicationsByAppId[$appId]
    $appDescription = Get-EffectiveString (Get-GraphResponseProperty -Response $appInfo -PropertyName 'description')
    $redirectUris = @()
    $webInfo = Get-GraphResponseProperty -Response $appInfo -PropertyName 'web'
    $redirectUrisRaw = if ($webInfo) { Get-GraphResponseProperty -Response $webInfo -PropertyName 'redirectUris' }
    if ($redirectUrisRaw) { $redirectUris = @($redirectUrisRaw) }

    $appProfile = [ordered]@{
        AppId                 = $appId
        Name                  = Get-EffectiveString (Get-GraphResponseProperty -Response $appInfo -PropertyName 'displayName')
        Tags                  = @()
        Publisher             = $null
        PublisherDomain       = $null
        AppOwnerOrgId         = $null
        VerifiedPublisherName = $null
        VerifiedPublisherId   = $null
        ServicePrincipalNames = @()
        ServicePrincipalType  = $null
    }

    $vendorInfo = Get-MicrosoftVendorInfo -SpProfile $appProfile -TenantDomains $tenantDomains -TenantDisplayName $tenantDisplayName
    $isStaleApp = $false
    if ($staleAppIds.Contains($appId)) { $isStaleApp = $true }

    $tags = @('No Service Principal')
    if ($vendorInfo.IsMicrosoft) { $tags += 'Microsoft' }
    elseif ($vendorInfo.IsThirdParty) { $tags += '3rd Party' }
    elseif ($vendorInfo.IsHomeTenant) { $tags += 'Home Tenant' }
    if ($isStaleApp) { $tags += 'staleApps' }

    $appSummaries += [pscustomobject]@{
        clientId      = "app:$appId"
        clientName    = $appProfile.Name
        clientAppId   = $appId
        notes         = $appDescription
        ssoState      = Get-SsoState -ReplyUrls $redirectUris -PreferredSingleSignOnMode $null
    staleApp      = $isStaleApp
        status        = 'pass'
        tags          = $tags
        isMicrosoft   = $vendorInfo.IsMicrosoft
        isThirdParty  = $vendorInfo.IsThirdParty
        isHomeTenant  = $vendorInfo.IsHomeTenant
        isHidden      = $false
        vendorInfo    = $vendorInfo
        adminGrants   = @()
        userGrants    = @()
        allScopes     = @()
    }
}

Update-Progress -Message 'Generate CSV/HTML report'

        $impactSummary = @{
            microsoft = [ordered]@{ critical = 0; elevated = 0; low = 0 }
            thirdParty = [ordered]@{ critical = 0; elevated = 0; low = 0 }
            homeTenant = [ordered]@{ critical = 0; elevated = 0; low = 0 }
        }

        foreach ($summary in $appSummaries) {
            $vendorKey = if ($summary.isMicrosoft) {
                'microsoft'
            }
            elseif ($summary.isThirdParty) {
                'thirdParty'
            }
            elseif ($summary.isHomeTenant) {
                'homeTenant'
            }
            else {
                $null
            }

            if (-not $vendorKey) { continue }

            $grants = @()
            if ($summary.adminGrants) { $grants += $summary.adminGrants }
            if ($summary.userGrants) { $grants += $summary.userGrants }

            foreach ($grant in $grants) {
                if (-not $grant) { continue }
                $impactKey = switch ($grant.impact) {
                    'critical' { 'critical' }
                    'low' { 'low' }
                    default { 'elevated' }
                }
                $impactSummary[$vendorKey][$impactKey]++
            }
        }

$totalApps = @($appSummaries).Count
$failCount = @($appSummaries | Where-Object status -eq 'fail').Count
$passCount = @($appSummaries | Where-Object status -eq 'pass').Count
$warnCount = @($appSummaries | Where-Object status -eq 'warn').Count
$microsoftCount = @($appSummaries | Where-Object isMicrosoft).Count
$thirdPartyCount = @($appSummaries | Where-Object isThirdParty).Count
$homeTenantCount = @($appSummaries | Where-Object isHomeTenant).Count
$enterpriseCount = @($appSummaries | Where-Object { -not $_.isMicrosoft }).Count
$staleCount = @($appSummaries | Where-Object staleApp).Count

$outDir = (Get-Location).Path
$timestamp = Get-Date -Format 'yyyyMMdd_HHmm'
$csvPath = Join-Path $outDir ("OAuth2_Consent_Apps_$timestamp.csv")

$appSummaries |
    Select-Object clientName,
                  clientAppId,
                  status,
                  @{n='isMicrosoft';e={$_.isMicrosoft}},
                  @{n='isThirdParty';e={$_.isThirdParty}},
                  @{n='isHomeTenant';e={$_.isHomeTenant}},
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
                  @{n='notes';e={$_.notes}},
                  @{n='ssoState';e={$_.ssoState}},
                  @{n='staleApp';e={$_.staleApp}},
                  @{n='tags';e={$_.tags -join ';'}},
                  @{n='adminScopes';e={($_.adminGrants | ForEach-Object { "{0}:{1}" -f $_.resource, $_.scope }) -join ';'}},
                  @{n='userScopes'; e={($_.userGrants  | ForEach-Object { "{0}:{1}" -f $_.resource, $_.scope }) -join ';'}},
                  @{n='userCount';  e={@($_.userGrants  | ForEach-Object { $_.users } | Where-Object { $_ } | Select-Object -Unique).Count}} |
    Export-Csv -NoTypeInformation -Encoding UTF8 -Path $csvPath

$report = [pscustomobject]@{
    generatedAt = (Get-Date).ToString('s')
    tenantId    = $tenantId
    totals      = @{ total=$totalApps; pass=$passCount; warn=$warnCount; fail=$failCount; microsoft=$microsoftCount; thirdParty=$thirdPartyCount; homeTenant=$homeTenantCount; enterprise=$enterpriseCount; stale=$staleCount }
    impactSummary = $impactSummary
    items       = $appSummaries
}

$json = $report | ConvertTo-Json -Depth 12
$jsonSafe = $json.Replace('</script>','<\/script>').Replace('<','\u003c')
$htmlTemplate = @'
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>OAuth2 Consent Report</title>
<style>
  @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600&family=JetBrains+Mono:wght@400;500&family=Space+Grotesk:wght@600;700&display=swap');
  :root{
    --red:#E63946;--blue:#00B4D8;--amber:#FFB703;--green:#2DC653;
    --white:#F0F0F0;--muted:#888;--muted2:#555;
    --bg:#0A0A0F;--surface:#12121A;--elevated:#1A1A2E;--border:#1E2030;
    --bg-red:rgba(230,57,70,.12);--bg-amber:rgba(255,183,3,.1);
    --bg-green:rgba(45,198,83,.08);--bg-blue:rgba(0,180,216,.1);
    --font-head:'Space Grotesk',sans-serif;
    --font-body:'Inter',sans-serif;
    --font-mono:'JetBrains Mono',monospace;
  }
  *,*::before,*::after{box-sizing:border-box;margin:0;padding:0}
  html,body{height:100%;background:var(--bg);color:var(--white);font:13px/1.5 var(--font-body);-webkit-font-smoothing:antialiased}
  a{color:var(--blue);text-decoration:none}
  a:hover{text-decoration:underline;color:#33c9e8}
  ::selection{background:var(--red);color:var(--white)}
  header{position:sticky;top:0;background:var(--bg);border-bottom:2px solid var(--red);z-index:10}
  .wrap{max-width:1260px;margin:0 auto;padding:10px 20px}
  .brand-row{display:flex;align-items:center;justify-content:space-between;gap:12px;padding-bottom:10px;border-bottom:1px solid var(--border);margin-bottom:10px}
  .brand-name{font-family:var(--font-head);font-size:19px;font-weight:700;letter-spacing:.14em;text-transform:uppercase;color:var(--white);line-height:1}
  .brand-name .dot{color:var(--red)}
  .brand-sub{font-size:11px;color:var(--muted);font-family:var(--font-mono);margin-top:3px;letter-spacing:.04em}
  .brand-meta{text-align:right;font-family:var(--font-mono);font-size:11px;color:var(--muted);line-height:1.6}
  .brand-meta b{color:var(--white);font-size:17px;display:block;font-family:var(--font-head)}
  .impact-section{border:1px solid var(--border);border-left:3px solid var(--red);background:var(--surface);border-radius:3px;padding:10px 14px;margin-bottom:10px}
  .impact-heading{font-size:9px;font-family:var(--font-mono);letter-spacing:.14em;text-transform:uppercase;color:var(--muted);margin-bottom:8px}
  .impact-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(200px,1fr));gap:8px}
  .impact-card{background:var(--elevated);border-radius:2px;padding:8px 12px;border-left:3px solid var(--border)}
  .impact-card[data-vendor="microsoft"]{border-left-color:var(--blue)}
  .impact-card[data-vendor="thirdParty"]{border-left-color:var(--amber)}
  .impact-card[data-vendor="homeTenant"]{border-left-color:var(--green)}
  .impact-card h5{margin:0 0 1px;font-size:11px;font-family:var(--font-head);color:var(--white);font-weight:600}
  .impact-total{font-size:10px;color:var(--muted);margin-bottom:7px;font-family:var(--font-mono)}
  .impact-bars{display:flex;flex-direction:column;gap:4px}
  .impact-bar{display:grid;grid-template-columns:52px 1fr 28px;gap:5px;align-items:center;font-size:10px;font-family:var(--font-mono)}
  .impact-bar-label{text-transform:uppercase;letter-spacing:.06em;color:var(--muted)}
  .impact-bar-track{height:4px;background:var(--bg);border-radius:999px;overflow:hidden}
  .impact-bar-fill{height:100%;border-radius:999px;transition:width .3s}
  .impact-bar[data-level="critical"] .impact-bar-fill{background:var(--red)}
  .impact-bar[data-level="elevated"] .impact-bar-fill{background:var(--amber)}
  .impact-bar[data-level="low"] .impact-bar-fill{background:var(--green)}
  .impact-count{font-weight:600;text-align:right}
  .filter-row{display:flex;gap:5px;align-items:center;flex-wrap:wrap;margin-bottom:5px}
  .filter-row:last-child{margin-bottom:0}
  .filter-label{font-size:9px;font-family:var(--font-mono);letter-spacing:.12em;text-transform:uppercase;color:var(--muted2);min-width:40px;flex-shrink:0}
  .filter-sep{width:1px;height:16px;background:var(--border);margin:0 2px;flex-shrink:0}
  .chip{display:inline-flex;align-items:center;gap:5px;border:1px solid var(--border);background:var(--surface);color:var(--muted);padding:3px 9px;border-radius:2px;font-size:11px;font-family:var(--font-mono);cursor:pointer;user-select:none;transition:border-color .15s,color .15s,background .15s;line-height:1.5}
  .chip b{font-weight:600;color:var(--white)}
  .chip:hover{border-color:#333;color:var(--white)}
  .chip-red.active{border-color:var(--red);background:var(--bg-red);color:var(--red)}
  .chip-red.active b{color:var(--red)}
  .chip-amber.active{border-color:var(--amber);background:var(--bg-amber);color:var(--amber)}
  .chip-amber.active b{color:var(--amber)}
  .chip-green.active{border-color:var(--green);background:var(--bg-green);color:var(--green)}
  .chip-green.active b{color:var(--green)}
  .chip-blue.active{border-color:var(--blue);background:var(--bg-blue);color:var(--blue)}
  .chip-blue.active b{color:var(--blue)}
  .chip-toggle.active{border-color:#444;background:var(--elevated);color:var(--white)}
  .dot-s{width:6px;height:6px;border-radius:999px;flex-shrink:0}
  .dot-red{background:var(--red)}.dot-amber{background:var(--amber)}
  .dot-green{background:var(--green)}.dot-blue{background:var(--blue)}
  .dot-mixed{background:linear-gradient(135deg,var(--amber) 50%,var(--blue) 50%)}
  .dot-grey{background:var(--muted2)}
  input[type="search"]{background:var(--surface);border:1px solid var(--border);color:var(--white);padding:4px 10px;border-radius:2px;font-family:var(--font-body);font-size:12px;width:200px;transition:border-color .15s}
  input[type="search"]::placeholder{color:var(--muted2)}
  input[type="search"]:focus{outline:none;border-color:var(--blue)}
  button{background:var(--surface);border:1px solid var(--border);color:var(--muted);padding:4px 12px;border-radius:2px;cursor:pointer;font-family:var(--font-body);font-size:12px;transition:border-color .15s,color .15s}
  button:hover{border-color:#444;color:var(--white)}
  #list{max-width:1260px;margin:6px auto 0;display:flex;flex-direction:column;gap:6px;padding:0 20px 20px}
  .list-counter{max-width:1260px;margin:6px auto 0;padding:0 20px;color:var(--muted);font-size:11px;font-family:var(--font-mono)}
  .list-counter b{color:var(--white)}
  .card{background:var(--surface);border:1px solid var(--border);border-radius:3px;overflow:hidden;border-left:3px solid var(--border)}
  .card.status-fail{border-left-color:var(--red)}
  .card.status-warn{border-left-color:var(--amber)}
  .card.status-pass{border-left-color:var(--green)}
  .item{display:grid;grid-template-columns:1fr auto;gap:12px;align-items:start;padding:11px 14px;border-bottom:1px solid var(--border)}
  .app-name{font-family:var(--font-head);font-size:14px;font-weight:600;color:var(--white)}
  .app-id{font-family:var(--font-mono);font-size:10px;color:var(--muted);margin-top:3px}
  .tags{display:flex;gap:4px;flex-wrap:wrap;margin-top:7px}
  .tag{font-family:var(--font-mono);font-size:9px;letter-spacing:.06em;text-transform:uppercase;padding:2px 6px;border-radius:2px;border:1px solid var(--border);color:var(--muted);cursor:default}
  .tag-admin{background:rgba(45,198,83,.1);border-color:rgba(45,198,83,.35);color:#7fffaa}
  .tag-user{background:rgba(0,180,216,.1);border-color:rgba(0,180,216,.35);color:#7fe8ff}
  .tag-microsoft{background:rgba(0,180,216,.07);border-color:rgba(0,180,216,.25);color:#aadaf0}
  .tag-thirdparty{background:rgba(255,183,3,.07);border-color:rgba(255,183,3,.25);color:#ffe080}
  .tag-home{background:rgba(45,198,83,.07);border-color:rgba(45,198,83,.25);color:#a0f0c0}
  .tag-hidden{background:rgba(136,136,136,.08);border-color:rgba(136,136,136,.25);color:var(--muted);font-style:italic}
  .tag-stale{background:rgba(255,183,3,.07);border-color:rgba(255,183,3,.25);color:#ffe080}
  .tag-clickable{cursor:pointer;transition:filter .15s}
  .tag-clickable:hover{filter:brightness(1.25)}
  .tag.tag-active{outline:1px solid var(--blue)}
  .actions{display:flex;gap:5px;align-items:center;flex-shrink:0;padding-top:1px}
  details{background:var(--bg)}
  details[open] summary{border-bottom:1px solid var(--border)}
  summary{list-style:none;padding:7px 14px;cursor:pointer;color:var(--muted);font-size:10px;font-family:var(--font-mono);letter-spacing:.06em;text-transform:uppercase}
  summary::-webkit-details-marker{display:none}
  summary::before{content:'\25b6  ';font-size:8px;color:var(--muted2)}
  details[open] summary::before{content:'\25bc  '}
  .sec-title{font-size:9px;font-family:var(--font-mono);letter-spacing:.1em;text-transform:uppercase;color:var(--muted);margin:10px 0 5px}
  .meta-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(190px,1fr));gap:5px 18px;margin:0 0 12px}
  .meta-label{display:block;font-size:9px;font-family:var(--font-mono);letter-spacing:.08em;text-transform:uppercase;color:var(--muted);margin-bottom:2px}
  .meta-value{font-size:12px;color:#bbb}
  .divider{height:1px;background:var(--border);margin:8px 0}
  .consent-row{padding:5px 8px;border-left:3px solid transparent;border-radius:2px;margin:2px 0;font-size:11px;display:flex;align-items:center;flex-wrap:wrap;gap:5px}
  .consent-row[data-impact="critical"]{border-color:var(--red);background:var(--bg-red)}
  .consent-row[data-impact="elevated"]{border-color:var(--amber);background:var(--bg-amber)}
  .consent-row[data-impact="low"]{border-color:var(--green);background:var(--bg-green)}
  .consent-row[data-impact="unknown"]{border-color:var(--border);background:rgba(255,255,255,.02)}
  .pill{font-family:var(--font-mono);font-size:9px;letter-spacing:.06em;padding:2px 6px;border-radius:2px;border:1px solid var(--border);color:var(--muted);text-transform:uppercase;flex-shrink:0}
  .scope-badge{font-family:var(--font-mono);font-weight:500;font-size:10px;border-radius:2px;padding:2px 6px;border:1px solid transparent;flex-shrink:0}
  .scope-critical{background:var(--bg-red);border-color:rgba(230,57,70,.4);color:#ff8a8f}
  .scope-elevated{background:var(--bg-amber);border-color:rgba(255,183,3,.4);color:#ffe080}
  .scope-low{background:var(--bg-green);border-color:rgba(45,198,83,.4);color:#7fffaa}
  .scope-unknown{background:rgba(255,255,255,.04);border-color:var(--border);color:var(--muted)}
  .portal-link{display:inline-flex;align-items:center;gap:3px;font-size:10px;font-family:var(--font-mono);color:var(--blue);text-decoration:none;border:1px solid rgba(0,180,216,.3);border-radius:2px;padding:1px 5px;opacity:.85;transition:opacity .15s,border-color .15s;flex-shrink:0}
  .portal-link:hover{opacity:1;border-color:var(--blue);text-decoration:none}
  .btn-portal{display:inline-flex;align-items:center;gap:4px;font-size:10px;font-family:var(--font-mono);color:var(--blue);text-decoration:none;background:var(--bg-blue);border:1px solid rgba(0,180,216,.35);border-radius:2px;padding:4px 9px;cursor:pointer;transition:background .15s,border-color .15s;white-space:nowrap}
  .btn-portal:hover{background:rgba(0,180,216,.18);border-color:var(--blue)}
  .copy-btn{display:inline-flex;align-items:center;gap:3px;font-size:10px;font-family:var(--font-mono);color:var(--muted);background:transparent;border:1px solid var(--border);border-radius:2px;padding:1px 5px;cursor:pointer;transition:color .15s,border-color .15s;flex-shrink:0}
  .copy-btn:hover{color:var(--amber);border-color:var(--amber)}
  .footer{padding:20px;color:var(--muted);text-align:center;border-top:1px solid var(--border);font-size:11px;font-family:var(--font-mono)}
  .footer a{color:var(--blue)}
  .footer .sig{display:inline-flex;gap:8px;align-items:center}
  .icon{width:15px;height:15px;vertical-align:-3px;fill:var(--blue)}
</style>
</head>
<body>
<header>
  <div class="wrap">
    <div class="brand-row">
      <div>
        <div class="brand-name">Consentra<span class="dot">•</span></div>
        <div class="brand-sub">Entra ID — OAuth Consent Report</div>
      </div>
      <div class="brand-meta">
        <b id="tTotal">0</b> enterprise apps
        <span id="genAt"></span>
      </div>
    </div>
    <div class="impact-section" aria-label="Permission impact summary">
      <div class="impact-heading">Permission impact by vendor</div>
      <div class="impact-grid" id="impactGrid"></div>
    </div>
    <div class="filter-row">
      <span class="filter-label">Risk</span>
      <span id="chipCritical" class="chip chip-red" role="button" tabindex="0" aria-pressed="false"><span class="dot-s dot-red"></span><b id="tFail">0</b> Critical scopes</span>
      <span id="chipElevated" class="chip chip-amber" role="button" tabindex="0" aria-pressed="false"><span class="dot-s dot-amber"></span><b id="tWarn">0</b> Elevated scopes</span>
      <span id="chipLow" class="chip chip-green" role="button" tabindex="0" aria-pressed="false"><span class="dot-s dot-green"></span><b id="tPass">0</b> Low impact only</span>
    </div>
    <div class="filter-row">
      <span class="filter-label">Type</span>
      <span id="chipEnterprise" class="chip chip-blue" role="button" tabindex="0" aria-pressed="false"><span class="dot-s dot-mixed"></span><b id="tEnterprise">0</b> Enterprise</span>
      <span id="chipMicrosoft" class="chip chip-blue" role="button" tabindex="0" aria-pressed="false"><span class="dot-s dot-blue"></span><b id="tMicrosoft">0</b> Microsoft</span>
      <span id="chipThirdParty" class="chip chip-amber" role="button" tabindex="0" aria-pressed="false"><span class="dot-s dot-amber"></span><b id="tThirdParty">0</b> 3rd Party</span>
      <span id="chipHomeTenant" class="chip chip-green" role="button" tabindex="0" aria-pressed="false"><span class="dot-s dot-green"></span><b id="tHomeTenant">0</b> Home Tenant</span>
      <span class="filter-sep"></span>
      <span id="chipStale" class="chip chip-amber" role="button" tabindex="0" aria-pressed="false"><span class="dot-s dot-amber"></span><b id="tStale">0</b> Stale</span>
      <span id="chipHideNoGrants" class="chip chip-toggle" role="button" tabindex="0" aria-pressed="false"><span class="dot-s dot-grey"></span>Hide no-grant apps</span>
    </div>
    <div class="filter-row">
      <span class="filter-label">Find</span>
      <input id="q" type="search" placeholder="Search app name…" />
      <span id="chipHasNotes" class="chip chip-toggle" role="button" tabindex="0" aria-pressed="false" title="Show only apps with notes">Has notes</span>
      <span id="chipNoNotes" class="chip chip-toggle" role="button" tabindex="0" aria-pressed="false" title="Show only apps without notes">No notes</span>
      <span class="filter-sep"></span>
      <button id="sortToggle" type="button" aria-pressed="false" title="Sort apps alphabetically">Sort A→Z</button>
      <button id="exportCsv" title="Export filtered apps to CSV">Export CSV</button>
    </div>
  </div>
</header>
<div class="list-counter" aria-live="polite"><span style="font-size:9px;letter-spacing:.1em;color:var(--muted2)">SHOWING</span> <b id="visibleCount">0</b> apps</div>
<div id="list"></div>
<div class="footer">
  <span class="sig">
    <a href="https://whennotif.com" target="_blank" rel="noopener" title="whennotif•">whennotif<span style="color:var(--red)">•</span></a>
    — created by <strong>Jürgen Waldl</strong>
    <a href="https://www.linkedin.com/in/j%C3%BCrgen-waldl-6592837b/" target="_blank" rel="noopener" title="LinkedIn">
      <svg class="icon" viewBox="0 0 24 24" aria-hidden="true">
        <path d="M4.98 3.5C4.98 4.88 3.86 6 2.5 6S0 4.88 0 3.5 1.12 1 2.5 1 4.98 2.12 4.98 3.5zM0 8h5v16H0zM8 8h4.8v2.2h.07c.67-1.2 2.3-2.46 4.73-2.46 5.06 0 6 3.33 6 7.66V24h-5V16.4c0-1.8-.03-4.1-2.5-4.1-2.5 0-2.88 1.95-2.88 3.97V24H8z"/>
      </svg>
    </a>
  </span>
  <div>generated — <span id="genAt2"></span></div>
</div>
<script id="reportData" type="application/json">__REPORT_JSON__</script>
<script>
const report = JSON.parse(document.getElementById("reportData").textContent);

// --- Entra portal URL helper ---
const _tenantId = report.tenantId || "";
const _portalBase = _tenantId ? `https://entra.microsoft.com/${_tenantId}` : "https://entra.microsoft.com";
function entraPermUrl(spObjectId, appId){
    let url = `${_portalBase}/#view/Microsoft_AAD_IAM/ManagedAppMenuBlade/~/Permissions/objectId/${encodeURIComponent(spObjectId)}`;
    if (appId) url += `/appId/${encodeURIComponent(appId)}`;
    return url;
}
function entraUserConsentUrl(principalId, grantId, spObjectId){
    return `https://portal.azure.com/#view/Microsoft_AAD_IAM/AssignmentDetailBlade/objectId/${encodeURIComponent(principalId)}/isUser~/true/assignmentObjectId/${encodeURIComponent(grantId)}/spObjectId/${encodeURIComponent(spObjectId)}/principalId/${encodeURIComponent(principalId)}`;
}

// --- State & Elements ---
const state = { q:"", status:"all", consent:"any", vendor:"any", stale:false, hideNoGrants:false, notes:"any", sort:"asc" };
const $q = document.getElementById("q");
const $list = document.getElementById("list");
const $impactGrid = document.getElementById("impactGrid");

// Utils & impact metadata
function escapeHtml(s){ return String(s).replace(/[&<>"']/g,m=>({ "&":"&amp;","<":"&lt;",">":"&gt;","\"":"&quot;","'":"&#039;"}[m])) }
function asArray(x){ return Array.isArray(x) ? x : (x ? [x] : []); }
const scopeImpactMeta = {
    critical: { key:"critical", className:"scope-critical", label:"Critical" },
    elevated: { key:"elevated", className:"scope-elevated", label:"Elevated" },
    low: { key:"low", className:"scope-low", label:"Low impact" },
    unknown: { key:"unknown", className:"scope-unknown", label:"Unknown" }
};
function getImpactMeta(impact){
    const key = (impact || "elevated").toLowerCase();
    return scopeImpactMeta[key] || scopeImpactMeta.elevated;
}
const vendorImpactMeta = {
    microsoft: { title: "Microsoft apps" },
    thirdParty: { title: "3rd party apps" },
    homeTenant: { title: "Home tenant apps" }
};
function renderImpactGrid(){
    if(!$impactGrid) return;
    const summary = report.impactSummary || {};
    const cards = Object.entries(vendorImpactMeta).map(([vendorKey, meta])=>{
        const counts = summary[vendorKey] || {};
        const critical = Number(counts.critical) || 0;
        const elevated = Number(counts.elevated) || 0;
        const low = Number(counts.low) || 0;
        const total = critical + elevated + low;
        const max = Math.max(critical, elevated, low, 1);
        const rows = [
            { label: "High", level:"critical", value: critical },
            { label: "Elevated", level:"elevated", value: elevated },
            { label: "Low", level:"low", value: low }
        ].map(row => {
            const width = max === 0 ? 0 : Math.round((row.value / max) * 100);
            return `<div class="impact-bar" data-level="${row.level}"><span class="impact-bar-label">${row.label}</span><div class="impact-bar-track"><div class="impact-bar-fill" style="width:${width}%"></div></div><span class="impact-count">${row.value}</span></div>`;
        }).join("");
        const totalLabel = total === 1 ? 'permission' : 'permissions';
        return `<div class="impact-card" data-vendor="${vendorKey}"><h5>${meta.title}</h5><div class="impact-total">${total} ${totalLabel}</div><div class="impact-bars">${rows}</div></div>`;
    }).join("");
    $impactGrid.innerHTML = cards;
}

// Chips
const $chipCritical = document.getElementById("chipCritical");
const $chipElevated = document.getElementById("chipElevated");
const $chipLow = document.getElementById("chipLow");
const $chipEnterprise = document.getElementById("chipEnterprise");
const $chipMicrosoft = document.getElementById("chipMicrosoft");
const $chipThirdParty = document.getElementById("chipThirdParty");
const $chipHomeTenant = document.getElementById("chipHomeTenant");
const $chipStale = document.getElementById("chipStale");
const $visibleCount = document.getElementById("visibleCount");
const $sortToggle = document.getElementById("sortToggle");

// Header totals
document.getElementById("tTotal").textContent = report.totals.total;
document.getElementById("tPass").textContent  = report.totals.pass;
document.getElementById("tWarn").textContent  = report.totals.warn;
document.getElementById("tFail").textContent  = report.totals.fail;
document.getElementById("tEnterprise").textContent = report.totals.enterprise || 0;
document.getElementById("tMicrosoft").textContent = report.totals.microsoft;
document.getElementById("tThirdParty").textContent = report.totals.thirdParty;
document.getElementById("tHomeTenant").textContent = report.totals.homeTenant;
document.getElementById("tStale").textContent = report.totals.stale || 0;
document.getElementById("genAt").textContent  = report.generatedAt;
const _genAt2 = document.getElementById("genAt2"); if(_genAt2) _genAt2.textContent = report.generatedAt;

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
function toggleStatusFilter(filter){
  state.status = state.status === filter ? "all" : filter;
  updateStatusChips();
  render();
}
function updateStatusChips(){
  [
    [$chipCritical,"fail"],
    [$chipElevated,"warn"],
    [$chipLow,"pass"]
  ].forEach(([chip,value])=>{
    if(!chip) return;
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
function toggleStaleFilter(){
        state.stale = !state.stale;
        updateStaleChip();
        render();
}
function updateVendorChips(){
  const vendorMapping = [
    [$chipEnterprise, "enterprise"],
    [$chipMicrosoft, "microsoft"],
    [$chipThirdParty, "thirdparty"],
    [$chipHomeTenant, "home"]
  ];
  vendorMapping.forEach(([chip,value])=>{
    if(!chip) return;
    const active = state.vendor === value;
    chip.classList.toggle("active", active);
    chip.setAttribute("aria-pressed", String(active));
  });
}
function updateStaleChip(){
    if(!$chipStale) return;
    $chipStale.classList.toggle("active", state.stale);
    $chipStale.setAttribute("aria-pressed", String(state.stale));
}
const $chipHideNoGrants = document.getElementById("chipHideNoGrants");
const $chipHasNotes = document.getElementById("chipHasNotes");
const $chipNoNotes = document.getElementById("chipNoNotes");
function toggleHideNoGrants(){
    state.hideNoGrants = !state.hideNoGrants;
    if($chipHideNoGrants){ $chipHideNoGrants.classList.toggle("active", state.hideNoGrants); $chipHideNoGrants.setAttribute("aria-pressed", String(state.hideNoGrants)); }
    render();
}
function toggleNotesFilter(val){
    state.notes = (state.notes === val) ? "any" : val;
    updateNotesChips();
    render();
}
function updateNotesChips(){
    if($chipHasNotes){ $chipHasNotes.classList.toggle("active", state.notes === "has"); $chipHasNotes.setAttribute("aria-pressed", String(state.notes === "has")); }
    if($chipNoNotes){ $chipNoNotes.classList.toggle("active", state.notes === "none"); $chipNoNotes.setAttribute("aria-pressed", String(state.notes === "none")); }
}
$chipCritical.addEventListener("click", ()=> toggleStatusFilter("fail"));
$chipElevated.addEventListener("click", ()=> toggleStatusFilter("warn"));
$chipLow.addEventListener("click", ()=> toggleStatusFilter("pass"));
$chipEnterprise.addEventListener("click", ()=> toggleVendorFilter("enterprise"));
$chipMicrosoft.addEventListener("click", ()=> toggleVendorFilter("microsoft"));
$chipThirdParty.addEventListener("click", ()=> toggleVendorFilter("thirdparty"));
$chipHomeTenant.addEventListener("click", ()=> toggleVendorFilter("home"));
$chipStale.addEventListener("click", toggleStaleFilter);
if($chipHideNoGrants) $chipHideNoGrants.addEventListener("click", toggleHideNoGrants);
if($chipHasNotes) $chipHasNotes.addEventListener("click", ()=> toggleNotesFilter("has"));
if($chipNoNotes) $chipNoNotes.addEventListener("click", ()=> toggleNotesFilter("none"));
[$chipCritical,$chipElevated,$chipLow,$chipEnterprise,$chipMicrosoft,$chipThirdParty,$chipHomeTenant,$chipStale,$chipHideNoGrants,$chipHasNotes,$chipNoNotes].forEach(chip=>{
  if(!chip) return;
  chip.addEventListener("keydown", event=>{
    if(event.key === "Enter" || event.key === " "){
      event.preventDefault();
      chip.click();
    }
  });
});
updateStatusChips();
updateVendorChips();
updateStaleChip();
updateNotesChips();
renderImpactGrid();

function toggleConsentFilter(consent){
  state.consent = (state.consent === consent) ? "any" : consent;
  render();
}

$list.addEventListener("click", (event)=>{
    const tag = event.target.closest('.tag[data-consent], .tag[data-vendor], .tag[data-stale]');
  if(!tag) return;
  const consent = tag.getAttribute('data-consent');
  const vendor = tag.getAttribute('data-vendor');
    const stale = tag.getAttribute('data-stale');
  if(consent){
    toggleConsentFilter(consent);
  } else if(vendor){
    toggleVendorFilter(vendor);
    } else if(stale){
        toggleStaleFilter();
  }
});

$list.addEventListener("keydown", (event)=>{
    const tag = event.target.closest('.tag[data-consent], .tag[data-vendor], .tag[data-stale]');
  if(!tag) return;
  if(event.key === "Enter" || event.key === " "){
    event.preventDefault();
    const consent = tag.getAttribute('data-consent');
    const vendor = tag.getAttribute('data-vendor');
        const stale = tag.getAttribute('data-stale');
    if(consent){
      toggleConsentFilter(consent);
    } else if(vendor){
      toggleVendorFilter(vendor);
        } else if(stale){
            toggleStaleFilter();
    }
  }
});

function tagHtml(t){
  const label = escapeHtml(t);
  const lower = String(t).toLowerCase();
  const isAdmin = lower === "admin consent";
  const isUser  = lower === "user consent";
  const isMicrosoft = lower === "microsoft";
  const isThirdParty = lower === "3rd party";
  const isHomeTenant = lower === "home tenant";
  const isHidden = lower === "hidden";
    const isStale = lower === "staleapps";
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
  if (isHomeTenant){
    const active = state.vendor === "home";
    const classes = ["tag","tag-home","tag-clickable"];
    if(active){ classes.push("tag-active"); }
    return `<span class="${classes.join(" ")}" data-vendor="home" role="button" tabindex="0" aria-pressed="${active}">${label}</span>`;
  }
  if (isHidden){
    return `<span class="tag tag-hidden">${label}</span>`;
  }
    if (isStale){
        const classes = ["tag","tag-stale","tag-clickable"];
        if(state.stale){ classes.push("tag-active"); }
        return `<span class="${classes.join(" ")}" data-stale="true" role="button" tabindex="0" aria-pressed="${state.stale}">${label}</span>`;
    }
  return `<span class="tag">${label}</span>`;
}
function filterItems(items){
  let out = items;
  if(state.status !== "all"){ out = out.filter(x => (x.status||"") === state.status); }
  if(state.consent === "admin"){ out = out.filter(x => (x.adminGrants||[]).length > 0); }
  if(state.consent === "user"){ out = out.filter(x => (x.userGrants||[]).length > 0); }
  if(state.vendor === "enterprise"){ out = out.filter(x => !x.isMicrosoft); }
  if(state.vendor === "microsoft"){ out = out.filter(x => x.isMicrosoft); }
  if(state.vendor === "thirdparty"){ out = out.filter(x => x.isThirdParty); }
  if(state.vendor === "home"){ out = out.filter(x => x.isHomeTenant); }
  if(state.stale){ out = out.filter(x => (x.tags||[]).some(t => String(t).toLowerCase() === "staleapps")); }
  if(state.hideNoGrants){ out = out.filter(x => (x.adminGrants||[]).length > 0 || (x.userGrants||[]).length > 0); }
  if(state.notes === "has"){ out = out.filter(x => x.notes && String(x.notes).trim()); }
  if(state.notes === "none"){ out = out.filter(x => !x.notes || !String(x.notes).trim()); }
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
        const _spId = (item.clientId && !item.clientId.startsWith("app:")) ? item.clientId : null;
        const _permUrl = (_spId && item.clientAppId) ? entraPermUrl(_spId, item.clientAppId) : null;
        const adminRows = (item.adminGrants||[]).map(a=>{
            const impact = getImpactMeta(a.impact);
            const scopeBadge = `<span class="scope-badge ${impact.className}" title="${impact.label}">${escapeHtml(a.scope)}</span>`;
            const rowTitle = impact.label ? ` title="${impact.label} scope"` : "";
            const revokeLink = _permUrl ? `<a class="portal-link" href="${_permUrl}" target="_blank" rel="noopener" title="Revoke in Entra portal">↗ Revoke</a>` : "";
            const copyBtn = a.grantId ? `<button class="copy-btn" onclick="copyPsRevoke('${escapeHtml(a.grantId)}')" title="Copy PowerShell command to revoke this grant&#10;Remove-MgOauth2PermissionGrant -OAuth2PermissionGrantId '${escapeHtml(a.grantId)}'">⧉ Copy PS</button>` : "";
            return `<div class="consent-row" data-impact="${impact.key}"${rowTitle}><span class="pill">Admin</span> ${escapeHtml(a.resource)} • ${scopeBadge}${revokeLink}${copyBtn}</div>`;
        }).join("");
        const userRows  = (item.userGrants ||[]).map(u=>{
            const impact = getImpactMeta(u.impact);
            const scopeBadge = `<span class="scope-badge ${impact.className}" title="${impact.label}">${escapeHtml(u.scope)}</span>`;
            const rowTitle = impact.label ? ` title="${impact.label} scope"` : "";
            const details = asArray(u.userDetails);
            const usersHtml = details.length
                ? details.map(d => {
                    const directUrl = (d.principalId && d.assignmentId && _spId)
                        ? entraUserConsentUrl(d.principalId, d.assignmentId, _spId)
                        : null;
                    const copyBtn = d.grantId ? `<button class="copy-btn" onclick="copyPsRevoke('${escapeHtml(d.grantId)}')" title="Copy PowerShell command to revoke this user's grant&#10;Remove-MgOauth2PermissionGrant -OAuth2PermissionGrantId '${escapeHtml(d.grantId)}'">⧉ Copy PS</button>` : "";
                    const userLabel = directUrl
                        ? `<a class="portal-link" href="${directUrl}" target="_blank" rel="noopener" title="Remove assignment for this user in Azure portal">${escapeHtml(d.upn||d.principalId)} ↗</a>`
                        : `<span>${escapeHtml(d.upn||d.principalId)}</span>`;
                    return userLabel + copyBtn;
                }).join(" ")
                : escapeHtml(asArray(u.users).join(", "));
            const overviewLink = _permUrl ? `<a class="portal-link" href="${_permUrl}" target="_blank" rel="noopener" title="All permissions overview in Entra portal">Permissions ↗</a>` : "";
            return `<div class="consent-row" data-impact="${impact.key}"${rowTitle}><span class="pill">User</span> ${escapeHtml(u.resource)} • ${scopeBadge} • ${usersHtml} ${overviewLink}</div>`;
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
    const vendorClass = info.IsMicrosoft ? "Microsoft" : (info.IsThirdParty ? "3rd Party" : (info.IsHomeTenant ? "Home tenant" : "Custom / unknown"));
    const vendorClassEsc = escapeHtml(vendorClass);
    const verifiedPublisherRaw = info.VerifiedPublisherName || info.Publisher || "—";
    const verifiedPublisherEsc = escapeHtml(verifiedPublisherRaw);
    const ownerOrgEsc = info.OwnerOrg ? escapeHtml(info.OwnerOrg) : "—";
    const spTagsEsc = (info.Tags && info.Tags.length) ? escapeHtml(info.Tags.join(", ")) : "—";
    const publisherDomainEsc = info.PublisherDomain ? escapeHtml(info.PublisherDomain) : "—";
    const spNames = Array.isArray(info.ServicePrincipalNames) ? info.ServicePrincipalNames : [];
    const spNameDisplay = spNames.length ? escapeHtml(spNames.slice(0,3).join(", ")) + (spNames.length > 3 ? " …" : "") : "—";
        const notes = item.notes ? escapeHtml(item.notes) : "—";
        const ssoState = item.ssoState ? escapeHtml(item.ssoState) : "—";
    const metadataBlock = `
        <div class="sec-title">App metadata</div>
        <div class="meta-grid">
          <div><span class="meta-label">Vendor type</span><span class="meta-value">${vendorClassEsc}</span></div>
          <div><span class="meta-label">Verified publisher</span><span class="meta-value">${verifiedPublisherEsc}</span></div>
          <div><span class="meta-label">Publisher domain</span><span class="meta-value">${publisherDomainEsc}</span></div>
          <div><span class="meta-label">App owner org</span><span class="meta-value">${ownerOrgEsc}</span></div>
          <div><span class="meta-label">Service principal names</span><span class="meta-value">${spNameDisplay}</span></div>
          <div><span class="meta-label">SP tags</span><span class="meta-value">${spTagsEsc}</span></div>
                    <div><span class="meta-label">SSO state</span><span class="meta-value">${ssoState}</span></div>
                    <div><span class="meta-label">Notes</span><span class="meta-value">${notes}</span></div>
        </div>
        <div class="divider"></div>`;
        return `
            <div class="card status-${item.status}">
                <div class="item">
                    <div>
                        <div class="app-name">${escapeHtml(item.clientName||"(no name)")}</div>
                        <div class="app-id">${escapeHtml(item.clientAppId||"")} · ${scopeArr.length} scope(s)</div>
                        <div class="tags">${tags}</div>
                    </div>
                    <div class="actions">
                        ${_permUrl ? `<a class="btn-portal" href="${_permUrl}" target="_blank" rel="noopener" title="Manage permissions in Entra portal">Entra ↗</a>` : ""}
                        <button onclick='copyJson(${JSON.stringify(item).replace(/</g,"\\u003c")})' title="Copy JSON">⧉ JSON</button>
                    </div>
                </div>
                <details>
                  <summary>Details &amp; Permissions</summary>
                  <div style="padding:10px 14px">
                    ${metadataBlock}${sections}${none}
                  </div>
                </details>
            </div>`;
  }).join("");
}

function copyJson(obj){ navigator.clipboard.writeText(JSON.stringify(obj,null,2)); }
function copyPsRevoke(grantId){ navigator.clipboard.writeText(`Remove-MgOauth2PermissionGrant -OAuth2PermissionGrantId '${grantId}'`); }

// Export CSV (current filter selection)
document.getElementById("exportCsv").addEventListener("click", () => {
  const items = filterItems(report.items);
        const headers = ["clientName","clientAppId","status","isMicrosoft","isThirdParty","isHomeTenant","isHidden","staleApp","ssoState","notes","tags","adminScopes","userScopes","userCount"];
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
      isHomeTenant: x.isHomeTenant ? "true" : "false",
      isHidden: x.isHidden ? "true" : "false",
    staleApp: x.staleApp ? "true" : "false",
            ssoState: x.ssoState || "",
            notes: x.notes || "",
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

$jsonSafe = $json.Replace('</script>','<\/script>')
$html = $htmlTemplate.Replace('__REPORT_JSON__', $jsonSafe)

$ts = Get-Date -Format "yyyyMMdd_HHmm"
$htmlPath = Join-Path $outDir ("OAuth2_Consent_Report_$ts.html")
Set-Content -Path $htmlPath -Encoding UTF8 -Value $html

Complete-Progress

Write-Host "`nDone:" -ForegroundColor Cyan
Write-Host "  CSV : $csvPath"
Write-Host "  HTML: $htmlPath"
