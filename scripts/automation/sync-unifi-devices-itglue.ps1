#Requires -Version 7.4
#Requires -Modules MspGdap

<#
.SYNOPSIS
    Documents UniFi network devices as IT Glue configurations, using your SharePoint site match register.
.DESCRIPTION
    For each UniFi site the script:
      1. looks up the IT Glue organisation in the SharePoint match register (by default
         'UniFi - IT Glue match register', whose ITGlue lookup column points at the
         'ITGlue Org Register' list kept by sync-itglue-organisations-sharepoint.ps1), read through
         Microsoft Graph with MspGdap in your own partner tenant,
      2. reads the site's devices from the official UniFi Network integration API with an API key
         (no controller username or password), and
      3. compares them, by MAC address, with the organisation's IT Glue configurations.

    Without -Apply the script only reports which configurations it would create or update. With
    -Apply it creates missing configurations. Existing configurations are left alone unless you
    also use -UpdateExisting, which then updates their name and primary IP address. This matches
    the original, which only overwrote existing configurations when you turned that on, because
    technicians often rename configurations or add their own details in IT Glue.

    Secrets are read at run time from a SecretManagement vault (Get-Secret): the UniFi API key
    (created in UniFi Network under Settings, Control Plane, Integrations) and the IT Glue API key.
    Nothing is written to disk.

    This rewrite covers device configurations only. The original also built a UniFi Site flexible
    asset with Wi-Fi, LAN and WAN, port forward, VPN and alarm tables from the legacy controller
    API. Those endpoints are not part of the official integration API, and IT Glue's own UniFi
    integration now documents them. New configurations also don't get the manufacturer, model or
    serial number IDs that the original created in IT Glue. The UniFi model is written to the
    configuration notes instead. The script that built the match register by comparing site names
    and client MAC addresses is not rewritten, so add or fix matches in the SharePoint list by hand.
    Check the integration API version on your console, because UniFi may change it.

    This script works in your partner tenant and your UniFi console only, so it takes -SiteId
    instead of -TenantId or -AllCustomers. CustomerName in each row is the IT Glue organisation.
.PARAMETER UniFiBaseUri
    Base address of the UniFi Network integration API, for example
    https://unifi.contoso.com/proxy/network/integration on a UniFi OS console. Must be HTTPS.
.PARAMETER UniFiApiKeySecretName
    Name of the secret that holds the UniFi Network API key.
.PARAMETER ITGlueApiKeySecretName
    Name of the secret that holds the IT Glue API key.
.PARAMETER VaultName
    SecretManagement vault that holds both secrets. Defaults to your default vault.
.PARAMETER ITGlueBaseUri
    IT Glue API address for your data centre.
.PARAMETER ConfigurationTypeId
    IT Glue configuration type ID for new configurations, for example your 'Network Device' type.
.PARAMETER ConfigurationStatusId
    IT Glue configuration status ID for new configurations, for example your 'Active' status.
.PARAMETER SiteId
    Microsoft Graph site ID of the SharePoint site in your partner tenant. Defaults to root.
.PARAMETER MatchListName
    Display name of the UniFi to IT Glue match register list.
.PARAMETER OrgListName
    Display name of the IT Glue organisation register list.
.PARAMETER UpdateExisting
    Also update the name and primary IP address of existing configurations that differ from
    UniFi. Without this switch existing configurations are reported but never changed.
.PARAMETER Apply
    Create IT Glue configurations (and update existing ones with -UpdateExisting). Without this
    switch the script only reports. Combine with -WhatIf to preview the changes.
.PARAMETER OutputPath
    Optional path of a CSV file for the results.
.EXAMPLE
    ./sync-unifi-devices-itglue.ps1 -UniFiBaseUri 'https://unifi.contoso.com/proxy/network/integration' -UniFiApiKeySecretName 'UniFiApiKey' -ITGlueApiKeySecretName 'ITGlueApiKey' -ConfigurationTypeId 101 -ConfigurationStatusId 201

    Reports which UniFi devices are missing or out of date in IT Glue.
.EXAMPLE
    ./sync-unifi-devices-itglue.ps1 -UniFiBaseUri 'https://unifi.contoso.com/proxy/network/integration' -UniFiApiKeySecretName 'UniFiApiKey' -ITGlueApiKeySecretName 'ITGlueApiKey' -ConfigurationTypeId 101 -ConfigurationStatusId 201 -UpdateExisting -Apply -WhatIf

    Shows which configurations would be created, and which existing ones would be updated, without changing anything.
.NOTES
    Replaces the original 2019 method: the legacy UniFi controller API with an admin username and
    password (POST /api/login on port 8443), a client secret app created with the AzureAD module,
    a client credentials token from the v1.0 endpoint (oauth2/token with resource=), and the IT
    Glue API key, client secret and UniFi password pasted into the script.
    Required GDAP roles: none. The script works in your partner tenant, where the technician needs
    read access to the SharePoint site.
    Required partner app permissions: Microsoft Graph Sites.ReadWrite.All (delegated, read only used,
    Sites.Read.All is the least privileged).
.LINK
    https://gcit.com.au/knowledge-base/sync-unifi-sites-with-it-glue/
.LINK
    ../../docs/07-migrating-from-dap-msonline.md
.LINK
    ../../docs/08-unattended-automation.md
.LINK
    https://learn.microsoft.com/en-us/graph/api/listitem-list?view=graph-rest-1.0
#>
[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium')]
[OutputType([pscustomobject])]
param(
    [Parameter(Mandatory)]
    [ValidatePattern('^https://[^\s/]+(/[^\s]*)?$')]
    [string]$UniFiBaseUri,

    [Parameter(Mandatory)]
    [ValidateNotNullOrEmpty()]
    [string]$UniFiApiKeySecretName,

    [Parameter(Mandatory)]
    [ValidateNotNullOrEmpty()]
    [string]$ITGlueApiKeySecretName,

    [string]$VaultName,

    [ValidateSet('https://api.itglue.com', 'https://api.eu.itglue.com', 'https://api.au.itglue.com')]
    [string]$ITGlueBaseUri = 'https://api.itglue.com',

    [Parameter(Mandatory)]
    [ValidateRange('Positive')]
    [long]$ConfigurationTypeId,

    [Parameter(Mandatory)]
    [ValidateRange('Positive')]
    [long]$ConfigurationStatusId,

    [ValidateNotNullOrEmpty()]
    [string]$SiteId = 'root',

    [ValidateNotNullOrEmpty()]
    [string]$MatchListName = 'UniFi - IT Glue match register',

    [ValidateNotNullOrEmpty()]
    [string]$OrgListName = 'ITGlue Org Register',

    [switch]$UpdateExisting,

    [switch]$Apply,

    [string]$OutputPath
)

$ErrorActionPreference = 'Stop'
$results = [System.Collections.Generic.List[object]]::new()
$UniFiBaseUri = $UniFiBaseUri.TrimEnd('/')

function ConvertTo-ResultRow {
    param([System.Collections.IDictionary]$Values = @{})
    $row = [ordered]@{}
    foreach ($column in 'CustomerTenantId', 'CustomerName', 'Status', 'UniFiSite', 'ITGlueOrganisationId', 'DeviceName', 'Model', 'MacAddress', 'IpAddress', 'ConfigurationId', 'Action', 'Detail') {
        $row[$column] = if ($Values.Contains($column)) { $Values[$column] } else { $null }
    }
    [pscustomobject]$row
}

function Get-SecretText {
    param([Parameter(Mandatory)][string]$Name, [string]$Vault)
    $secretParams = @{ Name = $Name }
    if ($Vault) { $secretParams.Vault = $Vault }
    $secret = Get-Secret @secretParams
    if ($secret -is [securestring]) { [System.Net.NetworkCredential]::new('', $secret).Password } else { [string]$secret }
}

function Get-ITGlueCollection {
    # Follows JSON:API links.next, but only within the IT Glue API host, so the key never goes elsewhere.
    param([Parameter(Mandatory)][string]$Path, [Parameter(Mandatory)][hashtable]$Headers, [Parameter(Mandatory)][string]$BaseUri)
    $next = "$BaseUri/$Path"
    while ($next) {
        if (-not $next.StartsWith("$BaseUri/", [System.StringComparison]::OrdinalIgnoreCase)) {
            throw "Refusing to follow an IT Glue link outside $BaseUri."
        }
        $page = Invoke-RestMethod -Method GET -Uri $next -Headers $Headers -ContentType 'application/vnd.api+json'
        foreach ($item in @($page.data)) { $item }
        $next = if ($page.links -and $page.links.next) { [string]$page.links.next } else { $null }
    }
}

function Get-UniFiCollection {
    # The integration API pages with offset and limit and returns totalCount.
    param([Parameter(Mandatory)][string]$Path, [Parameter(Mandatory)][hashtable]$Headers, [Parameter(Mandatory)][string]$BaseUri)
    $offset = 0
    do {
        $separator = if ($Path.Contains('?')) { '&' } else { '?' }
        $page = Invoke-RestMethod -Method GET -Uri "$BaseUri/$Path$($separator)offset=$offset&limit=200" -Headers $Headers
        $data = @($page.data)
        foreach ($item in $data) { $item }
        $offset += $data.Count
        $total = if ($null -ne $page.totalCount) { [int]$page.totalCount } else { $offset }
    } while ($data.Count -gt 0 -and $offset -lt $total)
}

function Get-ListItem {
    param([Parameter(Mandatory)][string]$Name, [Parameter(Mandatory)][string]$Site)
    $escaped = $Name -replace "'", "''"
    $list = Invoke-MspGraphRequest -PartnerTenant -Method GET -Uri "v1.0/sites/$Site/lists?`$filter=displayName eq '$escaped'&`$select=id,displayName" | Select-Object -First 1
    if (-not $list) { throw "SharePoint list '$Name' was not found on site $Site." }
    @(Invoke-MspGraphRequest -PartnerTenant -Method GET -Uri "v1.0/sites/$Site/lists/$($list.id)/items?`$expand=fields&`$top=999")
}

function ConvertTo-MacKey {
    param([string]$Mac)
    if (-not $Mac) { return $null }
    ($Mac -replace '[^0-9A-Fa-f]', '').ToLowerInvariant()
}

# Map UniFi site references to IT Glue organisations through the two SharePoint lists.
$orgItems = @{}
foreach ($item in Get-ListItem -Name $OrgListName -Site $SiteId) { $orgItems[[string]$item.id] = $item }
$siteMap = @{}
foreach ($item in Get-ListItem -Name $MatchListName -Site $SiteId) {
    $lookupId = [string]$item.fields.ITGlueLookupId
    if (-not $item.fields.UnifiSiteName -or -not $lookupId -or -not $orgItems.ContainsKey($lookupId)) { continue }
    $org = $orgItems[$lookupId]
    $siteMap[([string]$item.fields.UnifiSiteName).ToLowerInvariant()] = [pscustomobject]@{
        OrganisationId   = [string][int64]$org.fields.ITGlueID
        OrganisationName = [string]$org.fields.Title
        CustomerTenantId = [string]$org.fields.CustomerTenantId
    }
}

$uniFiHeaders = @{ 'X-API-KEY' = (Get-SecretText -Name $UniFiApiKeySecretName -Vault $VaultName); Accept = 'application/json' }
$itGlueHeaders = @{ 'x-api-key' = (Get-SecretText -Name $ITGlueApiKeySecretName -Vault $VaultName); Accept = 'application/vnd.api+json' }

try {
    $sites = @(Get-UniFiCollection -Path 'v1/sites' -Headers $uniFiHeaders -BaseUri $UniFiBaseUri)
    foreach ($site in $sites) {
        $reference = if ($site.internalReference) { [string]$site.internalReference } else { [string]$site.name }
        $match = $siteMap[$reference.ToLowerInvariant()]
        if (-not $match) {
            $row = ConvertTo-ResultRow -Values @{ Status = 'Skipped'; UniFiSite = "$($site.name) ($reference)"; Action = 'Unmatched'; Detail = "Add this site to '$MatchListName' to sync it." }
            $results.Add($row)
            $row
            continue
        }

        try {
            $devices = @(Get-UniFiCollection -Path "v1/sites/$($site.id)/devices" -Headers $uniFiHeaders -BaseUri $UniFiBaseUri)
            $configurations = @(Get-ITGlueCollection -Path "configurations?filter[organization_id]=$($match.OrganisationId)&page[size]=1000" -Headers $itGlueHeaders -BaseUri $ITGlueBaseUri)
        }
        catch {
            $row = ConvertTo-ResultRow -Values @{ CustomerTenantId = $match.CustomerTenantId; CustomerName = $match.OrganisationName; Status = 'Failed'; UniFiSite = [string]$site.name; ITGlueOrganisationId = $match.OrganisationId; Detail = $_.Exception.Message }
            $results.Add($row)
            $row
            continue
        }

        $byMac = @{}
        foreach ($configuration in $configurations) {
            $key = ConvertTo-MacKey -Mac ([string]$configuration.attributes.'mac-address')
            if ($key -and -not $byMac.ContainsKey($key)) { $byMac[$key] = $configuration }
        }

        foreach ($device in $devices) {
            $deviceName = if ($device.name) { [string]$device.name } else { "UniFi $($device.model)" }
            $values = @{
                CustomerTenantId     = $match.CustomerTenantId
                CustomerName         = $match.OrganisationName
                Status               = 'Succeeded'
                UniFiSite            = [string]$site.name
                ITGlueOrganisationId = $match.OrganisationId
                DeviceName           = $deviceName
                Model                = [string]$device.model
                MacAddress           = [string]$device.macAddress
                IpAddress            = [string]$device.ipAddress
            }
            try {
                $existing = $byMac[(ConvertTo-MacKey -Mac ([string]$device.macAddress))]
                if ($existing) {
                    $values.ConfigurationId = [string]$existing.id
                    $same = ([string]$existing.attributes.name -eq $deviceName) -and ([string]$existing.attributes.'primary-ip' -eq [string]$device.ipAddress)
                    $values.Action = if ($same) { 'None' } elseif ($UpdateExisting) { 'Update' } else { 'Differs' }
                }
                else {
                    $values.Action = 'Create'
                }

                if ($values.Action -notin 'None', 'Differs') {
                    if (-not $Apply) {
                        $values.Action = "$($values.Action) (report only)"
                    }
                    elseif ($PSCmdlet.ShouldProcess("IT Glue organisation $($match.OrganisationName)", "$($values.Action) configuration $deviceName")) {
                        if ($values.Action -eq 'Create') {
                            $attributes = [ordered]@{
                                'organization-id'         = [long]$match.OrganisationId
                                'name'                    = $deviceName
                                'hostname'                = $deviceName
                                'primary-ip'              = [string]$device.ipAddress
                                'mac-address'             = [string]$device.macAddress
                                'configuration-type-id'   = $ConfigurationTypeId
                                'configuration-status-id' = $ConfigurationStatusId
                                'notes'                   = "Synced from UniFi site $($site.name). Model $($device.model)."
                            }
                            $body = @{ data = @{ type = 'configurations'; attributes = $attributes } } | ConvertTo-Json -Depth 5
                            $created = Invoke-RestMethod -Method POST -Uri "$ITGlueBaseUri/configurations" -Headers $itGlueHeaders -ContentType 'application/vnd.api+json' -Body $body
                            $values.ConfigurationId = [string]$created.data.id
                            $values.Action = 'Created'
                        }
                        else {
                            $attributes = [ordered]@{ 'name' = $deviceName; 'primary-ip' = [string]$device.ipAddress }
                            $body = @{ data = @{ type = 'configurations'; attributes = $attributes } } | ConvertTo-Json -Depth 5
                            $null = Invoke-RestMethod -Method PATCH -Uri "$ITGlueBaseUri/configurations/$($existing.id)" -Headers $itGlueHeaders -ContentType 'application/vnd.api+json' -Body $body
                            $values.Action = 'Updated'
                        }
                    }
                    else {
                        $values.Action = 'WhatIf'
                    }
                }
            }
            catch {
                $values.Status = 'Failed'
                $values.Detail = $_.Exception.Message
            }
            $row = ConvertTo-ResultRow -Values $values
            $results.Add($row)
            $row
        }
    }
}
finally {
    $uniFiHeaders = $null
    $itGlueHeaders = $null
}

if ($OutputPath -and $results.Count -gt 0) {
    $results | Export-Csv -LiteralPath $OutputPath -NoTypeInformation -Encoding utf8
    Write-Verbose "Saved $($results.Count) rows to $OutputPath"
}
