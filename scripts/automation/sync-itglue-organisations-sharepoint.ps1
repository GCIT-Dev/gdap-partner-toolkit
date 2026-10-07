#Requires -Version 7.4
#Requires -Modules MspGdap

<#
.SYNOPSIS
    Keeps a SharePoint list in your partner tenant in step with your IT Glue organisations.
.DESCRIPTION
    Reads every organisation from the IT Glue API and compares it with a SharePoint list (by
    default 'ITGlue Org Register' on your root site) in your own partner tenant, through Microsoft
    Graph with MspGdap (Invoke-MspGraphRequest -PartnerTenant). Each list item holds the
    organisation's name (Title), short name, IT Glue ID and, with -MatchCustomers, the Microsoft
    tenant ID of the GDAP customer with the same name.

    Without -Apply the script only reports what it would create, update or remove. With -Apply it
    creates missing items and updates changed ones. Items for organisations that no longer exist
    in IT Glue are only removed with -RemoveStale. The list itself is only created with
    -CreateList.

    No secrets are stored in the script. The IT Glue API key is read at run time from a
    SecretManagement vault (Get-Secret), which can be a local SecretStore or Azure Key Vault. Graph
    access uses the technician's MspGdap token. For a scheduled run, host the script on Azure
    Functions 4.x with PowerShell 7.6 and follow docs/08 (Pattern B), or give a separate app the
    Sites.Selected permission on just this site.

    This script works in your partner tenant only, so it takes -SiteId instead of -TenantId or
    -AllCustomers. Each result row still carries CustomerTenantId (when matched) and CustomerName.
.PARAMETER ITGlueApiKeySecretName
    Name of the secret in your SecretManagement vault that holds the IT Glue API key.
.PARAMETER VaultName
    SecretManagement vault that holds the IT Glue API key. Defaults to your default vault.
.PARAMETER ITGlueBaseUri
    IT Glue API address for your data centre.
.PARAMETER SiteId
    Microsoft Graph site ID of the SharePoint site in your partner tenant. Defaults to root.
.PARAMETER ListName
    Display name of the SharePoint list.
.PARAMETER MatchCustomers
    Fill the CustomerTenantId column by matching each organisation name to Get-MspCustomer.
.PARAMETER CreateList
    Create the list (with ShortName, ITGlueID and CustomerTenantId columns) if it doesn't exist. Needs -Apply.
.PARAMETER RemoveStale
    Remove list items whose IT Glue organisation no longer exists. Needs -Apply.
.PARAMETER Apply
    Create and update list items. Without this switch the script only reports.
    Combine with -WhatIf to preview the changes.
.PARAMETER OutputPath
    Optional path of a CSV file for the results.
.EXAMPLE
    ./sync-itglue-organisations-sharepoint.ps1 -ITGlueApiKeySecretName 'ITGlueApiKey' -MatchCustomers

    Reports how the list differs from IT Glue, including which organisations match a GDAP customer.
.EXAMPLE
    ./sync-itglue-organisations-sharepoint.ps1 -ITGlueApiKeySecretName 'ITGlueApiKey' -ITGlueBaseUri 'https://api.au.itglue.com' -CreateList -Apply -WhatIf

    Shows what would be created in a new list for an Australian IT Glue account, without changing anything.
.NOTES
    Replaces the original 2019 method: an Azure Functions 1.x experimental PowerShell timer
    function, a client secret app created with the AzureAD module, a client credentials token
    from the v1.0 endpoint (oauth2/token with resource=), and the IT Glue API key and client secret
    pasted into the script or kept in plain app settings.
    Required GDAP roles: none. The script works in your partner tenant, where the technician needs
    edit rights on the SharePoint site.
    Required partner app permissions: Microsoft Graph Sites.ReadWrite.All (delegated) for list items,
    and Sites.Manage.All (delegated) for -CreateList, which Sites.FullControl.All in
    partner-app.full.json covers. With -MatchCustomers, the partner tenant reads that
    Get-MspCustomer needs.
.LINK
    https://gcit.com.au/knowledge-base/sync-it-glue-organisations-with-a-sharepoint-list-via-powershell/
.LINK
    ../../docs/07-migrating-from-dap-msonline.md
.LINK
    ../../docs/08-unattended-automation.md
.LINK
    https://learn.microsoft.com/en-us/graph/api/listitem-create?view=graph-rest-1.0
#>
[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium')]
[OutputType([pscustomobject])]
param(
    [Parameter(Mandatory)]
    [ValidateNotNullOrEmpty()]
    [string]$ITGlueApiKeySecretName,

    [string]$VaultName,

    [ValidateSet('https://api.itglue.com', 'https://api.eu.itglue.com', 'https://api.au.itglue.com')]
    [string]$ITGlueBaseUri = 'https://api.itglue.com',

    [ValidateNotNullOrEmpty()]
    [string]$SiteId = 'root',

    [ValidateNotNullOrEmpty()]
    [string]$ListName = 'ITGlue Org Register',

    [switch]$MatchCustomers,

    [switch]$CreateList,

    [switch]$RemoveStale,

    [switch]$Apply,

    [string]$OutputPath
)

$ErrorActionPreference = 'Stop'
$results = [System.Collections.Generic.List[object]]::new()

function ConvertTo-ResultRow {
    param([System.Collections.IDictionary]$Values = @{})
    $row = [ordered]@{}
    foreach ($column in 'CustomerTenantId', 'CustomerName', 'Status', 'ITGlueId', 'ShortName', 'ListItemId', 'Action', 'Detail') {
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

$itGlueHeaders = @{ 'x-api-key' = (Get-SecretText -Name $ITGlueApiKeySecretName -Vault $VaultName); Accept = 'application/vnd.api+json' }
try {
    Write-Verbose 'Reading IT Glue organisations'
    $organisations = @(Get-ITGlueCollection -Path 'organizations?page[size]=1000' -Headers $itGlueHeaders -BaseUri $ITGlueBaseUri)
}
finally {
    $itGlueHeaders = $null
}

$customerByName = @{}
if ($MatchCustomers) {
    foreach ($customer in @(Get-MspCustomer -IncludeGdapStatus)) {
        if ($customer.DisplayName) { $customerByName[([string]$customer.DisplayName).Trim().ToLowerInvariant()] = [string]$customer.TenantId }
    }
}

$escapedName = $ListName -replace "'", "''"
$list = Invoke-MspGraphRequest -PartnerTenant -Method GET -Uri "v1.0/sites/$SiteId/lists?`$filter=displayName eq '$escapedName'&`$select=id,displayName" | Select-Object -First 1
$existingItems = @()

if (-not $list) {
    if ($Apply -and $CreateList) {
        if ($PSCmdlet.ShouldProcess("partner tenant site $SiteId", "Create SharePoint list '$ListName'")) {
            $listBody = @{
                displayName = $ListName
                columns     = @(
                    @{ name = 'ShortName'; text = @{} }
                    @{ name = 'ITGlueID'; number = @{}; indexed = $true }
                    @{ name = 'CustomerTenantId'; text = @{} }
                )
                list        = @{ template = 'genericList' }
            }
            $list = Invoke-MspGraphRequest -PartnerTenant -Method POST -Uri "v1.0/sites/$SiteId/lists" -Body $listBody -Confirm:$false
        }
    }
    elseif (-not $CreateList) {
        Write-Warning "List '$ListName' was not found on site $SiteId. Use -CreateList -Apply to create it. Reporting what would be added."
    }
}
else {
    $itemUri = "v1.0/sites/$SiteId/lists/$($list.id)/items?`$expand=fields(`$select=Title,ShortName,ITGlueID,CustomerTenantId)&`$top=999"
    $existingItems = @(Invoke-MspGraphRequest -PartnerTenant -Method GET -Uri $itemUri)
}

$itemsById = @{}
foreach ($item in $existingItems) {
    if ($null -ne $item.fields.ITGlueID) { $itemsById[[string][int64]$item.fields.ITGlueID] = $item }
}

foreach ($organisation in $organisations) {
    $name = [string]$organisation.attributes.name
    $shortName = [string]$organisation.attributes.'short-name'
    $tenant = if ($customerByName.ContainsKey($name.Trim().ToLowerInvariant())) { $customerByName[$name.Trim().ToLowerInvariant()] } else { $null }
    $values = @{ CustomerTenantId = $tenant; CustomerName = $name; ITGlueId = [string]$organisation.id; ShortName = $shortName; Status = 'Succeeded' }

    try {
        $fields = [ordered]@{ Title = $name; ShortName = $shortName; ITGlueID = [int64]$organisation.id }
        if ($MatchCustomers) { $fields.CustomerTenantId = $tenant }
        $existing = $itemsById[[string][int64]$organisation.id]

        if (-not $existing) {
            $values.Action = 'Create'
        }
        else {
            $values.ListItemId = [string]$existing.id
            $changed = ([string]$existing.fields.Title -ne $name) -or ([string]$existing.fields.ShortName -ne $shortName)
            if ($MatchCustomers -and [string]$existing.fields.CustomerTenantId -ne [string]$tenant) { $changed = $true }
            $values.Action = if ($changed) { 'Update' } else { 'None' }
        }

        if ($values.Action -ne 'None') {
            if (-not $Apply) {
                $values.Action = "$($values.Action) (report only)"
            }
            elseif (-not $list) {
                $values.Action = "$($values.Action) (list missing)"
            }
            elseif ($PSCmdlet.ShouldProcess("SharePoint list '$ListName'", "$($values.Action) item for $name")) {
                if ($values.Action -eq 'Create') {
                    $created = Invoke-MspGraphRequest -PartnerTenant -Method POST -Uri "v1.0/sites/$SiteId/lists/$($list.id)/items" -Body @{ fields = $fields } -Confirm:$false
                    $values.ListItemId = [string]$created.id
                    $values.Action = 'Created'
                }
                else {
                    $null = Invoke-MspGraphRequest -PartnerTenant -Method PATCH -Uri "v1.0/sites/$SiteId/lists/$($list.id)/items/$($existing.id)/fields" -Body $fields -Confirm:$false
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

$currentIds = @($organisations | ForEach-Object { [string][int64]$_.id })
foreach ($key in @($itemsById.Keys)) {
    if ($currentIds -contains $key) { continue }
    $item = $itemsById[$key]
    $values = @{ CustomerTenantId = [string]$item.fields.CustomerTenantId; CustomerName = [string]$item.fields.Title; ITGlueId = $key; ListItemId = [string]$item.id; Status = 'Succeeded'; Action = 'Stale (report only)' }
    try {
        if ($Apply -and $RemoveStale) {
            if ($PSCmdlet.ShouldProcess("SharePoint list '$ListName'", "Remove stale item $($item.fields.Title)")) {
                $null = Invoke-MspGraphRequest -PartnerTenant -Method DELETE -Uri "v1.0/sites/$SiteId/lists/$($list.id)/items/$($item.id)" -Confirm:$false
                $values.Action = 'Removed'
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

if ($OutputPath -and $results.Count -gt 0) {
    $results | Export-Csv -LiteralPath $OutputPath -NoTypeInformation -Encoding utf8
    Write-Verbose "Saved $($results.Count) rows to $OutputPath"
}
