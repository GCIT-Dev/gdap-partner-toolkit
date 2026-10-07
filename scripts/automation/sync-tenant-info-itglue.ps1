#Requires -Version 7.4
#Requires -Modules MspGdap

<#
.SYNOPSIS
    Collects Microsoft 365 tenant details for customers and optionally writes them to IT Glue flexible assets.

.DESCRIPTION
    For each customer the script reads, through Microsoft Graph:

    - the organisation name, initial domain and verified domains (GET /organization),
    - licence counts (GET /subscribedSkus), and
    - licensed users with their aliases and licences (GET /users).

    It returns one row per customer. With -ITGlueFlexibleAssetTypeId and -Apply it also creates or updates
    one flexible asset per customer in IT Glue, with the same traits as the original (tenant-name,
    tenant-id, initial-domain, verified-domains, licenses, licensed-users). The asset is matched on the
    tenant-id trait.

    Customers are matched to IT Glue organisations with -ITGlueOrganizationMapPath (a CSV with TenantId and
    OrganizationId columns). Without a map, the script matches verified domains against the email domains
    of IT Glue contacts, as the original did.

    The IT Glue API key is read from a SecretManagement vault at run time. It is never stored in the
    script. IT Glue also has its own Microsoft 365 integration, which may be all you need.

    The TenantInfoSync Azure Function in the functions folder runs the script on a timer.

.PARAMETER TenantId
    One or more customer tenant IDs (GUID) or verified domains. Accepts pipeline input, including the
    output of Get-MspCustomer.

.PARAMETER AllCustomers
    Process every customer returned by Get-MspCustomer.

.PARAMETER ITGlueFlexibleAssetTypeId
    ID of the "Office 365" flexible asset type in IT Glue. Without it nothing is sent to IT Glue.

.PARAMETER ITGlueApiKeySecretName
    Name of the SecretManagement secret that holds the IT Glue API key.

.PARAMETER ITGlueVaultName
    SecretManagement vault that holds the IT Glue API key. Defaults to the default vault.

.PARAMETER ITGlueBaseUri
    IT Glue API address for your region.

.PARAMETER ITGlueOrganizationMapPath
    CSV with TenantId and OrganizationId columns that maps customer tenants to IT Glue organisations.

.PARAMETER Apply
    Create or update the IT Glue flexible assets. Without it the script only reports what it would do.

.PARAMETER OutputPath
    Optional path for a CSV export of the results.

.EXAMPLE
    ./sync-tenant-info-itglue.ps1 -AllCustomers -OutputPath ./tenant-info.csv

    Collects domains, licences and licensed users for every customer without touching IT Glue.

.EXAMPLE
    ./sync-tenant-info-itglue.ps1 -AllCustomers -ITGlueFlexibleAssetTypeId 12345 -ITGlueApiKeySecretName 'ITGlueApiKey' -ITGlueOrganizationMapPath ./itglue-map.csv -Apply -WhatIf

    Shows which IT Glue flexible assets would be created or updated.

.NOTES
    Replaces the original 2018 method: MSOnline and DAP (Connect-MsolService -Credential,
    Get-MsolPartnerContract, Get-MsolCompanyInformation, Get-MsolDomain, Get-MsolAccountSku, Get-MsolUser)
    run interactively or from an Azure Functions v1 timer function with an AES-encrypted stored password,
    with the IT Glue API key pasted into the script.
    Required GDAP roles: Global Reader.
    Required partner app permissions: Microsoft Graph delegated User.Read (organisation), User.Read.All
    (users, User.ReadWrite.All is in the full manifest), LicenseAssignment.Read.All or
    Directory.ReadWrite.All (subscribedSkus).

.LINK
    https://gcit.com.au/knowledge-base/sync-office-365-tenant-info-itglue/

.LINK
    docs/07-migrating-from-dap-msonline.md

.LINK
    docs/08-unattended-automation.md
#>
[CmdletBinding(SupportsShouldProcess, DefaultParameterSetName = 'Tenant')]
[OutputType([pscustomobject])]
param(
    [Parameter(Mandatory, ParameterSetName = 'Tenant', ValueFromPipeline, ValueFromPipelineByPropertyName)]
    [Alias('CustomerTenantId')]
    [ValidateNotNullOrEmpty()]
    [string[]]$TenantId,

    [Parameter(Mandatory, ParameterSetName = 'All')]
    [switch]$AllCustomers,

    [ValidateRange(1, [long]::MaxValue)]
    [long]$ITGlueFlexibleAssetTypeId,

    [string]$ITGlueApiKeySecretName = 'ITGlueApiKey',

    [string]$ITGlueVaultName,

    [ValidateSet('https://api.itglue.com', 'https://api.eu.itglue.com', 'https://api.au.itglue.com')]
    [string]$ITGlueBaseUri = 'https://api.itglue.com',

    [string]$ITGlueOrganizationMapPath,

    [switch]$Apply,

    [string]$OutputPath
)

begin {
    $requested = [System.Collections.Generic.List[string]]::new()
    $results = [System.Collections.Generic.List[object]]::new()
    $itGlueHeaders = $null
    $itGlueBase = $ITGlueBaseUri.TrimEnd('/')
    $organisationMap = @{}
    $contactDomains = $null

    function ConvertTo-ResultRow {
        param($Customer, $Info, $OrganisationId, $Action, $ErrorMessage)
        [pscustomobject][ordered]@{
            CustomerTenantId     = $Customer.TenantId
            CustomerName         = $Customer.Name
            InitialDomain        = $Info.InitialDomain
            VerifiedDomains      = ($Info.Domains -join ', ')
            LicenceCount         = $Info.LicenceCount
            LicensedUserCount    = $Info.LicensedUserCount
            ITGlueOrganizationId = $OrganisationId
            Action               = $Action
            Error                = $ErrorMessage
        }
    }

    function ConvertTo-HtmlCell {
        param([object]$Value)
        [System.Net.WebUtility]::HtmlEncode([string]$Value)
    }

    function Invoke-ITGlueRequest {
        param([string]$Method, [string]$Path, [object]$Body)
        # Only ever send the API key to the configured IT Glue host, including for paging links.
        $uri = if ($Path.StartsWith("$itGlueBase/", [System.StringComparison]::OrdinalIgnoreCase)) {
            $Path
        }
        elseif ($Path -match '^[a-z]+://') {
            throw "Refusing to send the IT Glue API key to $Path."
        }
        else {
            "$itGlueBase/$Path"
        }
        if ($Method -eq 'GET') {
            $items = [System.Collections.Generic.List[object]]::new()
            while ($uri) {
                $page = Invoke-RestMethod -Method GET -Uri $uri -Headers $itGlueHeaders -ContentType 'application/vnd.api+json'
                foreach ($item in @($page.data)) { $items.Add($item) }
                $next = if ($page.links -and $page.links.next) { [string]$page.links.next } else { $null }
                if ($next -and -not $next.StartsWith("$itGlueBase/", [System.StringComparison]::OrdinalIgnoreCase)) {
                    throw "Refusing to follow an IT Glue paging link to another host: $next"
                }
                $uri = $next
            }
            return $items
        }
        $json = ConvertTo-Json -InputObject $Body -Depth 10
        Invoke-RestMethod -Method $Method -Uri $uri -Headers $itGlueHeaders -ContentType 'application/vnd.api+json' -Body $json
    }
}

process {
    foreach ($id in @($TenantId)) {
        if ($id) { $requested.Add($id) }
    }
}

end {
    if ($ITGlueFlexibleAssetTypeId) {
        if (-not (Get-Command -Name Get-Secret -ErrorAction SilentlyContinue)) {
            throw 'IT Glue sync needs Microsoft.PowerShell.SecretManagement and a vault that holds the API key.'
        }
        $secretParams = @{ Name = $ITGlueApiKeySecretName; AsPlainText = $true; ErrorAction = 'Stop' }
        if ($ITGlueVaultName) { $secretParams.Vault = $ITGlueVaultName }
        $itGlueHeaders = @{ 'x-api-key' = (Get-Secret @secretParams) }
        if ($ITGlueOrganizationMapPath) {
            Import-Csv -Path $ITGlueOrganizationMapPath | ForEach-Object { $organisationMap[([string]$_.TenantId).ToLowerInvariant()] = [string]$_.OrganizationId }
        }
    }

    $customers = if ($AllCustomers) {
        @(Get-MspCustomer | ForEach-Object { [pscustomobject]@{ TenantId = $_.TenantId; Name = $_.DisplayName } })
    }
    else {
        @($requested | Select-Object -Unique | ForEach-Object { [pscustomobject]@{ TenantId = $_; Name = $null } })
    }

    foreach ($customer in $customers) {
        try {
            $organisation = Invoke-MspGraphRequest -TenantId $customer.TenantId -Uri 'organization?$select=id,displayName,verifiedDomains' | Select-Object -First 1
            if (-not $customer.Name) { $customer.Name = $organisation.displayName }
            if ($organisation.id) { $customer.TenantId = $organisation.id }

            $domains = @($organisation.verifiedDomains | ForEach-Object { [string]$_.name } | Sort-Object)
            $initialDomain = @($organisation.verifiedDomains | Where-Object { $_.isInitial } | ForEach-Object { [string]$_.name }) | Select-Object -First 1

            $skuNames = @{}
            $licenceRows = foreach ($sku in @(Invoke-MspGraphRequest -TenantId $customer.TenantId -Uri 'subscribedSkus')) {
                $skuNames[[string]$sku.skuId] = $sku.skuPartNumber
                $enabled = [int]$sku.prepaidUnits.enabled
                $consumed = [int]$sku.consumedUnits
                '<tr><td>{0}</td><td>{1}</td><td>{2}</td><td>{3}</td></tr>' -f (ConvertTo-HtmlCell $sku.skuPartNumber), $enabled, $consumed, ([math]::Max(0, $enabled - $consumed))
            }
            $users = @(Invoke-MspGraphRequest -TenantId $customer.TenantId -Uri 'users?$select=displayName,userPrincipalName,proxyAddresses,assignedLicenses&$top=999' |
                    Where-Object { @($_.assignedLicenses).Count -gt 0 } | Sort-Object -Property userPrincipalName)
            $userRows = foreach ($user in $users) {
                $aliases = @($user.proxyAddresses | Where-Object { $_ -clike 'smtp:*' -and $_ -notmatch '\.onmicrosoft\.com$' } | ForEach-Object { ConvertTo-HtmlCell $_.Substring(5) }) -join '<br/>'
                $licenceNames = @($user.assignedLicenses | ForEach-Object { ConvertTo-HtmlCell $skuNames[[string]$_.skuId] }) -join '<br/>'
                '<tr><td>{0}</td><td><strong>{1}</strong><br/>{2}</td><td>{3}</td></tr>' -f (ConvertTo-HtmlCell $user.displayName), (ConvertTo-HtmlCell $user.userPrincipalName), $aliases, $licenceNames
            }

            $info = [pscustomobject]@{
                InitialDomain     = $initialDomain
                Domains           = $domains
                LicenceCount      = @($licenceRows).Count
                LicensedUserCount = $users.Count
                LicenceTable      = '<table class="table table-bordered table-hover"><thead><tr><th>Licence</th><th>Active</th><th>Consumed</th><th>Unused</th></tr></thead><tbody>' + (@($licenceRows) -join '') + '</tbody></table>'
                UserTable         = '<table class="table table-bordered table-hover"><thead><tr><th>Display name</th><th>Addresses</th><th>Assigned licences</th></tr></thead><tbody>' + (@($userRows) -join '') + '</tbody></table>'
            }

            $organisationId = $null
            $action = 'Collected'
            if ($ITGlueFlexibleAssetTypeId) {
                $organisationId = $organisationMap[([string]$customer.TenantId).ToLowerInvariant()]
                if (-not $organisationId -and -not $ITGlueOrganizationMapPath) {
                    if ($null -eq $contactDomains) {
                        $contactDomains = @{}
                        foreach ($contact in (Invoke-ITGlueRequest -Method GET -Path 'contacts?page[size]=1000')) {
                            foreach ($email in @($contact.attributes.'contact-emails')) {
                                $emailDomain = ([string]$email.value -split '@')[-1].ToLowerInvariant()
                                if ($emailDomain -and -not $contactDomains.ContainsKey($emailDomain)) { $contactDomains[$emailDomain] = [string]$contact.attributes.'organization-id' }
                            }
                        }
                    }
                    $organisationId = @($domains | ForEach-Object { $contactDomains[$_.ToLowerInvariant()] } | Where-Object { $_ }) | Select-Object -First 1
                }

                if (-not $organisationId) {
                    $action = 'NoITGlueMatch'
                }
                else {
                    $traits = [ordered]@{
                        'tenant-name'      = $customer.Name
                        'tenant-id'        = $customer.TenantId
                        'initial-domain'   = $initialDomain
                        'verified-domains' = ($domains -join ', ')
                        'licenses'         = $info.LicenceTable
                        'licensed-users'   = $info.UserTable
                    }
                    $existing = @(Invoke-ITGlueRequest -Method GET -Path "flexible_assets?filter[organization_id]=$organisationId&filter[flexible_asset_type_id]=$ITGlueFlexibleAssetTypeId") |
                        Where-Object { [string]$_.attributes.traits.'tenant-id' -eq [string]$customer.TenantId } | Select-Object -First 1
                    $verb = if ($existing) { 'Update' } else { 'Create' }
                    if (-not $Apply) {
                        $action = "Would$verb"
                    }
                    elseif ($PSCmdlet.ShouldProcess("IT Glue organisation $organisationId", "$verb Office 365 flexible asset for $($customer.Name)")) {
                        if ($existing) {
                            $body = @{ data = @{ type = 'flexible-assets'; attributes = @{ traits = $traits } } }
                            $null = Invoke-ITGlueRequest -Method PATCH -Path "flexible_assets/$($existing.id)" -Body $body
                            $action = 'Updated'
                        }
                        else {
                            $body = @{ data = @{ type = 'flexible-assets'; attributes = @{ 'organization-id' = $organisationId; 'flexible-asset-type-id' = $ITGlueFlexibleAssetTypeId; traits = $traits } } }
                            $null = Invoke-ITGlueRequest -Method POST -Path 'flexible_assets' -Body $body
                            $action = 'Created'
                        }
                    }
                    else {
                        $action = 'WhatIf'
                    }
                }
            }

            $row = ConvertTo-ResultRow -Customer $customer -Info $info -OrganisationId $organisationId -Action $action
            $results.Add($row)
            $row
        }
        catch {
            Write-Warning "Customer $($customer.TenantId): $($_.Exception.Message)"
            $row = ConvertTo-ResultRow -Customer $customer -Action 'Failed' -ErrorMessage $_.Exception.Message
            $results.Add($row)
            $row
        }
    }

    if ($OutputPath) {
        $results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding utf8
    }
}
