#Requires -Version 7.4
#Requires -Modules MspGdap

<#
.SYNOPSIS
    Collects Microsoft Secure Score by category for customers and optionally writes it to IT Glue flexible
    assets.

.DESCRIPTION
    For each customer the script reads the latest Secure Score (GET /security/secureScores?$top=1) and the
    control profiles (GET /security/secureScoreControlProfiles) through Microsoft Graph. It returns one row
    per customer with the overall score and the score for each category (Identity, Data, Device, Apps,
    Infrastructure).

    With -ITGlueFlexibleAssetTypeId and -Apply it also creates or updates one "Microsoft Secure Score"
    flexible asset per customer, with the same traits as the original (overview, per-category control tables
    and scores, tenant-name, tenant-id, default-domain). The asset is matched on the tenant-id trait.

    Customers are matched to IT Glue organisations with -ITGlueOrganizationMapPath (a CSV with TenantId and
    OrganizationId columns, which replaces the original's SharePoint match register). Without a map, the
    script matches verified domains against the email domains of IT Glue contacts, as the original did.

    The IT Glue API key is read from a SecretManagement vault at run time and is never stored in the script.

.PARAMETER TenantId
    One or more customer tenant IDs (GUID) or verified domains. Accepts pipeline input, including the
    output of Get-MspCustomer.

.PARAMETER AllCustomers
    Process every customer returned by Get-MspCustomer.

.PARAMETER ITGlueFlexibleAssetTypeId
    ID of the "Microsoft Secure Score" flexible asset type in IT Glue. Without it nothing is sent to IT Glue.

.PARAMETER ITGlueOrganizationMapPath
    CSV with TenantId and OrganizationId columns that maps customer tenants to IT Glue organisations.
    Without it, customers are matched by the email domains of IT Glue contacts.

.PARAMETER ITGlueApiKeySecretName
    Name of the SecretManagement secret that holds the IT Glue API key.

.PARAMETER ITGlueVaultName
    SecretManagement vault that holds the IT Glue API key. Defaults to the default vault.

.PARAMETER ITGlueBaseUri
    IT Glue API address for your region.

.PARAMETER Apply
    Create or update the IT Glue flexible assets. Without it the script only reports what it would do.

.PARAMETER OutputPath
    Optional path for a CSV export of the results.

.EXAMPLE
    ./sync-secure-score-itglue.ps1 -AllCustomers -OutputPath ./secure-score-categories.csv

    Collects Secure Score by category for every customer without touching IT Glue.

.EXAMPLE
    ./sync-secure-score-itglue.ps1 -AllCustomers -ITGlueFlexibleAssetTypeId 12345 -ITGlueOrganizationMapPath ./itglue-map.csv -Apply -WhatIf

    Shows which IT Glue flexible assets would be created or updated.

.NOTES
    Replaces the original 2018 method: an AdminAgents (DAP) multi-tenant app created with the AzureAD module,
    a hard-coded client secret and IT Glue API key, the v1 token endpoint with resource=, the Graph contracts
    list and beta Secure Score endpoints.
    Required GDAP roles: Security Reader (or Global Reader).
    Required partner app permissions: Microsoft Graph delegated SecurityEvents.Read.All
    (SecurityEvents.ReadWrite.All is in the full manifest) and User.Read (organisation).

.LINK
    https://gcit.com.au/knowledge-base/sync-microsoft-secure-scores-with-it-glue/

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

    [string]$ITGlueOrganizationMapPath,

    [string]$ITGlueApiKeySecretName = 'ITGlueApiKey',

    [string]$ITGlueVaultName,

    [ValidateSet('https://api.itglue.com', 'https://api.eu.itglue.com', 'https://api.au.itglue.com')]
    [string]$ITGlueBaseUri = 'https://api.itglue.com',

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
    $categories = @('Identity', 'Data', 'Device', 'Apps', 'Infrastructure')

    function ConvertTo-HtmlCell {
        param([object]$Value)
        [System.Net.WebUtility]::HtmlEncode([string]$Value)
    }

    function ConvertTo-ResultRow {
        param($Customer, $Summary, $OrganisationId, $Action, $ErrorMessage)
        [pscustomobject][ordered]@{
            CustomerTenantId     = $Customer.TenantId
            CustomerName         = $Customer.Name
            DefaultDomain        = $Summary.DefaultDomain
            SecureScore          = $Summary.Score
            MaxScore             = $Summary.MaxScore
            IdentityScore        = $Summary.Scores['Identity']
            DataScore            = $Summary.Scores['Data']
            DeviceScore          = $Summary.Scores['Device']
            AppsScore            = $Summary.Scores['Apps']
            InfrastructureScore  = $Summary.Scores['Infrastructure']
            ITGlueOrganizationId = $OrganisationId
            Action               = $Action
            Error                = $ErrorMessage
        }
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
            $defaultDomain = @($organisation.verifiedDomains | Where-Object { $_.isDefault } | ForEach-Object { [string]$_.name }) | Select-Object -First 1

            $score = Invoke-MspGraphRequest -TenantId $customer.TenantId -Uri 'security/secureScores?$top=1' -MaxPages 1 | Select-Object -First 1
            if (-not $score) { throw 'No Secure Score was returned for this tenant.' }
            $profiles = @{}
            Invoke-MspGraphRequest -TenantId $customer.TenantId -Uri 'security/secureScoreControlProfiles' |
                ForEach-Object { $profiles[[string]$_.id] = $_ }

            $categoryScores = @{}
            $categoryTables = @{}
            foreach ($category in $categories) {
                $controls = @($score.controlScores | Where-Object { [string]$_.controlCategory -eq $category })
                $categoryScores[$category] = [math]::Round((($controls | Measure-Object -Property score -Sum).Sum), 2)
                $rows = foreach ($control in ($controls | Sort-Object -Property { $profiles[[string]$_.controlName].rank })) {
                    $controlProfile = $profiles[[string]$control.controlName]
                    $title = if ($controlProfile.title) { $controlProfile.title } else { $control.controlName }
                    '<tr><td>{0}</td><td>{1} of {2}</td><td>{3}</td></tr>' -f (ConvertTo-HtmlCell $title), [math]::Round([double]$control.score, 2), (ConvertTo-HtmlCell $controlProfile.maxScore), (ConvertTo-HtmlCell $controlProfile.userImpact)
                }
                $categoryTables[$category] = '<table class="table table-bordered table-hover"><thead><tr><th>Control</th><th>Score</th><th>User impact</th></tr></thead><tbody>' + (@($rows) -join '') + '</tbody></table>'
            }
            $summary = [pscustomobject]@{
                DefaultDomain = $defaultDomain
                Score         = [math]::Round([double]$score.currentScore, 2)
                MaxScore      = [math]::Round([double]$score.maxScore, 2)
                Scores        = $categoryScores
            }

            $organisationId = $null
            $action = 'Collected'
            if ($ITGlueFlexibleAssetTypeId) {
                $organisationId = $organisationMap[([string]$customer.TenantId).ToLowerInvariant()]
                if (-not $organisationId -and -not $ITGlueOrganizationMapPath) {
                    if ($null -eq $contactDomains) {
                        $contactDomains = @{}
                        foreach ($contact in (Invoke-ITGlueRequest -Method GET -Path 'contacts?page[size]=1000')) {
                            foreach ($contactEmail in @($contact.attributes.'contact-emails')) {
                                $emailDomain = ([string]$contactEmail.value -split '@')[-1].ToLowerInvariant()
                                if ($emailDomain -and -not $contactDomains.ContainsKey($emailDomain)) { $contactDomains[$emailDomain] = [string]$contact.attributes.'organization-id' }
                            }
                        }
                    }
                    $domainNames = @($organisation.verifiedDomains | ForEach-Object { ([string]$_.name).ToLowerInvariant() })
                    $organisationId = @($domainNames | ForEach-Object { $contactDomains[$_] } | Where-Object { $_ }) | Select-Object -First 1
                }
                if (-not $organisationId) {
                    $action = 'NoITGlueMatch'
                }
                else {
                    $overview = '<p>Secure Score {0} of {1} ({2}%), measured {3}.</p>' -f $summary.Score, $summary.MaxScore, $(if ($summary.MaxScore) { [math]::Round(100 * $summary.Score / $summary.MaxScore, 1) } else { 0 }), (ConvertTo-HtmlCell $score.createdDateTime)
                    $traits = [ordered]@{
                        'overview'                = $overview
                        'identity-controls'       = $categoryTables['Identity']
                        'data-controls'           = $categoryTables['Data']
                        'device-controls'         = $categoryTables['Device']
                        'apps-controls'           = $categoryTables['Apps']
                        'infrastructure-controls' = $categoryTables['Infrastructure']
                        'tenant-name'             = $customer.Name
                        'secure-score'            = [int]$summary.Score
                        'identity-score'          = [int]$categoryScores['Identity']
                        'data-score'              = [int]$categoryScores['Data']
                        'device-score'            = [int]$categoryScores['Device']
                        'apps-score'              = [int]$categoryScores['Apps']
                        'infrastructure-score'    = [int]$categoryScores['Infrastructure']
                        'tenant-id'               = $customer.TenantId
                        'default-domain'          = $defaultDomain
                    }
                    $existing = @(Invoke-ITGlueRequest -Method GET -Path "flexible_assets?filter[organization_id]=$organisationId&filter[flexible_asset_type_id]=$ITGlueFlexibleAssetTypeId") |
                        Where-Object { [string]$_.attributes.traits.'tenant-id' -eq [string]$customer.TenantId } | Select-Object -First 1
                    $verb = if ($existing) { 'Update' } else { 'Create' }
                    if (-not $Apply) {
                        $action = "Would$verb"
                    }
                    elseif ($PSCmdlet.ShouldProcess("IT Glue organisation $organisationId", "$verb Secure Score flexible asset for $($customer.Name)")) {
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

            $row = ConvertTo-ResultRow -Customer $customer -Summary $summary -OrganisationId $organisationId -Action $action
            $results.Add($row)
            $row
        }
        catch {
            Write-Warning "Customer $($customer.TenantId): $($_.Exception.Message)"
            $row = ConvertTo-ResultRow -Customer $customer -Summary ([pscustomobject]@{ Scores = @{} }) -Action 'Failed' -ErrorMessage $_.Exception.Message
            $results.Add($row)
            $row
        }
    }

    if ($OutputPath) {
        $results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding utf8
    }
}
