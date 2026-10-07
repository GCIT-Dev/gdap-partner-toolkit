#Requires -Version 7.4
#Requires -Modules MspGdap

<#
.SYNOPSIS
    Checks that you can connect to each GDAP customer through the partner app, and with which roles.
.DESCRIPTION
    This is the modern replacement for the Secure Application Model set-up check. The original
    article built a multi-tenant app with a client secret, pre-consented it by adding it to the
    AdminAgents group (DAP only), stored a refresh token and secret in plain text files, and
    swapped the refresh token for MSOnline and Azure AD Graph tokens.

    With MspGdap the partner app, its certificate and each technician's refresh token are set up
    once (docs/02 to docs/04). This script then proves, read-only, for each customer:
      1. GDAP: which roles you effectively hold through active relationships (Test-MspGdapAccess),
      2. Consent: that the partner app is consented with every expected scope
         (Test-MspPartnerAppConsent),
      3. Graph: that a delegated Microsoft Graph call reaches the right tenant
         (GET /organization), and
      4. Exchange (with -TestExchange): that Connect-MspExchangeOnline connects and
         Get-OrganizationConfig answers.

    The script changes nothing. Run it after onboarding a customer and before scheduling any of
    the other scripts against them.
.PARAMETER TenantId
    One or more customer tenant IDs (GUID) or verified domains. Accepts pipeline input.
.PARAMETER AllCustomers
    Run against every customer returned by Get-MspCustomer -IncludeGdapStatus.
.PARAMETER ManifestPath
    Optional manifest to compare the customer's consent against, for example
    ./manifests/partner-app.minimal.json. Defaults to the partner app's registered permissions.
.PARAMETER TestExchange
    Also connect to Exchange Online in each customer.
.PARAMETER OutputPath
    Optional path of a CSV file for the results.
.EXAMPLE
    ./test-delegated-tenant-access.ps1 -AllCustomers -OutputPath ./access-check.csv

    Checks GDAP roles, consent and Graph access for every customer and saves the results.
.EXAMPLE
    ./test-delegated-tenant-access.ps1 -TenantId 'contoso.onmicrosoft.com' -ManifestPath ./manifests/partner-app.minimal.json -TestExchange

    Checks one customer against the minimal manifest, including an Exchange Online connection.
.NOTES
    Replaces the original 2019 method: the Secure Application Model built with the AzureAD module
    (New-AzureADApplication, New-AzureADApplicationPasswordCredential, Add-AzureADGroupMember to
    AdminAgents), the PartnerCenter module (New-PartnerAccessToken), Azure AD Graph tokens from the
    v1.0 token endpoint, Connect-MsolService -AdGraphAccessToken, a client secret and refresh token
    in plain text files, and basic authentication for Exchange.
    Required GDAP roles: any role in the customer (Global Reader recommended). Exchange
    Administrator or Global Reader for -TestExchange.
    Required partner app permissions: Microsoft Graph DelegatedAdminRelationship.ReadWrite.All
    (partner tenant, read only used), Application.ReadWrite.All or Directory.ReadWrite.All
    (customer, read only used, to read the service principal and its grants) and
    Organization.Read.All or Directory.ReadWrite.All (customer, GET /organization), all delegated.
    Office 365 Exchange Online Exchange.Manage (delegated) for -TestExchange.
.LINK
    https://gcit.com.au/knowledge-base/how-to-connect-to-delegated-office-365-tenants-using-the-secure-app-model/
.LINK
    ../../docs/04-preconsent-customers.md
.LINK
    ../../docs/07-migrating-from-dap-msonline.md
.LINK
    https://learn.microsoft.com/en-us/partner-center/developer/gdap-and-secure-application-model
#>
[CmdletBinding(DefaultParameterSetName = 'Tenant')]
[OutputType([pscustomobject])]
param(
    [Parameter(Mandatory, ParameterSetName = 'Tenant', ValueFromPipeline, ValueFromPipelineByPropertyName)]
    [Alias('CustomerTenantId')]
    [ValidateNotNullOrEmpty()]
    [string[]]$TenantId,

    [Parameter(Mandatory, ParameterSetName = 'AllCustomers')]
    [switch]$AllCustomers,

    [ValidateScript({ Test-Path -LiteralPath $_ -PathType Leaf })]
    [string]$ManifestPath,

    [switch]$TestExchange,

    [string]$OutputPath
)

begin {
    $ErrorActionPreference = 'Stop'
    $requested = [System.Collections.Generic.List[string]]::new()
    $results = [System.Collections.Generic.List[object]]::new()
    $resultColumns = @(
        'Status', 'GdapStatus', 'GdapCheck', 'EffectiveRoles', 'ConsentCheck', 'MissingScopes',
        'GraphCheck', 'InitialDomain', 'ExchangeCheck', 'Detail'
    )

    function ConvertTo-ResultRow {
        param(
            [Parameter(Mandatory)][object]$Customer,
            [System.Collections.IDictionary]$Values = @{}
        )
        $row = [ordered]@{
            CustomerTenantId = [string]$Customer.TenantId
            CustomerName     = [string]$Customer.DisplayName
        }
        foreach ($column in $resultColumns) {
            $row[$column] = if ($Values.Contains($column)) { $Values[$column] } else { $null }
        }
        [pscustomobject]$row
    }

    function Get-TargetCustomer {
        param([switch]$All, [string[]]$Requested)
        if ($All) {
            return @(Get-MspCustomer -IncludeGdapStatus)
        }
        foreach ($id in $Requested) {
            $match = $null
            try {
                $match = Get-MspCustomer -TenantId $id -IncludeGdapStatus | Select-Object -First 1
            }
            catch {
                Write-Warning "Could not look up '$id' in your customer list: $($_.Exception.Message)"
            }
            if ($match) { $match } else { [pscustomobject]@{ TenantId = $id; DisplayName = $id } }
        }
    }

    function Get-StepSummary {
        # MspGdap Test-* results carry a Steps list. Summarise the ones that did not pass.
        param([object]$Result)
        @($Result.Steps | Where-Object { $_.Status -notin 'Passed', 'Changed', 'Skipped' } | ForEach-Object { "$($_.Step): $($_.Status) $($_.Detail)".Trim() }) -join ' | '
    }

    function Format-MissingScope {
        param([object]$Missing)
        if (-not $Missing) { return $null }
        $parts = foreach ($key in @($Missing.Keys)) {
            $scopes = @($Missing[$key]) -join ' '
            if ($scopes) { "$($key): $scopes" }
        }
        @($parts) -join ' | '
    }
}

process {
    foreach ($id in $TenantId) {
        if ($id) { $requested.Add($id) }
    }
}

end {
    foreach ($customer in @(Get-TargetCustomer -All:$AllCustomers -Requested $requested.ToArray())) {
        $values = @{ Status = 'Succeeded' }
        if ($customer.PSObject.Properties['GdapStatus']) { $values.GdapStatus = $customer.GdapStatus }
        $problems = [System.Collections.Generic.List[string]]::new()

        try {
            $gdap = Test-MspGdapAccess -TenantId $customer.TenantId
            $values.GdapCheck = if ($gdap.Success) { 'Passed' } else { 'Failed' }
            $values.EffectiveRoles = @($gdap.EffectiveRoleNames) -join ', '
            if (-not $gdap.Success) { $problems.Add("GDAP: $(Get-StepSummary -Result $gdap)") }
        }
        catch {
            $values.GdapCheck = 'Error'
            $problems.Add("GDAP: $($_.Exception.Message)")
        }

        try {
            $consentParams = @{ TenantId = $customer.TenantId }
            if ($ManifestPath) { $consentParams.ManifestPath = $ManifestPath }
            $consent = Test-MspPartnerAppConsent @consentParams
            $values.ConsentCheck = if ($consent.Success) { 'Passed' } else { 'Failed' }
            $values.MissingScopes = Format-MissingScope -Missing $consent.MissingScopes
            if (-not $consent.Success) { $problems.Add("Consent: $(Get-StepSummary -Result $consent)") }
        }
        catch {
            $values.ConsentCheck = 'Error'
            $problems.Add("Consent: $($_.Exception.Message)")
        }

        try {
            $organisation = Invoke-MspGraphRequest -TenantId $customer.TenantId -Method GET -Uri 'v1.0/organization?$select=id,displayName,verifiedDomains' | Select-Object -First 1
            if ([string]$organisation.id -and $customer.TenantId -match '^[0-9a-fA-F-]{36}$' -and [string]$organisation.id -ne [string]$customer.TenantId) {
                throw "Graph answered for tenant $($organisation.id), not $($customer.TenantId)."
            }
            $values.GraphCheck = 'Passed'
            $values.InitialDomain = @($organisation.verifiedDomains | Where-Object { $_.isInitial } | ForEach-Object { $_.name }) -join ', '
        }
        catch {
            $values.GraphCheck = 'Failed'
            $problems.Add("Graph: $($_.Exception.Message)")
        }

        if ($TestExchange) {
            try {
                $null = Connect-MspExchangeOnline -TenantId $customer.TenantId
                $null = Get-OrganizationConfig
                $values.ExchangeCheck = 'Passed'
            }
            catch {
                $values.ExchangeCheck = 'Failed'
                $problems.Add("Exchange: $($_.Exception.Message)")
            }
            finally {
                Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
            }
        }

        if ($problems.Count -gt 0) {
            $values.Status = 'Failed'
            $values.Detail = $problems -join ' || '
        }
        $row = ConvertTo-ResultRow -Customer $customer -Values $values
        $results.Add($row)
        $row
    }

    if ($OutputPath -and $results.Count -gt 0) {
        $results | Export-Csv -LiteralPath $OutputPath -NoTypeInformation -Encoding utf8
        Write-Verbose "Saved $($results.Count) rows to $OutputPath"
    }
}
