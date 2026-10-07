#Requires -Version 7.4
#Requires -Modules MspGdap

<#
.SYNOPSIS
    Checks that you can reach Microsoft Graph (and optionally Exchange Online) in each customer tenant through GDAP.

.DESCRIPTION
    This is the modern replacement for "install the MSOnline module and run
    Connect-MsolService". There is nothing to install per customer and no
    partner-wide session to open. Install and configure MspGdap once
    (see the README and docs/03), then run this script to confirm that each
    customer is ready for the other scripts in this folder.

    For each customer it checks that:
    - the tenant ID or domain resolves,
    - the organisation can be read through Microsoft Graph (proves the partner
      app is consented and a token can be issued in that tenant),
    - users can be read (proves a GDAP role that can read the directory),
    - with -IncludeExchange, an Exchange Online session opens and a mailbox can
      be read, then the session is closed again.

    Every check is recorded separately, so you can see which step failed. The
    script only reads. It never changes a tenant.

    Prerequisites, once per workstation (PowerShell 7.4 or later):
        Install-PSResource -Name Microsoft.PowerShell.SecretManagement, Microsoft.PowerShell.SecretStore
        Install-PSResource -Name ExchangeOnlineManagement   (only for -IncludeExchange)
        Import-Module MspGdap
        Set-MspConfiguration and Register-MspPartnerToken (docs/03)

    For interactive work in a single tenant you can also use the Microsoft Graph
    PowerShell SDK (Install-PSResource -Name Microsoft.Graph) together with
    Connect-MspGraph -TenantId.

.PARAMETER TenantId
    One or more customer tenant IDs (GUID) or verified domains, such as
    contoso.onmicrosoft.com. Accepts pipeline input, including objects from
    Get-MspCustomer.

.PARAMETER AllCustomers
    Runs against every customer with an active GDAP relationship, as returned by
    Get-MspCustomer -IncludeGdapStatus.

.PARAMETER IncludeExchange
    Also opens an Exchange Online session with Connect-MspExchangeOnline, reads one
    mailbox with Get-EXOMailbox and disconnects.

.PARAMETER OutputPath
    Optional path of a CSV file for the results. The file is overwritten.

.EXAMPLE
    ./test-customer-graph-access.ps1 -TenantId 'contoso.onmicrosoft.com' -IncludeExchange

    Checks Graph and Exchange Online access to one customer.

.EXAMPLE
    ./test-customer-graph-access.ps1 -AllCustomers -OutputPath ./gdap-readiness.csv

    Checks every active GDAP customer and saves the results to CSV.

.NOTES
    Replaces the original 2017 method: Install-Module MSOnline and Connect-MsolService (MSOnline module, retired 30 May 2025, and mislabelled as the Azure Active Directory PowerShell module).
    Required GDAP roles: Directory Readers or Global Reader, plus Global Reader or Exchange Recipient Administrator for -IncludeExchange.
    Required partner app permissions: Microsoft Graph delegated User.Read and User.Read.All (covered in manifests/partner-app.full.json by User.Read and User.ReadWrite.All), and Office 365 Exchange Online delegated Exchange.Manage for -IncludeExchange.

.LINK
    https://gcit.com.au/knowledge-base/install-azure-active-directory-powershell-module/

.LINK
    docs/03-register-technician-token.md

.LINK
    docs/07-migrating-from-dap-msonline.md
#>
[CmdletBinding(DefaultParameterSetName = 'Tenant')]
[OutputType([pscustomobject])]
param(
    [Parameter(Mandatory, ParameterSetName = 'Tenant', ValueFromPipeline, ValueFromPipelineByPropertyName)]
    [ValidateNotNullOrEmpty()]
    [string[]]$TenantId,

    [Parameter(Mandatory, ParameterSetName = 'AllCustomers')]
    [switch]$AllCustomers,

    [switch]$IncludeExchange,

    [ValidateNotNullOrEmpty()]
    [string]$OutputPath
)

begin {
    $columns = @('CustomerTenantId', 'CustomerName', 'InitialDomain', 'DefaultDomain', 'GraphOrganisationRead', 'GraphUserRead', 'ExchangeConnect', 'Status', 'Error')
    $targets = [System.Collections.Generic.List[string]]::new()
    $knownNames = @{}
    $results = [System.Collections.Generic.List[object]]::new()

    function Get-ResultRow {
        param(
            [Parameter(Mandatory)][string[]]$Column,
            [Parameter(Mandatory)][System.Collections.IDictionary]$Value
        )
        $row = [ordered]@{}
        foreach ($name in $Column) {
            $row[$name] = if ($Value.Contains($name)) { $Value[$name] } else { $null }
        }
        [pscustomobject]$row
    }
}

process {
    if ($PSCmdlet.ParameterSetName -eq 'Tenant') {
        foreach ($item in $TenantId) { $targets.Add($item) }
    }
}

end {
    if ($AllCustomers) {
        foreach ($customer in @(Get-MspCustomer -IncludeGdapStatus | Where-Object { $_.GdapStatus -eq 'active' })) {
            $targets.Add($customer.TenantId)
            $knownNames[$customer.TenantId] = $customer.DisplayName
        }
    }

    foreach ($target in $targets) {
        $values = [ordered]@{
            CustomerTenantId      = $target
            CustomerName          = $null
            GraphOrganisationRead = 'NotRun'
            GraphUserRead         = 'NotRun'
            ExchangeConnect       = if ($IncludeExchange) { 'NotRun' } else { 'Skipped' }
        }
        $problems = [System.Collections.Generic.List[string]]::new()

        try {
            $tenant = Resolve-MspTenantId -Tenant $target
            $values.CustomerTenantId = $tenant
            $values.CustomerName = $knownNames[$tenant]

            try {
                $organisation = @(Invoke-MspGraphRequest -TenantId $tenant -Uri 'v1.0/organization?$select=id,displayName,verifiedDomains')[0]
                if (-not $values.CustomerName) { $values.CustomerName = $organisation.displayName }
                $values.InitialDomain = (@($organisation.verifiedDomains) | Where-Object { $_.isInitial } | Select-Object -First 1).name
                $values.DefaultDomain = (@($organisation.verifiedDomains) | Where-Object { $_.isDefault } | Select-Object -First 1).name
                $values.GraphOrganisationRead = 'Passed'
            }
            catch {
                $values.GraphOrganisationRead = 'Failed'
                $problems.Add("Organisation read: $($_.Exception.Message)")
            }

            try {
                $null = Invoke-MspGraphRequest -TenantId $tenant -Uri 'v1.0/users?$select=id&$top=1' -NoPaging
                $values.GraphUserRead = 'Passed'
            }
            catch {
                $values.GraphUserRead = 'Failed'
                $problems.Add("User read: $($_.Exception.Message)")
            }

            if ($IncludeExchange) {
                try {
                    Connect-MspExchangeOnline -TenantId $tenant | Out-Null
                    $null = Get-EXOMailbox -ResultSize 1 -ErrorAction Stop
                    $values.ExchangeConnect = 'Passed'
                }
                catch {
                    $values.ExchangeConnect = 'Failed'
                    $problems.Add("Exchange Online: $($_.Exception.Message)")
                }
                finally {
                    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
                }
            }
        }
        catch {
            $problems.Add($_.Exception.Message)
        }

        $values.Status = if ($problems.Count -eq 0) { 'OK' } else { 'Failed' }
        $values.Error = if ($problems.Count -gt 0) { $problems -join ' | ' } else { $null }
        if ($problems.Count -gt 0) {
            Write-Warning -Message "Customer $target failed: $($values.Error)"
        }

        $row = Get-ResultRow -Column $columns -Value $values
        $results.Add($row)
        $row
    }

    if ($OutputPath -and $results.Count -gt 0) {
        $results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding utf8 -WhatIf:$false -Confirm:$false
    }
}
