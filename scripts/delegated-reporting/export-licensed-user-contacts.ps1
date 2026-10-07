#Requires -Version 7.4
#Requires -Modules MspGdap

<#
.SYNOPSIS
    Exports your customers' licensed users as a contact list for a double opt-in newsletter invitation, with licence-based segments.

.DESCRIPTION
    The original article read every licensed user from every customer and
    subscribed them straight to a Mailchimp audience, with the Mailchimp API key
    typed into the script. Subscribing people to a marketing list without their
    own consent is likely to breach the Spam Act 2003 in Australia (and similar
    laws elsewhere), whatever your customer contract says. This script does not
    subscribe anyone.

    Instead it exports a contact list you can import into your email platform as
    pending (unconfirmed) contacts, so each person receives a confirmation email
    and is only added once they opt in (double opt-in). For each customer it
    returns one row per enabled, licensed member user with an email address:
    customer name, display name, given name and surname (taken from the display
    name when blank), email, licence names (SKU part numbers)
    and a Segment column built from those licences, plus ConsentStatus set to
    NotRequested.

    Before you use the list, check that your customer agreements and privacy
    policy allow you to contact your customers' staff, and keep the list only as
    long as you need it. Store any API key for your email platform in a
    SecretManagement vault, never in a script.

    The original also unsubscribed people who were no longer licensed. Do the same
    by comparing each new export with the audience in your email platform and
    archiving the contacts that are no longer listed.

    The script only reads. It never changes a tenant.

.PARAMETER TenantId
    One or more customer tenant IDs (GUID) or verified domains, such as
    contoso.onmicrosoft.com. Accepts pipeline input, including objects from
    Get-MspCustomer.

.PARAMETER AllCustomers
    Runs against every customer with an active GDAP relationship, as returned by
    Get-MspCustomer -IncludeGdapStatus.

.PARAMETER OutputPath
    Optional path of a CSV file for the results. The file is overwritten.

.EXAMPLE
    ./export-licensed-user-contacts.ps1 -AllCustomers -OutputPath ./newsletter-invitations.csv

    Exports licensed users from every active GDAP customer for a double opt-in invitation.

.EXAMPLE
    ./export-licensed-user-contacts.ps1 -TenantId 'contoso.onmicrosoft.com' | Group-Object Segment | Select-Object Name, Count

    Counts one customer's users per licence segment.

.NOTES
    Replaces the original 2017 method: Connect-MsolService, Get-MsolPartnerContract (DAP), Get-MsolUser -TenantId and Get-MsolCompanyInformation (MSOnline module, retired 30 May 2025), with users subscribed directly to a Mailchimp audience through API 3.0 and a plain-text API key (removed as unsafe).
    Required GDAP roles: Directory Readers or Global Reader.
    Required partner app permissions: Microsoft Graph delegated User.Read.All and LicenseAssignment.Read.All (covered in manifests/partner-app.full.json by User.ReadWrite.All and Directory.ReadWrite.All).

.LINK
    https://gcit.com.au/knowledge-base/sync-office-365-users-mailchimp-list/

.LINK
    https://www.acma.gov.au/avoid-sending-spam

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

    [ValidateNotNullOrEmpty()]
    [string]$OutputPath
)

begin {
    $columns = @('CustomerTenantId', 'CustomerName', 'DisplayName', 'GivenName', 'Surname', 'Email', 'Licences', 'Segment', 'ConsentStatus', 'Status', 'Error')
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

    function Get-LicenceSegment {
        param([string[]]$SkuPartNumber)
        $segments = foreach ($sku in $SkuPartNumber) {
            switch -Regex ($sku) {
                '^(SPE_E5|ENTERPRISEPREMIUM|SPE_E3|ENTERPRISEPACK|STANDARDPACK|SPE_F1|DESKLESSPACK)$' { 'Enterprise'; break }
                '^(SPB|O365_BUSINESS_PREMIUM|SMB_BUSINESS_PREMIUM|O365_BUSINESS_ESSENTIALS|SMB_BUSINESS_ESSENTIALS|O365_BUSINESS|SMB_BUSINESS)$' { 'Business'; break }
                '^(EXCHANGESTANDARD|EXCHANGEENTERPRISE|EXCHANGEDESKLESS)$' { 'Exchange Online only'; break }
                default { $null }
            }
        }
        $list = @($segments | Where-Object { $_ } | Sort-Object -Unique)
        if ($list.Count -eq 0) { 'Other' } else { $list -join ', ' }
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
        $tenant = $target
        $customerName = $null
        try {
            $tenant = Resolve-MspTenantId -Tenant $target
            $customerName = $knownNames[$tenant]
            if (-not $customerName) {
                $customerName = @(Invoke-MspGraphRequest -TenantId $tenant -Uri 'v1.0/organization?$select=displayName')[0].displayName
            }

            $skuNames = @{}
            foreach ($sku in @(Invoke-MspGraphRequest -TenantId $tenant -Uri 'v1.0/subscribedSkus?$select=skuId,skuPartNumber')) {
                $skuNames[[string]$sku.skuId] = $sku.skuPartNumber
            }

            $uri = 'v1.0/users?$select=id,displayName,givenName,surname,mail,userType,accountEnabled,assignedLicenses&$top=999'
            $users = @(Invoke-MspGraphRequest -TenantId $tenant -Uri $uri | Where-Object { $_.userType -eq 'Member' -and $_.accountEnabled -and $_.mail })
            foreach ($user in $users) {
                $licences = @($user.assignedLicenses | Where-Object { $_ } | ForEach-Object {
                        $id = [string]$_.skuId
                        if ($skuNames.ContainsKey($id)) { $skuNames[$id] } else { $id }
                    })
                if ($licences.Count -eq 0) { continue }

                # As in the original, fall back to the display name when the first or last name is blank.
                $givenName = $user.givenName
                $surname = $user.surname
                if ((-not $givenName -or -not $surname) -and $user.displayName) {
                    $parts = @(([string]$user.displayName).Trim() -split '\s+')
                    if (-not $givenName) { $givenName = $parts[0] }
                    if (-not $surname -and $parts.Count -gt 1) { $surname = ($parts[1..($parts.Count - 1)]) -join ' ' }
                }

                $row = Get-ResultRow -Column $columns -Value @{
                    CustomerTenantId = $tenant
                    CustomerName     = $customerName
                    DisplayName      = $user.displayName
                    GivenName        = $givenName
                    Surname          = $surname
                    Email            = $user.mail
                    Licences         = ($licences | Sort-Object) -join ', '
                    Segment          = Get-LicenceSegment -SkuPartNumber $licences
                    ConsentStatus    = 'NotRequested'
                    Status           = 'OK'
                }
                $results.Add($row)
                $row
            }
        }
        catch {
            Write-Warning -Message "Customer $target failed: $($_.Exception.Message)"
            $row = Get-ResultRow -Column $columns -Value @{
                CustomerTenantId = $tenant
                CustomerName     = $customerName
                Status           = 'Failed'
                Error            = $_.Exception.Message
            }
            $results.Add($row)
            $row
        }
    }

    if ($OutputPath -and $results.Count -gt 0) {
        $results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding utf8 -WhatIf:$false -Confirm:$false
    }
}
