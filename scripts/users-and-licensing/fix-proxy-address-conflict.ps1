#Requires -Version 7.4
#Requires -Modules MspGdap

<#
.SYNOPSIS
    Finds which recipient holds an email address that Exchange Online says is "already being used by the proxy address of another mailbox", and moves it safely.

.DESCRIPTION
    By default the script only reports. In each customer tenant it opens an
    Exchange Online session through GDAP and lists every recipient whose email
    addresses include -EmailAddress, showing whether it is that recipient's
    primary address.

    To free the address, name the mailbox that should give it up with -RemoveFrom
    and add -Apply. In this order, and only for the parts you ask for:
    1. -NewUserPrincipalName changes that mailbox user's user principal name
       through Microsoft Graph (cloud-only users).
    2. If the address is the mailbox's primary address, -NewPrimarySmtpAddress
       makes another address primary first. Without it the script stops, because
       a primary address cannot simply be removed.
    3. The conflicting address is removed with
       Set-Mailbox -EmailAddresses @{Remove='smtp:...'}, which leaves every other
       address in place. The original article used
       Set-Mailbox -EmailAddresses SMTP:..., which replaces the whole list and
       drops every other alias, including the .onmicrosoft.com address.
    4. -DefaultDomain makes a verified domain the tenant's default domain through
       Microsoft Graph, so new mailboxes stop getting addresses on the wrong domain.

    Each change goes through ShouldProcess and supports -WhatIf, and the mailbox
    is read back afterwards to confirm the address has gone. The Exchange Online
    session is closed after each customer.

.PARAMETER TenantId
    One or more customer tenant IDs (GUID) or verified domains, such as
    contoso.onmicrosoft.com. Accepts pipeline input. Normally one tenant.

.PARAMETER AllCustomers
    Searches every customer with an active GDAP relationship. Report only.

.PARAMETER EmailAddress
    The address that is reported as already in use, such as accounts@contoso.com.

.PARAMETER RemoveFrom
    The mailbox (user principal name, primary address or GUID) that should give up
    the address.

.PARAMETER NewPrimarySmtpAddress
    A new primary address for the -RemoveFrom mailbox. Needed only when the
    conflicting address is its primary address.

.PARAMETER NewUserPrincipalName
    A new user principal name for the -RemoveFrom mailbox's user account.

.PARAMETER DefaultDomain
    A verified domain to make the tenant's default domain.

.PARAMETER Apply
    Makes the requested changes. Without it the script only reports.

.PARAMETER OutputPath
    Optional path of a CSV file for the results. The file is overwritten.

.EXAMPLE
    ./fix-proxy-address-conflict.ps1 -TenantId 'contoso.onmicrosoft.com' -EmailAddress 'accounts@contoso.com'

    Shows which recipients hold accounts@contoso.com.

.EXAMPLE
    ./fix-proxy-address-conflict.ps1 -TenantId 'contoso.onmicrosoft.com' -EmailAddress 'accounts@contoso.com' -RemoveFrom 'accounts@fabrikam.com' -NewPrimarySmtpAddress 'accounts@fabrikam.com' -NewUserPrincipalName 'accounts@fabrikam.com' -Apply -WhatIf

    Previews moving the address off a shared mailbox. Remove -WhatIf to make the changes.

.NOTES
    Replaces the original 2017 method: Connect-MsolService, Set-MsolUserPrincipalName and Set-MsolDomain -IsDefault (MSOnline module, retired 30 May 2025), Exchange Online remote PowerShell with basic authentication, and Set-Mailbox -EmailAddresses SMTP:... which overwrote every alias.
    Required GDAP roles: Exchange Recipient Administrator (or Global Reader for the report), User Administrator for -NewUserPrincipalName, Domain Name Administrator for -DefaultDomain.
    Required partner app permissions: Office 365 Exchange Online delegated Exchange.Manage, Microsoft Graph delegated User.ReadWrite.All, and Microsoft Graph delegated Domain.ReadWrite.All for -DefaultDomain (all in manifests/partner-app.full.json).

.LINK
    https://gcit.com.au/knowledge-base/the-proxy-address-is-already-being-used-by-the-proxy-address-of-another-mailbox/

.LINK
    docs/05-exchange-access.md

.LINK
    docs/07-migrating-from-dap-msonline.md
#>
[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High', DefaultParameterSetName = 'Tenant')]
[OutputType([pscustomobject])]
param(
    [Parameter(Mandatory, ParameterSetName = 'Tenant', ValueFromPipeline, ValueFromPipelineByPropertyName)]
    [ValidateNotNullOrEmpty()]
    [string[]]$TenantId,

    [Parameter(Mandatory, ParameterSetName = 'AllCustomers')]
    [switch]$AllCustomers,

    [Parameter(Mandatory)]
    [ValidatePattern('^[^\s@'']+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}$')]
    [string]$EmailAddress,

    [ValidateNotNullOrEmpty()]
    [string]$RemoveFrom,

    [ValidatePattern('^[^\s@'']+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}$')]
    [string]$NewPrimarySmtpAddress,

    [ValidatePattern('^[^\s@'']+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}$')]
    [string]$NewUserPrincipalName,

    [ValidatePattern('^[A-Za-z0-9.-]+\.[A-Za-z]{2,}$')]
    [string]$DefaultDomain,

    [switch]$Apply,

    [ValidateNotNullOrEmpty()]
    [string]$OutputPath
)

begin {
    if ($AllCustomers -and ($RemoveFrom -or $DefaultDomain -or $Apply)) {
        throw '-AllCustomers is for reporting only. Name the customer with -TenantId to make changes.'
    }
    if (($NewPrimarySmtpAddress -or $NewUserPrincipalName) -and -not $RemoveFrom) {
        throw '-NewPrimarySmtpAddress and -NewUserPrincipalName apply to the -RemoveFrom mailbox. Add -RemoveFrom.'
    }

    $columns = @('CustomerTenantId', 'CustomerName', 'EmailAddress', 'HolderCount', 'Holders', 'RemoveFrom', 'UpnAction', 'PrimaryAction', 'RemoveAction', 'DefaultDomainAction', 'Status', 'Error')
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

    function Test-HasAddress {
        param($Addresses, [string]$Address)
        foreach ($entry in @($Addresses)) {
            if ([string]$entry -match '^smtp:(.+)$' -and $Matches[1] -eq $Address) { return $true }
        }
        return $false
    }

    function Test-IsPrimary {
        param($Addresses, [string]$Address)
        foreach ($entry in @($Addresses)) {
            if ([string]$entry -cmatch '^SMTP:(.+)$' -and $Matches[1] -eq $Address) { return $true }
        }
        return $false
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
        $values = [ordered]@{
            CustomerTenantId    = $target
            EmailAddress        = $EmailAddress
            RemoveFrom          = $RemoveFrom
            UpnAction           = 'None'
            PrimaryAction       = 'None'
            RemoveAction        = 'None'
            DefaultDomainAction = 'None'
            Status              = 'OK'
        }
        try {
            $tenant = Resolve-MspTenantId -Tenant $target
            $values.CustomerTenantId = $tenant
            $values.CustomerName = $knownNames[$tenant]
            if (-not $values.CustomerName) {
                $values.CustomerName = @(Invoke-MspGraphRequest -TenantId $tenant -Uri 'v1.0/organization?$select=displayName')[0].displayName
            }
            $label = "$($values.CustomerName) ($tenant)"

            Connect-MspExchangeOnline -TenantId $tenant | Out-Null
            try {
                $holders = @(Get-EXORecipient -Filter "EmailAddresses -eq 'smtp:$EmailAddress'" -Properties EmailAddresses, ExternalDirectoryObjectId -ResultSize Unlimited -ErrorAction Stop)
                $values.HolderCount = $holders.Count
                $values.Holders = ($holders | ForEach-Object {
                        $kind = if (Test-IsPrimary -Addresses $_.EmailAddresses -Address $EmailAddress) { 'primary' } else { 'alias' }
                        '{0} [{1}, {2}]' -f $_.PrimarySmtpAddress, $_.RecipientTypeDetails, $kind
                    }) -join ' | '

                if ($RemoveFrom) {
                    $mailbox = Get-EXOMailbox -Identity $RemoveFrom -Properties EmailAddresses, ExternalDirectoryObjectId -ErrorAction Stop
                    $mailboxId = [string]$mailbox.ExternalDirectoryObjectId
                    $identity = if ($mailbox.ExchangeGuid) { [string]$mailbox.ExchangeGuid } else { [string]$mailbox.UserPrincipalName }
                    $holdsAddress = Test-HasAddress -Addresses $mailbox.EmailAddresses -Address $EmailAddress
                    $isPrimary = Test-IsPrimary -Addresses $mailbox.EmailAddresses -Address $EmailAddress

                    if ($NewUserPrincipalName) {
                        if ($mailbox.UserPrincipalName -eq $NewUserPrincipalName) {
                            $values.UpnAction = 'AlreadySet'
                        }
                        elseif (-not $Apply) {
                            $values.UpnAction = 'WouldChange'
                        }
                        elseif ($PSCmdlet.ShouldProcess("$($mailbox.UserPrincipalName) in $label", "Change user principal name to $NewUserPrincipalName")) {
                            Invoke-MspGraphRequest -TenantId $tenant -Method PATCH -Uri "v1.0/users/$mailboxId" -Body @{ userPrincipalName = $NewUserPrincipalName } -Confirm:$false | Out-Null
                            $values.UpnAction = 'Changed'
                        }
                        else {
                            $values.UpnAction = 'WhatIf'
                        }
                    }

                    if (-not $holdsAddress) {
                        $values.RemoveAction = 'NotPresent'
                    }
                    elseif ($isPrimary -and -not $NewPrimarySmtpAddress) {
                        $values.PrimaryAction = 'NeedsNewPrimarySmtpAddress'
                        $values.RemoveAction = 'Blocked'
                        $values.Status = 'Warning'
                    }
                    else {
                        $primaryReady = -not $isPrimary
                        if ($isPrimary) {
                            if (-not $Apply) {
                                $values.PrimaryAction = 'WouldChange'
                            }
                            elseif ($PSCmdlet.ShouldProcess("$($mailbox.PrimarySmtpAddress) in $label", "Make $NewPrimarySmtpAddress the primary address")) {
                                Set-Mailbox -Identity $identity -EmailAddresses @{ Add = "SMTP:$NewPrimarySmtpAddress" } -ErrorAction Stop
                                $values.PrimaryAction = 'Changed'
                                $primaryReady = $true
                            }
                            else {
                                $values.PrimaryAction = 'WhatIf'
                            }
                        }

                        if (-not $Apply) {
                            $values.RemoveAction = 'WouldRemove'
                        }
                        elseif (-not $primaryReady) {
                            $values.RemoveAction = 'SkippedPrimaryUnchanged'
                        }
                        elseif ($PSCmdlet.ShouldProcess("$($mailbox.PrimarySmtpAddress) in $label", "Remove address $EmailAddress (other addresses stay)")) {
                            Set-Mailbox -Identity $identity -EmailAddresses @{ Remove = "smtp:$EmailAddress" } -ErrorAction Stop
                            $check = Get-EXOMailbox -Identity $identity -Properties EmailAddresses -ErrorAction Stop
                            if (Test-HasAddress -Addresses $check.EmailAddresses -Address $EmailAddress) {
                                $values.RemoveAction = 'Unconfirmed'
                                $values.Status = 'Warning'
                            }
                            else {
                                $values.RemoveAction = 'Removed'
                            }
                        }
                        else {
                            $values.RemoveAction = 'WhatIf'
                        }
                    }
                }
            }
            finally {
                Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
            }

            if ($DefaultDomain) {
                $domain = Invoke-MspGraphRequest -TenantId $tenant -Uri ('v1.0/domains/{0}?$select=id,isDefault,isVerified' -f $DefaultDomain)
                if ($domain.isDefault) {
                    $values.DefaultDomainAction = 'AlreadyDefault'
                }
                elseif (-not $domain.isVerified) {
                    $values.DefaultDomainAction = 'NotVerified'
                    $values.Status = 'Warning'
                }
                elseif (-not $Apply) {
                    $values.DefaultDomainAction = 'WouldChange'
                }
                elseif ($PSCmdlet.ShouldProcess($label, "Make $DefaultDomain the default domain")) {
                    Invoke-MspGraphRequest -TenantId $tenant -Method PATCH -Uri "v1.0/domains/$DefaultDomain" -Body @{ isDefault = $true } -Confirm:$false | Out-Null
                    $values.DefaultDomainAction = 'Changed'
                }
                else {
                    $values.DefaultDomainAction = 'WhatIf'
                }
            }
        }
        catch {
            Write-Warning -Message "Customer $target failed: $($_.Exception.Message)"
            $values.Status = 'Failed'
            $values.Error = $_.Exception.Message
        }

        $row = Get-ResultRow -Column $columns -Value $values
        $results.Add($row)
        $row
    }

    if ($OutputPath -and $results.Count -gt 0) {
        $results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding utf8 -WhatIf:$false -Confirm:$false
    }
}
