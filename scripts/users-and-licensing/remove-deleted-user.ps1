#Requires -Version 7.4
#Requires -Modules MspGdap

<#
.SYNOPSIS
    Lists soft-deleted users in a customer tenant with the days left until Microsoft purges them, to help clear the "Work or school account or personal account" prompt.

.DESCRIPTION
    When the same email address exists as a Microsoft Entra work account and as a
    personal Microsoft account, sign-in pages ask which one to use. Once you have
    removed the work account that should not exist (for example one created by a
    self-service sign-up), the deleted user stays in the tenant's deleted users
    list for 30 days and the prompt can stay until it is gone.

    The script lists the soft-deleted users in each tenant with the date they were
    deleted, the date Microsoft Entra starts to purge them automatically (30 days
    after deletion) and the whole days left until then. -UserPrincipalName narrows
    the list to the accounts you name, and each name that is not found gets its own
    NotFound row. Each row has a Guidance column, and the script writes the same
    guidance once to the information stream.

    The script changes nothing. It does not permanently delete users. In most cases
    waiting for the automatic purge is enough. When an account really has to go
    sooner, permanently delete that one account by hand in the Microsoft Entra
    admin center: Entra ID > Users > Deleted users, select the user, then Delete
    permanently (at least User Administrator). The Microsoft 365 admin center lists
    the same accounts under Users > Deleted users, for restoring.

    The toolkit deliberately does not hold a standing permission to permanently
    delete users (Microsoft Graph User.DeleteRestore.All). A permanent deletion
    cannot be undone by anyone, including Microsoft support, it is rarely needed
    because Microsoft Entra purges deleted users after 30 days anyway, and least
    privilege means a delegated permission that every technician token carries
    into every consented customer should not include an irreversible action that
    is needed a few times a year. The original article purged every deleted user
    in the tenant without any filter, which this script does not reproduce.

    Taking over an unmanaged (self-service) tenant is a separate step in the
    Microsoft Entra admin center (admin takeover). GDAP does not reach an
    unmanaged tenant, so this script applies once the domain and its users are in
    a managed customer tenant you have GDAP access to.

.PARAMETER TenantId
    One or more customer tenant IDs (GUID) or verified domains, such as
    contoso.onmicrosoft.com. Accepts pipeline input, including objects from
    Get-MspCustomer.

.PARAMETER AllCustomers
    Runs against every customer with an active GDAP relationship, as returned by
    Get-MspCustomer -IncludeGdapStatus.

.PARAMETER UserPrincipalName
    One or more original user principal names, email addresses or deleted object
    IDs to list. Deleted users get their object ID (without hyphens) added to the
    front of their user principal name, and names are matched exactly against
    that form, so a partial name never selects another account.

.PARAMETER OutputPath
    Optional path of a CSV file for the results. The file is overwritten.

.EXAMPLE
    ./remove-deleted-user.ps1 -TenantId 'contoso.onmicrosoft.com'

    Lists the soft-deleted users in one customer tenant with their purge dates.

.EXAMPLE
    ./remove-deleted-user.ps1 -TenantId 'contoso.onmicrosoft.com' -UserPrincipalName 'adele.vance@contoso.com'

    Shows whether that account is still in the deleted users list and how many
    days are left before Microsoft Entra purges it.

.NOTES
    Replaces the original 2016 method: Connect-MsolService and Get-MsolUser -ReturnDeletedUsers | Remove-MsolUser -RemoveFromRecycleBin -Force (MSOnline module, retired 30 May 2025, and an unfiltered purge of every deleted user). The permanent delete is now a deliberate manual step in the Microsoft Entra admin center.
    Required GDAP roles: Directory Readers or Global Reader to list. A manual permanent delete in the Microsoft Entra admin center needs at least User Administrator, and Microsoft documents that deleted users who held privileged admin roles can need a higher role (Privileged Authentication Administrator).
    Required partner app permissions: Microsoft Graph delegated User.Read.All is the least-privileged permission for directory/deletedItems/microsoft.graph.user. Directory.Read.All covers it and is in both manifests/partner-app.minimal.json and manifests/partner-app.full.json. User.DeleteRestore.All is not needed and is deliberately not in either manifest.

.LINK
    https://gcit.com.au/knowledge-base/remove-work-or-school-account-option/

.LINK
    docs/07-migrating-from-dap-msonline.md

.LINK
    https://learn.microsoft.com/entra/fundamentals/users-restore
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
    [string[]]$UserPrincipalName,

    [ValidateNotNullOrEmpty()]
    [string]$OutputPath
)

begin {
    $retentionDays = 30
    $columns = @('CustomerTenantId', 'CustomerName', 'DeletedObjectId', 'DisplayName', 'UserPrincipalName', 'OriginalUserPrincipalName', 'Mail', 'DeletedDateTime', 'PurgeDateTime', 'DaysUntilPurge', 'Status', 'Guidance', 'Error')
    $targets = [System.Collections.Generic.List[string]]::new()
    $knownNames = @{}
    $results = [System.Collections.Generic.List[object]]::new()

    $manualGuidance = 'Nothing to do in most cases: Microsoft Entra purges deleted users automatically {0} days after deletion. If the account must go sooner, permanently delete it in the Microsoft Entra admin center (Entra ID > Users > Deleted users > Delete permanently, at least User Administrator). This cannot be undone. The toolkit has no permanent-delete permission on purpose (least privilege for an irreversible, rarely needed action).' -f $retentionDays

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

    function Test-NameMatch {
        param($DeletedUser, [string]$Name)
        # A deleted user's userPrincipalName is its object ID without hyphens followed by the
        # original name, so match that exact form. A plain EndsWith would also match
        # jim.smith@contoso.com when you asked for m.smith@contoso.com.
        $deletedUpn = [string]$DeletedUser.userPrincipalName
        $prefix = ([string]$DeletedUser.id) -replace '-', ''
        if ([string]$DeletedUser.id -eq $Name) { return $true }
        if ($deletedUpn -and ($deletedUpn -eq $Name -or $deletedUpn -eq "$prefix$Name")) { return $true }
        if ([string]$DeletedUser.mail -and [string]$DeletedUser.mail -eq $Name) { return $true }
        return $false
    }

    function Get-OriginalName {
        param($DeletedUser)
        $deletedUpn = [string]$DeletedUser.userPrincipalName
        $prefix = ([string]$DeletedUser.id) -replace '-', ''
        if ($prefix -and $deletedUpn.StartsWith($prefix, [System.StringComparison]::OrdinalIgnoreCase)) {
            return $deletedUpn.Substring($prefix.Length)
        }
        return $deletedUpn
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

    $now = [datetimeoffset]::UtcNow

    foreach ($target in $targets) {
        $tenant = $target
        $customerName = $null
        try {
            $tenant = Resolve-MspTenantId -Tenant $target
            $customerName = $knownNames[$tenant]
            if (-not $customerName) {
                $customerName = @(Invoke-MspGraphRequest -TenantId $tenant -Uri 'v1.0/organization?$select=displayName')[0].displayName
            }

            $deleted = @(Invoke-MspGraphRequest -TenantId $tenant -Uri 'v1.0/directory/deletedItems/microsoft.graph.user?$select=id,displayName,userPrincipalName,mail,deletedDateTime')
            if ($UserPrincipalName) {
                $selected = [System.Collections.Generic.List[object]]::new()
                foreach ($name in $UserPrincipalName) {
                    $hits = @($deleted | Where-Object { Test-NameMatch -DeletedUser $_ -Name $name })
                    foreach ($hit in $hits) {
                        if (-not ($selected | Where-Object { $_.id -eq $hit.id })) { $selected.Add($hit) }
                    }
                    if ($hits.Count -eq 0) {
                        $row = Get-ResultRow -Column $columns -Value @{
                            CustomerTenantId  = $tenant
                            CustomerName      = $customerName
                            UserPrincipalName = $name
                            Status            = 'NotFound'
                            Guidance          = 'Not in the deleted users list. It may already be purged, restored or never deleted.'
                        }
                        $results.Add($row)
                        $row
                    }
                }
                $deleted = @($selected)
            }

            foreach ($user in $deleted) {
                $values = [ordered]@{
                    CustomerTenantId          = $tenant
                    CustomerName              = $customerName
                    DeletedObjectId           = $user.id
                    DisplayName               = $user.displayName
                    UserPrincipalName         = $user.userPrincipalName
                    OriginalUserPrincipalName = Get-OriginalName -DeletedUser $user
                    Mail                      = $user.mail
                    DeletedDateTime           = $user.deletedDateTime
                    Status                    = 'OK'
                    Guidance                  = $manualGuidance
                }

                if ($user.deletedDateTime) {
                    $deletedAt = if ($user.deletedDateTime -is [datetime]) {
                        [datetimeoffset]::new(([datetime]$user.deletedDateTime).ToUniversalTime())
                    }
                    else {
                        [datetimeoffset]::Parse([string]$user.deletedDateTime, [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::AssumeUniversal)
                    }
                    $purgeAt = $deletedAt.AddDays($retentionDays)
                    $values.PurgeDateTime = $purgeAt.UtcDateTime.ToString('yyyy-MM-ddTHH:mm:ssZ', [System.Globalization.CultureInfo]::InvariantCulture)
                    $values.DaysUntilPurge = [int][math]::Max(0, [math]::Ceiling(($purgeAt - $now).TotalDays))
                }

                $row = Get-ResultRow -Column $columns -Value $values
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

    if (@($results | Where-Object { $_.Status -eq 'OK' }).Count -gt 0) {
        Write-Information -MessageData $manualGuidance -InformationAction Continue
    }

    if ($OutputPath -and $results.Count -gt 0) {
        $results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding utf8
    }
}
