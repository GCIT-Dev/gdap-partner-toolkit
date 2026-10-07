#Requires -Version 7.4
#Requires -Modules MspGdap

<#
.SYNOPSIS
    Reports user details, licences and optionally MFA registration and last sign-in for customer tenants.

.DESCRIPTION
    Reads every user in each customer through Microsoft Graph (GET /users) and returns one flat row per
    user with account state, licences (by SKU part number), last password change, sync state and contact
    details.

    -IncludeAuthMethods adds MFA registration details from GET /reports/authenticationMethods/
    userRegistrationDetails. -IncludeSignInActivity adds the last interactive and non-interactive sign-in
    times. Both need Microsoft Entra ID P1 or P2 in the customer tenant. -IncludeMailboxType connects to
    Exchange Online through MspGdap and adds each user's mailbox type (UserMailbox, SharedMailbox and so
    on), as the original did.

    The rows are flat, so they can go straight to CSV, to a SharePoint list through Power Automate, or to
    a PSA. The UserDetailsReport Azure Function in the functions folder writes them to blob storage on a
    timer.

.PARAMETER TenantId
    One or more customer tenant IDs (GUID) or verified domains. Accepts pipeline input, including the
    output of Get-MspCustomer.

.PARAMETER AllCustomers
    Process every customer returned by Get-MspCustomer.

.PARAMETER LicensedOnly
    Only return users with at least one licence.

.PARAMETER IncludeAuthMethods
    Add MFA registration details (needs Microsoft Entra ID P1 or P2).

.PARAMETER IncludeSignInActivity
    Add last sign-in times (needs Microsoft Entra ID P1 or P2).

.PARAMETER IncludeMailboxType
    Add each user's mailbox type from Exchange Online (Get-EXOMailbox RecipientTypeDetails).

.PARAMETER OutputPath
    Optional path for a CSV export of the results.

.EXAMPLE
    ./get-user-details.ps1 -AllCustomers -LicensedOnly -OutputPath ./users.csv

    Exports every licensed user across all customers.

.EXAMPLE
    ./get-user-details.ps1 -TenantId 'contoso.onmicrosoft.com' -IncludeAuthMethods -IncludeSignInActivity

    Lists users in one customer with MFA registration and last sign-in times.

.NOTES
    Replaces the original 2017 method: an Azure Functions v1 timer function using MSOnline and DAP
    (Connect-MsolService -Credential with an AES-encrypted stored password, Get-MsolPartnerContract,
    Get-MsolUser -TenantId) and basic authentication remote PowerShell to Exchange (DelegatedOrg), posting
    to a Microsoft Flow HTTP trigger.
    Required GDAP roles: Global Reader (users, licences and, with -IncludeMailboxType, mailboxes), plus
    Reports Reader or Security Reader for -IncludeAuthMethods and -IncludeSignInActivity.
    Required partner app permissions: Microsoft Graph delegated User.Read.All (User.ReadWrite.All is in the
    full manifest), LicenseAssignment.Read.All or Directory.ReadWrite.All (subscribedSkus), and
    AuditLog.Read.All for -IncludeAuthMethods and -IncludeSignInActivity. Office 365 Exchange Online
    delegated Exchange.Manage for -IncludeMailboxType.

.LINK
    https://gcit.com.au/knowledge-base/report-on-office-365-users-using-azure-functions-microsoft-flow-sharepoint-online/

.LINK
    docs/07-migrating-from-dap-msonline.md

.LINK
    docs/08-unattended-automation.md
#>
[CmdletBinding(DefaultParameterSetName = 'Tenant')]
[OutputType([pscustomobject])]
param(
    [Parameter(Mandatory, ParameterSetName = 'Tenant', ValueFromPipeline, ValueFromPipelineByPropertyName)]
    [Alias('CustomerTenantId')]
    [ValidateNotNullOrEmpty()]
    [string[]]$TenantId,

    [Parameter(Mandatory, ParameterSetName = 'All')]
    [switch]$AllCustomers,

    [switch]$LicensedOnly,

    [switch]$IncludeAuthMethods,

    [switch]$IncludeSignInActivity,

    [switch]$IncludeMailboxType,

    [string]$OutputPath
)

begin {
    $requested = [System.Collections.Generic.List[string]]::new()
    $results = [System.Collections.Generic.List[object]]::new()
    $userFields = 'id,displayName,userPrincipalName,mail,accountEnabled,userType,createdDateTime,lastPasswordChangeDateTime,onPremisesSyncEnabled,assignedLicenses,proxyAddresses,jobTitle,department,usageLocation'

    function ConvertTo-ResultRow {
        param($Customer, $User, [string[]]$Licences, $Registration, $MailboxType, $ErrorMessage)
        [pscustomobject][ordered]@{
            CustomerTenantId         = $Customer.TenantId
            CustomerName             = $Customer.Name
            UserId                   = $User.id
            DisplayName              = $User.displayName
            UserPrincipalName        = $User.userPrincipalName
            Mail                     = $User.mail
            Aliases                  = (@($User.proxyAddresses) | Where-Object { $_ -clike 'smtp:*' } | ForEach-Object { $_.Substring(5) }) -join '; '
            AccountEnabled           = $User.accountEnabled
            UserType                 = $User.userType
            CreatedDateTime          = $User.createdDateTime
            LastPasswordChange       = $User.lastPasswordChangeDateTime
            MailboxType              = $MailboxType
            OnPremisesSyncEnabled    = [bool]$User.onPremisesSyncEnabled
            JobTitle                 = $User.jobTitle
            Department               = $User.department
            UsageLocation            = $User.usageLocation
            Licences                 = ($Licences | Sort-Object) -join '; '
            IsMfaRegistered          = $Registration.isMfaRegistered
            MethodsRegistered        = (@($Registration.methodsRegistered)) -join '; '
            IsAdmin                  = $Registration.isAdmin
            LastSignInDateTime       = $User.signInActivity.lastSignInDateTime
            LastNonInteractiveSignIn = $User.signInActivity.lastNonInteractiveSignInDateTime
            Error                    = $ErrorMessage
        }
    }
}

process {
    foreach ($id in @($TenantId)) {
        if ($id) { $requested.Add($id) }
    }
}

end {
    $customers = if ($AllCustomers) {
        @(Get-MspCustomer | ForEach-Object { [pscustomobject]@{ TenantId = $_.TenantId; Name = $_.DisplayName } })
    }
    else {
        @($requested | Select-Object -Unique | ForEach-Object { [pscustomobject]@{ TenantId = $_; Name = $null } })
    }
    $select = if ($IncludeSignInActivity) { "$userFields,signInActivity" } else { $userFields }

    foreach ($customer in $customers) {
        try {
            if (-not $customer.Name) {
                $organisation = Invoke-MspGraphRequest -TenantId $customer.TenantId -Uri 'organization?$select=id,displayName' | Select-Object -First 1
                $customer.Name = $organisation.displayName
                if ($organisation.id) { $customer.TenantId = $organisation.id }
            }

            $skuNames = @{}
            Invoke-MspGraphRequest -TenantId $customer.TenantId -Uri 'subscribedSkus' |
                ForEach-Object { $skuNames[[string]$_.skuId] = $_.skuPartNumber }

            $registrations = @{}
            if ($IncludeAuthMethods) {
                Invoke-MspGraphRequest -TenantId $customer.TenantId -Uri 'reports/authenticationMethods/userRegistrationDetails' |
                    ForEach-Object { $registrations[[string]$_.id] = $_ }
            }

            # The original read RecipientTypeDetails from Exchange to tell user and shared mailboxes apart.
            $mailboxTypes = @{}
            if ($IncludeMailboxType) {
                $null = Connect-MspExchangeOnline -TenantId $customer.TenantId
                try {
                    Get-EXOMailbox -ResultSize Unlimited -Properties ExternalDirectoryObjectId, RecipientTypeDetails |
                        ForEach-Object { $mailboxTypes[[string]$_.ExternalDirectoryObjectId] = [string]$_.RecipientTypeDetails }
                }
                finally {
                    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
                }
            }

            $users = @(Invoke-MspGraphRequest -TenantId $customer.TenantId -Uri "users?`$select=$select&`$top=999")
            foreach ($user in $users) {
                $licences = @($user.assignedLicenses | ForEach-Object {
                        $name = $skuNames[[string]$_.skuId]
                        if ($name) { $name } else { [string]$_.skuId }
                    })
                if ($LicensedOnly -and $licences.Count -eq 0) { continue }
                $row = ConvertTo-ResultRow -Customer $customer -User $user -Licences $licences -Registration $registrations[[string]$user.id] -MailboxType $mailboxTypes[[string]$user.id]
                $results.Add($row)
                $row
            }
        }
        catch {
            Write-Warning "Customer $($customer.TenantId): $($_.Exception.Message)"
            $row = ConvertTo-ResultRow -Customer $customer -ErrorMessage $_.Exception.Message
            $results.Add($row)
            $row
        }
    }

    if ($OutputPath) {
        $results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding utf8
    }
}
