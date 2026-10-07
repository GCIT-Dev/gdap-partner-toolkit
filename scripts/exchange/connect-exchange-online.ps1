#Requires -Version 7.4
#Requires -Modules MspGdap

<#
.SYNOPSIS
    Connects to Exchange Online in a customer tenant through GDAP, and checks the connection.

.DESCRIPTION
    The modern replacement for a "quick connect to Exchange Online" profile function.

    For each customer it connects with Connect-MspExchangeOnline (the technician's GDAP token,
    modern authentication, no stored password), reads the organisation configuration to prove
    the session works, and returns one row per customer. The session is closed again before
    the next customer.

    With -KeepConnected and a single customer, the session is left open so you can keep
    working in it, which is what the original profile function was for. Close it with
    Disconnect-ExchangeOnline or Disconnect-Msp when you finish.

.PARAMETER TenantId
    One or more customer tenant IDs (GUID) or verified domains, such as contoso.onmicrosoft.com.
    Accepts pipeline input, including objects from Get-MspCustomer.

.PARAMETER AllCustomers
    Test the connection to every customer returned by Get-MspCustomer that has an active GDAP
    relationship. Customers without one are listed with the status Skipped.

.PARAMETER KeepConnected
    Leave the Exchange Online session open after the check. Only allowed with one customer, so
    commands can never land in the wrong tenant.

.PARAMETER OutputPath
    Optional path of a CSV file to save the results to. The folder is created if needed.

.EXAMPLE
    ./connect-exchange-online.ps1 -TenantId 'contoso.onmicrosoft.com' -KeepConnected
    Get-EXOMailbox -ResultSize 10

    Connects to one customer and leaves the session open for interactive work.

.EXAMPLE
    ./connect-exchange-online.ps1 -AllCustomers -OutputPath 'C:\Reports\ExchangeConnections.csv'

    Checks that Exchange Online access works in every GDAP customer, and saves the results.

.EXAMPLE
    # Add a short profile function, like the original article did (run: notepad $PROFILE)
    function Connect-CustomerExchange { param([string]$Tenant) & 'C:\Scripts\connect-exchange-online.ps1' -TenantId $Tenant -KeepConnected }

    Recreates the original quick-connect shortcut on top of this script.

.NOTES
    Replaces the original 2015 method: a PowerShell profile function (Connect-EXOnline) that ran
    Get-Credential and opened a Basic authentication remote PowerShell session with
    New-PSSession -ConnectionUri https://outlook.office365.com/powershell-liveid/ and Import-PSSession.
    Required GDAP roles: Global Reader for the connection check, Exchange Administrator or
    Exchange Recipient Administrator for the work you do afterwards.
    Required partner app permissions: Exchange.Manage (Office 365 Exchange Online), User.Read
    (Microsoft Graph, initial domain lookup), and in the partner tenant Directory.Read.All or
    Directory.ReadWrite.All plus DelegatedAdminRelationship.ReadWrite.All (Microsoft Graph,
    customer list for Get-MspCustomer).

    For your own tenant (not a customer), run Connect-ExchangeOnline -UserPrincipalName
    admin@contoso.onmicrosoft.com from the ExchangeOnlineManagement module instead.

.LINK
    https://gcit.com.au/knowledge-base/how-to-set-up-a-quick-connection-to-exchange-online-via-powershell/

.LINK
    ../../docs/05-exchange-access.md
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

    [Parameter(ParameterSetName = 'Tenant')]
    [switch]$KeepConnected,

    [string]$OutputPath
)

begin {
    function Get-TargetCustomer {
        param([string[]]$Tenant, [switch]$All)
        if ($All) {
            foreach ($customer in @(Get-MspCustomer -IncludeGdapStatus -ErrorAction Stop)) {
                $skip = $null
                if ($customer.GdapStatus -ne 'active') { $skip = "No active GDAP relationship (status: $($customer.GdapStatus))." }
                [pscustomobject]@{ TenantId = $customer.TenantId; Name = $customer.DisplayName; Skip = $skip }
            }
            return
        }
        foreach ($entry in $Tenant) {
            $id = $entry
            $name = $entry
            try {
                $customer = @(Get-MspCustomer -TenantId $entry -IncludeGdapStatus -ErrorAction Stop) | Select-Object -First 1
                if ($customer) {
                    $id = $customer.TenantId
                    $name = $customer.DisplayName
                }
            }
            catch {
                Write-Verbose "Could not look up the customer name for $entry. $($_.Exception.Message)"
            }
            [pscustomobject]@{ TenantId = $id; Name = $name; Skip = $null }
        }
    }

    function Close-CustomerSession {
        param([object]$Connection)
        if (-not (Get-Command -Name 'Disconnect-ExchangeOnline' -ErrorAction SilentlyContinue)) { return }
        $disconnect = @{ Confirm = $false; WhatIf = $false; ErrorAction = 'SilentlyContinue' }
        if ($Connection -and $Connection.ConnectionId) { $disconnect['ConnectionId'] = $Connection.ConnectionId }
        Disconnect-ExchangeOnline @disconnect
    }

    function ConvertTo-ResultRow {
        param([object]$Customer, [hashtable]$Data)
        $row = [ordered]@{ CustomerTenantId = $Customer.TenantId; CustomerName = $Customer.Name }
        foreach ($column in 'Organization', 'OrganizationDisplayName', 'IsDehydrated', 'ConnectedAs', 'TokenExpiryTimeUTC', 'KeptOpen', 'Status', 'Error') {
            $row[$column] = $Data[$column]
        }
        [pscustomobject]$row
    }

    $tenantInput = [System.Collections.Generic.List[string]]::new()
    $results = [System.Collections.Generic.List[object]]::new()
}

process {
    foreach ($entry in $TenantId) {
        if ($entry) { $tenantInput.Add($entry) }
    }
}

end {
    if ($KeepConnected -and $tenantInput.Count -ne 1) {
        throw '-KeepConnected needs exactly one customer, so Exchange commands cannot run against the wrong tenant.'
    }

    $customers = @(Get-TargetCustomer -Tenant $tenantInput -All:$AllCustomers)
    foreach ($customer in $customers) {
        if ($customer.Skip) {
            $row = ConvertTo-ResultRow -Customer $customer -Data @{ Status = 'Skipped'; Error = $customer.Skip }
            $results.Add($row)
            $row
            continue
        }

        $connection = $null
        $keepOpen = $false
        try {
            Write-Verbose "Connecting to Exchange Online for $($customer.Name) ($($customer.TenantId))."
            $connection = Connect-MspExchangeOnline -TenantId $customer.TenantId -ErrorAction Stop
            $organization = Get-OrganizationConfig -ErrorAction Stop
            $keepOpen = [bool]$KeepConnected
            $row = ConvertTo-ResultRow -Customer $customer -Data @{
                Organization            = $connection.Organization
                OrganizationDisplayName = $organization.DisplayName
                IsDehydrated            = $organization.IsDehydrated
                ConnectedAs             = $connection.UserPrincipalName
                TokenExpiryTimeUTC      = $connection.TokenExpiryTimeUTC
                KeptOpen                = $keepOpen
                Status                  = 'Connected'
            }
            $results.Add($row)
            $row
        }
        catch {
            $keepOpen = $false
            Write-Warning "$($customer.Name) ($($customer.TenantId)) failed: $($_.Exception.Message)"
            $row = ConvertTo-ResultRow -Customer $customer -Data @{ Status = 'Failed'; Error = $_.Exception.Message }
            $results.Add($row)
            $row
        }
        finally {
            if (-not $keepOpen) { Close-CustomerSession -Connection $connection }
        }
    }

    if ($OutputPath) {
        $folder = Split-Path -Path $OutputPath -Parent
        if ($folder -and -not (Test-Path -LiteralPath $folder)) { $null = New-Item -ItemType Directory -Path $folder -Force -WhatIf:$false }
        $results | Export-Csv -LiteralPath $OutputPath -NoTypeInformation -Encoding utf8 -WhatIf:$false
        Write-Verbose "Saved $($results.Count) row(s) to $OutputPath."
    }
}
