#Requires -Version 7.4
#Requires -Modules MspGdap

<#
.SYNOPSIS
    Connects to Security and Compliance PowerShell (Microsoft Purview) in a customer tenant through GDAP.

.DESCRIPTION
    The modern replacement for a "quick connect to the Security and Compliance Center" profile
    function.

    For each customer it connects with Connect-MspSecurityCompliance (the technician's GDAP token,
    no stored password), runs Get-RetentionCompliancePolicy to prove the session works, and
    returns one row per customer. The session is closed again before the next customer.

    Connect-MspSecurityCompliance was confirmed working with a delegated GDAP token in a live test on
    7 October 2026 (Get-RetentionCompliancePolicy). Microsoft documents Connect-IPPSSession -AccessToken
    only briefly, so try one customer first. If it fails, connect interactively instead (see .NOTES).

    With -KeepConnected and a single customer, the session is left open so you can keep working
    in it. Close it with Disconnect-ExchangeOnline or Disconnect-Msp when you finish.

.PARAMETER TenantId
    One or more customer tenant IDs (GUID) or verified domains, such as contoso.onmicrosoft.com.
    Accepts pipeline input, including objects from Get-MspCustomer.

.PARAMETER AllCustomers
    Test the connection to every customer returned by Get-MspCustomer that has an active GDAP
    relationship. Customers without one are listed with the status Skipped.

.PARAMETER KeepConnected
    Leave the Security and Compliance session open after the check. Only allowed with one customer.

.PARAMETER OutputPath
    Optional path of a CSV file to save the results to. The folder is created if needed.

.EXAMPLE
    ./connect-security-compliance.ps1 -TenantId 'contoso.onmicrosoft.com' -KeepConnected
    Get-DlpCompliancePolicy

    Connects to one customer's Security and Compliance PowerShell and leaves the session open.

.EXAMPLE
    ./connect-security-compliance.ps1 -AllCustomers -OutputPath 'C:\Reports\PurviewConnections.csv'

    Checks Security and Compliance PowerShell access in every GDAP customer.

.NOTES
    Replaces the original 2017 method: a PowerShell profile function (Connect-SecurityCompliance)
    that ran Get-Credential and opened a Basic authentication remote PowerShell session with
    New-PSSession -ConnectionUri https://ps.compliance.protection.outlook.com/powershell-liveid/.
    Required GDAP roles: Compliance Administrator in each customer.
    Required partner app permissions: Exchange.Manage (Office 365 Exchange Online). The delegated
    token is requested for https://ps.compliance.protection.outlook.com. The live test needed no extra
    consent or manifest entry for that audience (see docs/05).
    Also User.Read (Microsoft Graph, initial domain lookup), and in the partner tenant
    Directory.Read.All or Directory.ReadWrite.All plus DelegatedAdminRelationship.ReadWrite.All
    (Microsoft Graph, customer list for Get-MspCustomer).

    Interactive fallback with the ExchangeOnlineManagement module:
    Connect-IPPSSession -UserPrincipalName admin@contoso-msp.onmicrosoft.com -DelegatedOrganization contoso.onmicrosoft.com -AzureADAuthorizationEndpointUri 'https://login.microsoftonline.com/organizations'

.LINK
    https://gcit.com.au/knowledge-base/set-quick-connection-office-365-security-compliance-center-via-powershell/

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
        foreach ($column in 'Organization', 'ConnectedAs', 'RetentionPolicyCount', 'Experimental', 'KeptOpen', 'Status', 'Error') {
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
        throw '-KeepConnected needs exactly one customer, so compliance commands cannot run against the wrong tenant.'
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
            Write-Verbose "Connecting to Security and Compliance PowerShell for $($customer.Name) ($($customer.TenantId))."
            $connection = Connect-MspSecurityCompliance -TenantId $customer.TenantId -ErrorAction Stop
            $policies = @(Get-RetentionCompliancePolicy -ErrorAction Stop)
            $keepOpen = [bool]$KeepConnected
            $row = ConvertTo-ResultRow -Customer $customer -Data @{
                Organization         = $connection.Organization
                ConnectedAs          = $connection.UserPrincipalName
                RetentionPolicyCount = $policies.Count
                Experimental         = $connection.Experimental
                KeptOpen             = $keepOpen
                Status               = 'Connected'
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
