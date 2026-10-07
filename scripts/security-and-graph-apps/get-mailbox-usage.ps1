#Requires -Version 7.4
#Requires -Modules MspGdap

<#
.SYNOPSIS
    Reports mailbox size against quota for customer tenants and flags mailboxes that are nearly full.

.DESCRIPTION
    Downloads the Microsoft 365 mailbox usage report for each customer through Microsoft Graph
    (GET /reports/getMailboxUsageDetail(period='D7'), returned as CSV) and returns one row per mailbox with
    storage used, the warning, prohibit send and prohibit send and receive quotas, deleted items, the
    percentage used and whether it is over -ThresholdPercent.

    If the customer conceals user details in reports (the default for many tenants), user principal names
    come back as hashed values and NamesConcealed is True. A Global Administrator in the customer can turn
    this off in the Microsoft 365 admin center (Settings, Org settings, Reports).

    The MailboxUsageReport Azure Function in the functions folder runs the script on a timer and queues a
    message for every mailbox over the threshold.

.PARAMETER TenantId
    One or more customer tenant IDs (GUID) or verified domains. Accepts pipeline input, including the
    output of Get-MspCustomer.

.PARAMETER AllCustomers
    Process every customer returned by Get-MspCustomer.

.PARAMETER ThresholdPercent
    Percentage of the prohibit send and receive quota at which a mailbox is flagged. Default 90.

.PARAMETER OverThresholdOnly
    Only return mailboxes at or over -ThresholdPercent.

.PARAMETER ExcludeUserPrincipalName
    Mailboxes to leave out, for example archive or journal mailboxes that are expected to be full. This
    replaces the "always ignore" option of the original SharePoint list.

.PARAMETER OutputPath
    Optional path for a CSV export of the results.

.EXAMPLE
    ./get-mailbox-usage.ps1 -AllCustomers -OverThresholdOnly -OutputPath ./nearly-full.csv

    Lists every mailbox at 90 per cent or more of its quota across all customers.

.EXAMPLE
    ./get-mailbox-usage.ps1 -TenantId 'contoso.onmicrosoft.com' -ThresholdPercent 80

    Lists every mailbox in one customer and flags those at 80 per cent or more.

.NOTES
    Replaces the original 2019 method: the AzureAD and AzureRM modules to create a multi-tenant app with a
    client secret and add its service principal to the AdminAgents group (DAP), the Graph contracts list
    and v1 token endpoint, and Azure Functions v1 (timer sync and HTTP acknowledgement functions).
    Required GDAP roles: Reports Reader (or Global Reader).
    Required partner app permissions: Microsoft Graph delegated Reports.Read.All (in the full manifest).

.LINK
    https://gcit.com.au/knowledge-base/sync-mailbox-usage-reports-with-a-sharepoint-list/

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

    [ValidateRange(1, 100)]
    [int]$ThresholdPercent = 90,

    [switch]$OverThresholdOnly,

    [string[]]$ExcludeUserPrincipalName,

    [string]$OutputPath
)

begin {
    $requested = [System.Collections.Generic.List[string]]::new()
    $results = [System.Collections.Generic.List[object]]::new()
    $threshold = $ThresholdPercent

    function ConvertFrom-ReportContent {
        param([object]$Content)
        $text = if ($Content -is [byte[]]) { [System.Text.Encoding]::UTF8.GetString($Content) } else { [string]$Content }
        $text = $text.TrimStart([char]0xFEFF)
        if ([string]::IsNullOrWhiteSpace($text)) { return @() }
        @($text -split '\r?\n' | Where-Object { $_ } | ConvertFrom-Csv)
    }

    function ConvertTo-GigaByte {
        param([object]$Bytes)
        if ([string]::IsNullOrEmpty([string]$Bytes)) { return $null }
        [math]::Round([double]$Bytes / 1GB, 2)
    }

    function ConvertTo-ResultRow {
        param($Customer, $Mailbox, $ErrorMessage)
        $used = if ($Mailbox) { [double]$Mailbox.'Storage Used (Byte)' } else { $null }
        $quota = if ($Mailbox -and $Mailbox.'Prohibit Send/Receive Quota (Byte)') { [double]$Mailbox.'Prohibit Send/Receive Quota (Byte)' } else { $null }
        $percent = if ($quota) { [math]::Round(100 * $used / $quota, 1) } else { $null }
        $upn = $Mailbox.'User Principal Name'
        [pscustomobject][ordered]@{
            CustomerTenantId           = $Customer.TenantId
            CustomerName               = $Customer.Name
            UserPrincipalName          = $upn
            DisplayName                = $Mailbox.'Display Name'
            RecipientType              = $Mailbox.'Recipient Type'
            StorageUsedGB              = if ($null -ne $used) { [math]::Round($used / 1GB, 2) } else { $null }
            ProhibitSendReceiveQuotaGB = if ($quota) { [math]::Round($quota / 1GB, 2) } else { $null }
            ProhibitSendQuotaGB        = ConvertTo-GigaByte $Mailbox.'Prohibit Send Quota (Byte)'
            IssueWarningQuotaGB        = ConvertTo-GigaByte $Mailbox.'Issue Warning Quota (Byte)'
            PercentUsed                = $percent
            OverThreshold              = if ($null -ne $percent) { $percent -ge $threshold } else { $null }
            ItemCount                  = $Mailbox.'Item Count'
            DeletedItemCount           = $Mailbox.'Deleted Item Count'
            DeletedItemSizeGB          = ConvertTo-GigaByte $Mailbox.'Deleted Item Size (Byte)'
            HasArchive                 = $Mailbox.'Has Archive'
            LastActivityDate           = $Mailbox.'Last Activity Date'
            NamesConcealed             = if ($upn) { [bool]($upn -match '^[0-9A-F]{32}$') } else { $null }
            Error                      = $ErrorMessage
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

    foreach ($customer in $customers) {
        try {
            if (-not $customer.Name) {
                $organisation = Invoke-MspGraphRequest -TenantId $customer.TenantId -Uri 'organization?$select=id,displayName' | Select-Object -First 1
                $customer.Name = $organisation.displayName
                if ($organisation.id) { $customer.TenantId = $organisation.id }
            }

            $content = Invoke-MspGraphRequest -TenantId $customer.TenantId -Uri "reports/getMailboxUsageDetail(period='D7')" -NoPaging
            $mailboxes = ConvertFrom-ReportContent -Content $content
            foreach ($mailbox in $mailboxes) {
                if ($mailbox.'Is Deleted' -eq 'True') { continue }
                if ($ExcludeUserPrincipalName -and $ExcludeUserPrincipalName -contains [string]$mailbox.'User Principal Name') { continue }
                $row = ConvertTo-ResultRow -Customer $customer -Mailbox $mailbox
                if ($OverThresholdOnly -and -not $row.OverThreshold) { continue }
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
