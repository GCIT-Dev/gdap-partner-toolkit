#Requires -Version 7.4
#Requires -Modules MspGdap

<#
.SYNOPSIS
    Reports Microsoft 365 user activity and storage for customer tenants and optionally writes a summary to
    IT Glue.

.DESCRIPTION
    For each customer the script downloads these Microsoft 365 usage reports for -Period through Microsoft
    Graph (GET /reports/<report>(period='<Period>'), each returned as CSV) and joins them by user principal
    name:

    - getOffice365ActiveUserDetail for licences and the last activity date in Exchange, OneDrive,
      SharePoint and Teams,
    - getMailboxUsageDetail for mailbox size,
    - getOneDriveUsageAccountDetail for OneDrive size and file count,
    - getEmailActivityUserDetail for emails sent, received and read,
    - getTeamsUserActivityUserDetail for Teams chat messages, calls and meetings, and
    - getSharePointActivityUserDetail and getOneDriveActivityUserDetail for files viewed or edited.

    It returns one row per user. With -ITGlueFlexibleAssetTypeId and -Apply it also creates or updates one
    flexible asset per customer that holds the activity table (traits tenant-name, tenant-id, report-date
    and user-activity). The original wrote one asset per user, which needs far more API calls. Create a
    flexible asset type with those four traits first.

    Customers are matched to IT Glue organisations with -ITGlueOrganizationMapPath (a CSV with TenantId and
    OrganizationId columns). Without a map, the script matches verified domains against the email domains
    of IT Glue contacts, as the original did.

    If the customer conceals user details in reports, user principal names come back as hashed values. A
    Global Administrator in the customer can turn this off in the Microsoft 365 admin center. The Yammer and
    Office activations reports the original used are now Viva Engage and Microsoft 365 Apps reports, and are
    not included.

.PARAMETER TenantId
    One or more customer tenant IDs (GUID) or verified domains. Accepts pipeline input, including the
    output of Get-MspCustomer.

.PARAMETER AllCustomers
    Process every customer returned by Get-MspCustomer.

.PARAMETER Period
    Report period: D7, D30, D90 (default, as in the original) or D180.

.PARAMETER ITGlueFlexibleAssetTypeId
    ID of the user activity flexible asset type in IT Glue. Without it nothing is sent to IT Glue.

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
    ./sync-user-activity-itglue.ps1 -AllCustomers -Period D30 -OutputPath ./user-activity.csv

    Exports 30-day activity and storage for every user in every customer.

.EXAMPLE
    ./sync-user-activity-itglue.ps1 -TenantId 'contoso.onmicrosoft.com' -ITGlueFlexibleAssetTypeId 12345 -ITGlueOrganizationMapPath ./itglue-map.csv -Apply -WhatIf

    Shows whether the IT Glue asset for one customer would be created or updated.

.NOTES
    Replaces the original 2019 method: an AdminAgents (DAP) multi-tenant app created with the AzureAD and
    AzureRM modules, a client secret and IT Glue API key in plain text, the v1 token endpoint, the Graph
    contracts list and beta usage reports.
    Required GDAP roles: Reports Reader (or Global Reader).
    Required partner app permissions: Microsoft Graph delegated Reports.Read.All (in the full manifest) and
    User.Read (organisation).

.LINK
    https://gcit.com.au/knowledge-base/sync-office-365-user-activity-and-usage-with-it-glue/

.LINK
    docs/07-migrating-from-dap-msonline.md
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

    [ValidateSet('D7', 'D30', 'D90', 'D180')]
    [string]$Period = 'D90',

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

    function ConvertFrom-ReportContent {
        param([object]$Content)
        $text = if ($Content -is [byte[]]) { [System.Text.Encoding]::UTF8.GetString($Content) } else { [string]$Content }
        $text = $text.TrimStart([char]0xFEFF)
        if ([string]::IsNullOrWhiteSpace($text)) { return @() }
        @($text -split '\r?\n' | Where-Object { $_ } | ConvertFrom-Csv)
    }

    function Get-UsageReport {
        param([string]$Tenant, [string]$Report, [string]$KeyColumn, [string]$ReportPeriod)
        $map = @{}
        $content = Invoke-MspGraphRequest -TenantId $Tenant -Uri "reports/$Report(period='$ReportPeriod')" -NoPaging
        foreach ($line in (ConvertFrom-ReportContent -Content $content)) {
            $key = [string]$line.$KeyColumn
            if ($key) { $map[$key.ToLowerInvariant()] = $line }
        }
        $map
    }

    function ConvertTo-GigaByte {
        param([object]$Bytes)
        if ([string]::IsNullOrEmpty([string]$Bytes)) { return $null }
        [math]::Round([double]$Bytes / 1GB, 2)
    }

    function ConvertTo-HtmlCell {
        param([object]$Value)
        [System.Net.WebUtility]::HtmlEncode([string]$Value)
    }

    function ConvertTo-ResultRow {
        param($Customer, $User, $Mailbox, $OneDrive, $Email, $Teams, $SharePoint, $OneDriveActivity, $Action, $ErrorMessage)
        [pscustomobject][ordered]@{
            CustomerTenantId       = $Customer.TenantId
            CustomerName           = $Customer.Name
            UserPrincipalName      = $User.'User Principal Name'
            DisplayName            = $User.'Display Name'
            AssignedProducts       = $User.'Assigned Products'
            ExchangeLastActivity   = $User.'Exchange Last Activity Date'
            OneDriveLastActivity   = $User.'OneDrive Last Activity Date'
            SharePointLastActivity = $User.'SharePoint Last Activity Date'
            TeamsLastActivity      = $User.'Teams Last Activity Date'
            EmailsSent             = $Email.'Send Count'
            EmailsReceived         = $Email.'Receive Count'
            EmailsRead             = $Email.'Read Count'
            TeamsChatMessages      = if ($Teams) { [int]$Teams.'Team Chat Message Count' + [int]$Teams.'Private Chat Message Count' } else { $null }
            TeamsCalls             = $Teams.'Call Count'
            TeamsMeetings          = $Teams.'Meeting Count'
            SharePointFilesUsed    = $SharePoint.'Viewed Or Edited File Count'
            OneDriveFilesUsed      = $OneDriveActivity.'Viewed Or Edited File Count'
            MailboxStorageGB       = ConvertTo-GigaByte $Mailbox.'Storage Used (Byte)'
            MailboxItemCount       = $Mailbox.'Item Count'
            OneDriveStorageGB      = ConvertTo-GigaByte $OneDrive.'Storage Used (Byte)'
            OneDriveFileCount      = $OneDrive.'File Count'
            ReportRefreshDate      = $User.'Report Refresh Date'
            Action                 = $Action
            Error                  = $ErrorMessage
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
            $domains = @($organisation.verifiedDomains | ForEach-Object { [string]$_.name })

            $active = Get-UsageReport -Tenant $customer.TenantId -Report 'getOffice365ActiveUserDetail' -KeyColumn 'User Principal Name' -ReportPeriod $Period
            $mailboxes = Get-UsageReport -Tenant $customer.TenantId -Report 'getMailboxUsageDetail' -KeyColumn 'User Principal Name' -ReportPeriod $Period
            $oneDrives = Get-UsageReport -Tenant $customer.TenantId -Report 'getOneDriveUsageAccountDetail' -KeyColumn 'Owner Principal Name' -ReportPeriod $Period
            $email = Get-UsageReport -Tenant $customer.TenantId -Report 'getEmailActivityUserDetail' -KeyColumn 'User Principal Name' -ReportPeriod $Period
            $teams = Get-UsageReport -Tenant $customer.TenantId -Report 'getTeamsUserActivityUserDetail' -KeyColumn 'User Principal Name' -ReportPeriod $Period
            $sharePoint = Get-UsageReport -Tenant $customer.TenantId -Report 'getSharePointActivityUserDetail' -KeyColumn 'User Principal Name' -ReportPeriod $Period
            $oneDriveActivity = Get-UsageReport -Tenant $customer.TenantId -Report 'getOneDriveActivityUserDetail' -KeyColumn 'User Principal Name' -ReportPeriod $Period

            $organisationId = $null
            $action = 'Collected'
            $userRows = foreach ($key in ($active.Keys | Sort-Object)) {
                $user = $active[$key]
                if ($user.'Is Deleted' -eq 'True') { continue }
                ConvertTo-ResultRow -Customer $customer -User $user -Mailbox $mailboxes[$key] -OneDrive $oneDrives[$key] -Email $email[$key] -Teams $teams[$key] -SharePoint $sharePoint[$key] -OneDriveActivity $oneDriveActivity[$key]
            }
            $userRows = @($userRows)

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
                    $organisationId = @($domains | ForEach-Object { $contactDomains[$_.ToLowerInvariant()] } | Where-Object { $_ }) | Select-Object -First 1
                }
                if (-not $organisationId) {
                    $action = 'NoITGlueMatch'
                }
                else {
                    $tableRows = foreach ($userRow in $userRows) {
                        '<tr><td>{0}</td><td>{1}</td><td>{2}</td><td>{3}</td><td>{4}</td><td>{5}</td><td>{6}</td></tr>' -f (ConvertTo-HtmlCell $userRow.UserPrincipalName), (ConvertTo-HtmlCell $userRow.AssignedProducts), (ConvertTo-HtmlCell $userRow.ExchangeLastActivity), (ConvertTo-HtmlCell $userRow.TeamsLastActivity), (ConvertTo-HtmlCell $userRow.OneDriveLastActivity), (ConvertTo-HtmlCell $userRow.MailboxStorageGB), (ConvertTo-HtmlCell $userRow.OneDriveStorageGB)
                    }
                    $table = '<table class="table table-bordered table-hover"><thead><tr><th>User</th><th>Licences</th><th>Exchange last active</th><th>Teams last active</th><th>OneDrive last active</th><th>Mailbox GB</th><th>OneDrive GB</th></tr></thead><tbody>' + (@($tableRows) -join '') + '</tbody></table>'
                    $traits = [ordered]@{
                        'tenant-name'   = $customer.Name
                        'tenant-id'     = $customer.TenantId
                        'report-date'   = (@($userRows | ForEach-Object { $_.ReportRefreshDate } | Where-Object { $_ }) | Select-Object -First 1)
                        'user-activity' = $table
                    }
                    $existing = @(Invoke-ITGlueRequest -Method GET -Path "flexible_assets?filter[organization_id]=$organisationId&filter[flexible_asset_type_id]=$ITGlueFlexibleAssetTypeId") |
                        Where-Object { [string]$_.attributes.traits.'tenant-id' -eq [string]$customer.TenantId } | Select-Object -First 1
                    $verb = if ($existing) { 'Update' } else { 'Create' }
                    if (-not $Apply) {
                        $action = "Would$verb"
                    }
                    elseif ($PSCmdlet.ShouldProcess("IT Glue organisation $organisationId", "$verb user activity flexible asset for $($customer.Name)")) {
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

            foreach ($userRow in $userRows) {
                $userRow.Action = $action
                $results.Add($userRow)
                $userRow
            }
        }
        catch {
            Write-Warning "Customer $($customer.TenantId): $($_.Exception.Message)"
            $row = ConvertTo-ResultRow -Customer $customer -Action 'Failed' -ErrorMessage $_.Exception.Message
            $results.Add($row)
            $row
        }
    }

    if ($OutputPath) {
        $results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding utf8
    }
}
