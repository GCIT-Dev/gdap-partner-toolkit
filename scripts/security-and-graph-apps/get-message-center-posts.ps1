#Requires -Version 7.4
#Requires -Modules MspGdap

<#
.SYNOPSIS
    Searches each customer's Microsoft 365 Message center and reports matching posts and their action dates.

.DESCRIPTION
    Reads Message center posts in each customer through the Microsoft Graph service communications API
    (GET /admin/serviceAnnouncement/messages) and returns the posts whose title or body contains
    -SearchText, with the services affected, the date by which action is required and the message text. Many posts are
    targeted at only some tenants, which is why each customer is queried separately.

    -ActionRequiredOnly limits the output to posts with an action required date. -Days limits it to posts
    changed in the last -Days days.

.PARAMETER TenantId
    One or more customer tenant IDs (GUID) or verified domains. Accepts pipeline input, including the
    output of Get-MspCustomer.

.PARAMETER AllCustomers
    Process every customer returned by Get-MspCustomer.

.PARAMETER SearchText
    Text to find in the post title or body. Leave it out to return every post.

.PARAMETER ActionRequiredOnly
    Only return posts that have an action required date.

.PARAMETER Days
    Only return posts changed in the last this many days. 0 (default) returns every post.

.PARAMETER OutputPath
    Optional path for a CSV export of the results.

.EXAMPLE
    ./get-message-center-posts.ps1 -AllCustomers -ActionRequiredOnly -Days 30 -OutputPath ./action-required.csv

    Lists every post changed in the last 30 days that needs action, across all customers.

.EXAMPLE
    ./get-message-center-posts.ps1 -TenantId 'contoso.onmicrosoft.com' -SearchText 'retire'

    Lists posts that mention a retirement in one customer.

.NOTES
    Replaces the original 2018 method: the AzureAD and AzureRM modules to create an app and add its service
    principal to the AdminAgents group (DAP), the v1 token endpoint with resource=https://manage.office.com,
    and the retired Office 365 Service Communications API (manage.office.com ServiceComms/Messages).
    Required GDAP roles: Message Center Reader (or Global Reader).
    Required partner app permissions: Microsoft Graph delegated ServiceMessage.Read.All (in the full
    manifest).
    Microsoft Learn: https://learn.microsoft.com/en-us/graph/api/serviceannouncement-list-messages

.LINK
    https://gcit.com.au/knowledge-base/retrieve-customers-office-365-message-center-info-via-powershell-and-office-365-management-api/

.LINK
    docs/04-preconsent-customers.md

.LINK
    docs/07-migrating-from-dap-msonline.md
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

    [string]$SearchText,

    [switch]$ActionRequiredOnly,

    [ValidateRange(0, 3650)]
    [int]$Days = 0,

    [string]$OutputPath
)

begin {
    $requested = [System.Collections.Generic.List[string]]::new()
    $results = [System.Collections.Generic.List[object]]::new()

    function ConvertTo-PlainText {
        # The original exported the message text. Graph returns it as HTML, so strip the tags for CSV.
        param([string]$Html)
        if (-not $Html) { return $null }
        $text = $Html -replace '<br\s*/?>|</p>|</li>', "`n" -replace '<[^>]+>', ''
        ([System.Net.WebUtility]::HtmlDecode($text) -replace '[ \t]+', ' ' -replace '(\s*\n){2,}', "`n").Trim()
    }

    function ConvertTo-ResultRow {
        param($Customer, $Message, $ErrorMessage)
        [pscustomobject][ordered]@{
            CustomerTenantId         = $Customer.TenantId
            CustomerName             = $Customer.Name
            MessageId                = $Message.id
            Title                    = $Message.title
            Category                 = $Message.category
            Severity                 = $Message.severity
            Services                 = (@($Message.services)) -join '; '
            IsMajorChange            = $Message.isMajorChange
            ActionRequiredByDateTime = $Message.actionRequiredByDateTime
            StartDateTime            = $Message.startDateTime
            LastModifiedDateTime     = $Message.lastModifiedDateTime
            Body                     = ConvertTo-PlainText -Html $Message.body.content
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
    $cutoff = if ($Days -gt 0) { [datetime]::UtcNow.AddDays(-$Days) } else { [datetime]::MinValue }

    foreach ($customer in $customers) {
        try {
            if (-not $customer.Name) {
                $organisation = Invoke-MspGraphRequest -TenantId $customer.TenantId -Uri 'organization?$select=id,displayName' | Select-Object -First 1
                $customer.Name = $organisation.displayName
                if ($organisation.id) { $customer.TenantId = $organisation.id }
            }

            $messages = @(Invoke-MspGraphRequest -TenantId $customer.TenantId -Uri 'admin/serviceAnnouncement/messages' -Headers @{ Prefer = 'odata.maxpagesize=1000' })
            foreach ($message in $messages) {
                if ($SearchText) {
                    $text = '{0} {1}' -f $message.title, $message.body.content
                    if ($text.IndexOf($SearchText, [System.StringComparison]::OrdinalIgnoreCase) -lt 0) { continue }
                }
                if ($ActionRequiredOnly -and -not $message.actionRequiredByDateTime) { continue }
                if ($Days -gt 0 -and $message.lastModifiedDateTime -and ([datetime]$message.lastModifiedDateTime).ToUniversalTime() -lt $cutoff) { continue }
                $row = ConvertTo-ResultRow -Customer $customer -Message $message
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
