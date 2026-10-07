#Requires -Version 7.4
#Requires -Modules MspGdap

<#
.SYNOPSIS
    Reports Microsoft Secure Score for customer tenants, optionally with every control.

.DESCRIPTION
    Reads the latest Microsoft Secure Score in each customer through Microsoft Graph
    (GET /security/secureScores?$top=1) and returns one row per customer with the current score, the
    maximum score, the percentage, the score for each category (Identity, Data, Device, Apps,
    Infrastructure) and the averages for tenants of a similar size and for all tenants.

    With -IncludeControls the script returns one row per Secure Score control instead, joined with
    GET /security/secureScoreControlProfiles for the title, maximum score, rank and remediation link.

    There is no need to create an app in each customer, give it a client secret, or call Azure AD Graph.
    The MspGdap partner app is consented once per customer, and the technician's GDAP role sets the limit.

.PARAMETER TenantId
    One or more customer tenant IDs (GUID) or verified domains. Accepts pipeline input, including the
    output of Get-MspCustomer.

.PARAMETER AllCustomers
    Process every customer returned by Get-MspCustomer.

.PARAMETER IncludeControls
    Return one row per Secure Score control instead of one row per customer.

.PARAMETER OutputPath
    Optional path for a CSV export of the results.

.EXAMPLE
    ./get-secure-scores.ps1 -AllCustomers -OutputPath ./secure-scores.csv

    Exports the current Secure Score for every customer.

.EXAMPLE
    ./get-secure-scores.ps1 -TenantId 'contoso.onmicrosoft.com' -IncludeControls | Where-Object Score -lt MaxScore

    Lists the controls one customer has not fully implemented.

.NOTES
    Replaces the original 2018 method: the AzureAD and AzureRM modules (New-AzureADApplication,
    New-AzureRmADServicePrincipal) to create an app with a client secret in each tenant, Azure AD Graph
    oauth2PermissionGrants, the v1 token endpoint with resource= and optionally the password (ROPC) grant,
    and the removed beta report reports/getTenantSecureScores.
    Required GDAP roles: Security Reader (or Global Reader).
    Required partner app permissions: Microsoft Graph delegated SecurityEvents.Read.All
    (SecurityEvents.ReadWrite.All is in the full manifest).

.LINK
    https://gcit.com.au/knowledge-base/automate-creation-azure-ad-applications-access-microsoft-graph-customer-tenants/

.LINK
    https://gcit.com.au/knowledge-base/automate-api-calls-microsoft-graph-using-powershell-azure-active-directory-applications/

.LINK
    docs/06-token-caching-and-validation.md

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

    [switch]$IncludeControls,

    [string]$OutputPath
)

begin {
    $requested = [System.Collections.Generic.List[string]]::new()
    $results = [System.Collections.Generic.List[object]]::new()

    function ConvertTo-ScoreRow {
        param($Customer, $Score, $ErrorMessage)
        $current = if ($Score) { [double]$Score.currentScore } else { $null }
        $max = if ($Score) { [double]$Score.maxScore } else { $null }
        $similar = @($Score.averageComparativeScores | Where-Object { $_.basis -eq 'TotalSeats' }) | Select-Object -First 1
        $allTenants = @($Score.averageComparativeScores | Where-Object { $_.basis -eq 'AllTenants' }) | Select-Object -First 1
        # The original exported account, data and device scores. Secure Score now reports them per control,
        # so add up the control scores for each category.
        $category = @{}
        foreach ($control in @($Score.controlScores)) {
            $name = [string]$control.controlCategory
            if (-not $name) { continue }
            $category[$name] = [double]$category[$name] + [double]$control.score
        }
        [pscustomobject][ordered]@{
            CustomerTenantId      = $Customer.TenantId
            CustomerName          = $Customer.Name
            ScoreDate             = $Score.createdDateTime
            CurrentScore          = $current
            MaxScore              = $max
            Percentage            = if ($max) { [math]::Round(100 * $current / $max, 1) } else { $null }
            IdentityScore         = if ($Score) { [math]::Round([double]$category['Identity'], 2) } else { $null }
            DataScore             = if ($Score) { [math]::Round([double]$category['Data'], 2) } else { $null }
            DeviceScore           = if ($Score) { [math]::Round([double]$category['Device'], 2) } else { $null }
            AppsScore             = if ($Score) { [math]::Round([double]$category['Apps'], 2) } else { $null }
            InfrastructureScore   = if ($Score) { [math]::Round([double]$category['Infrastructure'], 2) } else { $null }
            SimilarTenantsAverage = $similar.averageScore
            AllTenantsAverage     = $allTenants.averageScore
            LicensedUserCount     = $Score.licensedUserCount
            ActiveUserCount       = $Score.activeUserCount
            EnabledServices       = (@($Score.enabledServices)) -join '; '
            Error                 = $ErrorMessage
        }
    }

    function ConvertTo-ControlRow {
        param($Customer, $Control, $ControlProfile, $ErrorMessage)
        [pscustomobject][ordered]@{
            CustomerTenantId   = $Customer.TenantId
            CustomerName       = $Customer.Name
            ControlName        = $Control.controlName
            Title              = $ControlProfile.title
            Category           = $Control.controlCategory
            Score              = if ($Control) { [double]$Control.score } else { $null }
            MaxScore           = $ControlProfile.maxScore
            Rank               = $ControlProfile.rank
            UserImpact         = $ControlProfile.userImpact
            ImplementationCost = $ControlProfile.implementationCost
            ActionUrl          = $ControlProfile.actionUrl
            Error              = $ErrorMessage
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

            $score = Invoke-MspGraphRequest -TenantId $customer.TenantId -Uri 'security/secureScores?$top=1' -MaxPages 1 | Select-Object -First 1
            if (-not $score) { throw 'No Secure Score was returned for this tenant.' }

            if ($IncludeControls) {
                $profiles = @{}
                Invoke-MspGraphRequest -TenantId $customer.TenantId -Uri 'security/secureScoreControlProfiles' |
                    ForEach-Object { $profiles[[string]$_.id] = $_ }
                foreach ($control in @($score.controlScores)) {
                    $row = ConvertTo-ControlRow -Customer $customer -Control $control -ControlProfile $profiles[[string]$control.controlName]
                    $results.Add($row)
                    $row
                }
            }
            else {
                $row = ConvertTo-ScoreRow -Customer $customer -Score $score
                $results.Add($row)
                $row
            }
        }
        catch {
            Write-Warning "Customer $($customer.TenantId): $($_.Exception.Message)"
            $row = if ($IncludeControls) {
                ConvertTo-ControlRow -Customer $customer -ErrorMessage $_.Exception.Message
            }
            else {
                ConvertTo-ScoreRow -Customer $customer -ErrorMessage $_.Exception.Message
            }
            $results.Add($row)
            $row
        }
    }

    if ($OutputPath) {
        $results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding utf8
    }
}
