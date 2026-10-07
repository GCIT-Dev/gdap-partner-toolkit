#Requires -Version 7.4
#Requires -Modules MspGdap
#Requires -Modules Microsoft.PowerShell.SecretManagement

<#
.SYNOPSIS
    Checks the email addresses of every user in one or more customer tenants against Have I Been Pwned and exports the breaches found.

.DESCRIPTION
    For each customer the script reads the member users through Microsoft Graph
    (guests are skipped), takes each user's SMTP proxy addresses (excluding
    .onmicrosoft.com addresses) and looks each one up with the Have I Been Pwned
    (HIBP) breachedaccount API. It returns one row per breach, with the breach
    details (including HIBP's description, which contains HTML) and whether the user's password has been changed since HIBP added the
    breach (PasswordChangedSinceBreach). A customer with no breaches gets one row
    with Status NoBreachesFound and the number of addresses checked.

    The HIBP API key is read at run time from a SecretManagement vault, never from
    the script. Store it once with:
        Set-Secret -Name 'HibpApiKey' -Vault 'MspGdap' -SecureStringSecret (Read-Host -AsSecureString -Prompt 'HIBP API key')

    HIBP rate limits depend on your subscription. -DelayMilliseconds (default
    6500, which suits the lowest tier of 10 requests a minute) is the pause between
    lookups, and a 429 response is retried after the Retry-After time HIBP returns.
    For large tenants, HIBP's domain search for your customers' verified domains is
    cheaper and has no per-address limit.

    The script only reads. It never changes a tenant.

.PARAMETER TenantId
    One or more customer tenant IDs (GUID) or verified domains, such as
    contoso.onmicrosoft.com. Accepts pipeline input, including objects from
    Get-MspCustomer.

.PARAMETER AllCustomers
    Runs against every customer with an active GDAP relationship, as returned by
    Get-MspCustomer -IncludeGdapStatus.

.PARAMETER ApiKeySecretName
    Name of the SecretManagement secret that holds the HIBP API key. Default
    HibpApiKey.

.PARAMETER VaultName
    SecretManagement vault that holds the secret. Defaults to your default vault.

.PARAMETER DelayMilliseconds
    Pause between HIBP lookups. Default 6500.

.PARAMETER OutputPath
    Optional path of a CSV file for the results. The file is overwritten.

.EXAMPLE
    ./check-hibp-breaches.ps1 -TenantId 'contoso.onmicrosoft.com' -VaultName 'MspGdap' -OutputPath ./contoso-breaches.csv

    Checks every member user's addresses in one customer and saves the breaches to CSV.

.EXAMPLE
    ./check-hibp-breaches.ps1 -AllCustomers -DelayMilliseconds 1600 | Where-Object { $_.Status -eq 'Breached' -and -not $_.PasswordChangedSinceBreach }

    Checks every active GDAP customer on a higher HIBP tier and shows the accounts whose password is older than the breach.

.NOTES
    Replaces the original 2018 method: Connect-MsolService, Get-MsolPartnerContract -All (DAP) and Get-MsolUser -TenantId -All (MSOnline module, retired 30 May 2025), an HIBP API key typed into the script, and a TLS 1.2 workaround for Windows PowerShell 5.1.
    Required GDAP roles: Directory Readers or Global Reader.
    Required partner app permissions: Microsoft Graph delegated User.Read.All (covered in manifests/partner-app.full.json by User.ReadWrite.All).

.LINK
    https://gcit.com.au/knowledge-base/check-office-365-accounts-against-have-i-been-pwned-breaches/

.LINK
    https://haveibeenpwned.com/API/v3

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
    [string]$ApiKeySecretName = 'HibpApiKey',

    [string]$VaultName,

    [ValidateRange(0, 600000)]
    [int]$DelayMilliseconds = 6500,

    [ValidateNotNullOrEmpty()]
    [string]$OutputPath
)

begin {
    $columns = @('CustomerTenantId', 'CustomerName', 'UserPrincipalName', 'Email', 'LastPasswordChange', 'BreachName', 'BreachTitle', 'BreachDate', 'BreachAdded', 'Description', 'DataClasses', 'IsVerified', 'IsFabricated', 'IsSensitive', 'IsRetired', 'IsSpamList', 'PasswordChangedSinceBreach', 'Status', 'Error')
    $targets = [System.Collections.Generic.List[string]]::new()
    $knownNames = @{}
    $results = [System.Collections.Generic.List[object]]::new()
    $baseUri = 'https://haveibeenpwned.com/api/v3/breachedaccount'

    $secretParams = @{ Name = $ApiKeySecretName; AsPlainText = $true; ErrorAction = 'Stop' }
    if ($VaultName) { $secretParams.Vault = $VaultName }
    $hibpHeaders = @{
        'hibp-api-key' = Get-Secret @secretParams
        'user-agent'   = 'MspGdap-HIBP-check'
    }

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

    function Get-HibpBreach {
        param([string]$Address, [hashtable]$Headers, [string]$Uri, [int]$MaxRetries = 3)
        $target = '{0}/{1}?truncateResponse=false' -f $Uri, [uri]::EscapeDataString($Address)
        for ($attempt = 0; $attempt -le $MaxRetries; $attempt++) {
            try {
                return @(Invoke-RestMethod -Method Get -Uri $target -Headers $Headers -ErrorAction Stop)
            }
            catch {
                $response = $_.Exception.Response
                $status = if ($response) { [int]$response.StatusCode } else { 0 }
                if ($status -eq 404) { return @() }
                if ($status -eq 429 -and $attempt -lt $MaxRetries) {
                    $wait = 10
                    if ($response.Headers.RetryAfter -and $response.Headers.RetryAfter.Delta) {
                        $wait = [math]::Ceiling($response.Headers.RetryAfter.Delta.TotalSeconds) + 1
                    }
                    Write-Verbose -Message "HIBP rate limit reached. Waiting $wait seconds."
                    Start-Sleep -Seconds $wait
                    continue
                }
                throw
            }
        }
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

    try {
        foreach ($target in $targets) {
            $tenant = $target
            $customerName = $null
            try {
                $tenant = Resolve-MspTenantId -Tenant $target
                $customerName = $knownNames[$tenant]
                if (-not $customerName) {
                    $customerName = @(Invoke-MspGraphRequest -TenantId $tenant -Uri 'v1.0/organization?$select=displayName')[0].displayName
                }

                $users = @(Invoke-MspGraphRequest -TenantId $tenant -Uri 'v1.0/users?$select=id,userPrincipalName,userType,proxyAddresses,lastPasswordChangeDateTime&$top=999' |
                        Where-Object { $_.userType -ne 'Guest' -and $_.userPrincipalName -notmatch '#EXT#' })

                $checked = 0
                $customerRows = $results.Count
                foreach ($user in $users) {
                    $addresses = @($user.proxyAddresses | Where-Object { $_ -match '^smtp:' -and $_ -notmatch '\.onmicrosoft\.com$' } | ForEach-Object { ($_ -split ':', 2)[1] } | Sort-Object -Unique)
                    foreach ($address in $addresses) {
                        try {
                            $breaches = @(Get-HibpBreach -Address $address -Headers $hibpHeaders -Uri $baseUri)
                            foreach ($breach in $breaches) {
                                $changed = $null
                                if ($user.lastPasswordChangeDateTime -and $breach.AddedDate) {
                                    $changed = ([datetime]$user.lastPasswordChangeDateTime) -gt ([datetime]$breach.AddedDate)
                                }
                                $row = Get-ResultRow -Column $columns -Value @{
                                    CustomerTenantId           = $tenant
                                    CustomerName               = $customerName
                                    UserPrincipalName          = $user.userPrincipalName
                                    Email                      = $address
                                    LastPasswordChange         = $user.lastPasswordChangeDateTime
                                    BreachName                 = $breach.Name
                                    BreachTitle                = $breach.Title
                                    BreachDate                 = $breach.BreachDate
                                    BreachAdded                = $breach.AddedDate
                                    Description                = $breach.Description
                                    DataClasses                = @($breach.DataClasses) -join ', '
                                    IsVerified                 = $breach.IsVerified
                                    IsFabricated               = $breach.IsFabricated
                                    IsSensitive                = $breach.IsSensitive
                                    IsRetired                  = $breach.IsRetired
                                    IsSpamList                 = $breach.IsSpamList
                                    PasswordChangedSinceBreach = $changed
                                    Status                     = 'Breached'
                                }
                                $results.Add($row)
                                $row
                            }
                        }
                        catch {
                            Write-Warning -Message "HIBP lookup for $address failed: $($_.Exception.Message)"
                            $row = Get-ResultRow -Column $columns -Value @{
                                CustomerTenantId  = $tenant
                                CustomerName      = $customerName
                                UserPrincipalName = $user.userPrincipalName
                                Email             = $address
                                Status            = 'LookupFailed'
                                Error             = $_.Exception.Message
                            }
                            $results.Add($row)
                            $row
                        }
                        $checked++
                        if ($DelayMilliseconds -gt 0) { Start-Sleep -Milliseconds $DelayMilliseconds }
                    }
                }

                if ($results.Count -eq $customerRows) {
                    # Nothing found, so say so (the original printed "No Breach detected").
                    $row = Get-ResultRow -Column $columns -Value @{
                        CustomerTenantId = $tenant
                        CustomerName     = $customerName
                        Email            = "$checked address(es) checked"
                        Status           = 'NoBreachesFound'
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
    }
    finally {
        $hibpHeaders = $null
    }

    if ($OutputPath -and $results.Count -gt 0) {
        $results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding utf8 -WhatIf:$false -Confirm:$false
    }
}
