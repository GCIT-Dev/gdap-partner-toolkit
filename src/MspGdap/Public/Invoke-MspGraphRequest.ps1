function Invoke-MspGraphRequest {
    <#
    .SYNOPSIS
        Calls Microsoft Graph in a customer tenant (or your partner tenant) with
        cached delegated tokens, paging and throttling handling.
    .DESCRIPTION
        - Gets the bearer token through Get-MspAccessToken, so a valid cached
          token is reused and a refresh happens only when needed.
        - GET collections are followed through @odata.nextLink and each item
          in value is written to the pipeline. -NoPaging stops after the
          first page and still writes its items.
        - Responses that are not collections (a single entity, report CSV
          content, $count text) are returned unchanged.
        - HTTP 429, 503 and 504 are retried after Retry-After, or with
          exponential backoff when no Retry-After is given.
        - HTTP 401 triggers one forced token refresh and retry.
        - POST, PATCH, PUT and DELETE support -WhatIf and -Confirm.
        - The token is only ever sent to https://graph.microsoft.com,
          including nextLink URLs.
    .PARAMETER TenantId
        Customer tenant GUID or verified domain. Mandatory unless -PartnerTenant is used.
    .PARAMETER PartnerTenant
        Call Graph in your own partner tenant deliberately.
    .PARAMETER Uri
        Relative path (users, /groups?$top=999) or absolute Graph URI.
    .PARAMETER Method
        GET (default), POST, PATCH, PUT or DELETE.
    .PARAMETER Body
        Request body. Objects and hashtables are converted to JSON.
    .PARAMETER ApiVersion
        v1.0 (default) or beta. Ignored when -Uri already names a version.
    .PARAMETER Headers
        Extra headers such as @{ ConsistencyLevel = 'eventual' }.
    .PARAMETER NoPaging
        Read the first page only. For a GET collection (a response with a value
        array) the items of that page are written to the pipeline, the same as
        paged mode, and @odata.nextLink is not followed (a Verbose message says
        when more pages exist). Responses that are not collections are returned
        unchanged.
    .PARAMETER MaxPages
        Stop after this many pages. 0 (default) means no limit.
    .PARAMETER MaxRetries
        Retries for throttled or unavailable responses.
    .PARAMETER UserPrincipalName
        Technician whose refresh token is used.
    .EXAMPLE
        Invoke-MspGraphRequest -TenantId 'contoso.onmicrosoft.com' -Uri 'users?$select=id,userPrincipalName'
    .EXAMPLE
        Invoke-MspGraphRequest -TenantId $tid -Uri "users/$userId" -Method PATCH -Body @{ usageLocation = 'AU' } -WhatIf
    .EXAMPLE
        Invoke-MspGraphRequest -PartnerTenant -Uri 'tenantRelationships/delegatedAdminRelationships'
    #>
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium', DefaultParameterSetName = 'Tenant')]
    [OutputType([object])]
    param(
        [Parameter(Mandatory, ParameterSetName = 'Tenant')]
        [ValidateNotNullOrEmpty()]
        [string]$TenantId,

        [Parameter(Mandatory, ParameterSetName = 'Partner')]
        [switch]$PartnerTenant,

        [Parameter(Mandatory, Position = 0)]
        [ValidateNotNullOrEmpty()]
        [string]$Uri,

        [ValidateSet('GET', 'POST', 'PATCH', 'PUT', 'DELETE')]
        [string]$Method = 'GET',

        [object]$Body,

        [ValidateSet('v1.0', 'beta')]
        [string]$ApiVersion = 'v1.0',

        [hashtable]$Headers,

        [switch]$NoPaging,

        [ValidateRange(0, 100000)]
        [int]$MaxPages = 0,

        [ValidateRange(0, 10)]
        [int]$MaxRetries = 5,

        [string]$UserPrincipalName
    )

    try {
        $config = Get-MspConfigurationInternal -RequireComplete
        $tenant = Resolve-MspTargetTenant -TenantId $TenantId -PartnerTenant:$PartnerTenant -Configuration $config
        $target = Resolve-MspGraphUri -Uri $Uri -ApiVersion $ApiVersion

        if ($Method -ne 'GET') {
            $label = if ($PartnerTenant) { "partner tenant $tenant" } else { "tenant $tenant" }
            if (-not $PSCmdlet.ShouldProcess($label, "Graph $Method $($target.PathAndQuery)")) {
                return
            }
        }

        $tokenParams = @{ Resource = 'Graph' }
        if ($UserPrincipalName) { $tokenParams.UserPrincipalName = $UserPrincipalName }
        if ($PartnerTenant) { $tokenParams.PartnerTenant = $true } else { $tokenParams.TenantId = $tenant }

        $next = $target
        $pages = 0
        while ($next) {
            $response = $null
            $refreshed = $false
            while ($true) {
                $requestHeaders = Get-MspAuthHeader @tokenParams -ForceRefresh:$refreshed -AdditionalHeaders $Headers
                $restParams = @{
                    Uri         = $next
                    Method      = $Method
                    Headers     = $requestHeaders
                    MaxRetries  = $MaxRetries
                    ServiceName = 'Graph'
                }
                if ($PSBoundParameters.ContainsKey('Body') -and $null -ne $Body) { $restParams.Body = $Body }
                try {
                    $response = Invoke-MspRestWithRetry @restParams
                    break
                }
                catch {
                    $status = $_.Exception.Data['StatusCode']
                    if ($status -eq 401 -and -not $refreshed) {
                        Write-Verbose 'Graph returned 401. Refreshing the token once and retrying.'
                        $refreshed = $true
                        continue
                    }
                    throw
                }
                finally {
                    $requestHeaders = $null
                    $restParams = $null
                }
            }
            $pages++

            $isPage = $Method -eq 'GET' -and $response -is [pscustomobject] -and $response.PSObject.Properties['value']
            if ($isPage -and $NoPaging) {
                # -NoPaging: the first page's items when value is an array. Anything else is returned as is.
                $isPage = $response.value -is [System.Collections.IList]
            }
            if ($isPage) {
                foreach ($item in @($response.value)) { $item }
                $nextLink = if ($response.PSObject.Properties['@odata.nextLink']) { [string]$response.'@odata.nextLink' } else { $null }
                if ($NoPaging) {
                    if ($nextLink) { Write-Verbose 'Graph returned more pages (@odata.nextLink). -NoPaging returned the first page only.' }
                    $next = $null
                }
                elseif ($nextLink -and ($MaxPages -eq 0 -or $pages -lt $MaxPages)) {
                    $next = Resolve-MspGraphUri -Uri $nextLink
                }
                else {
                    $next = $null
                }
            }
            else {
                if ($null -ne $response) { $response }
                $next = $null
            }
        }
    }
    catch {
        $PSCmdlet.ThrowTerminatingError($_)
    }
}
