function Invoke-MspGraphBatch {
    <#
    .SYNOPSIS
        Sends many Graph requests through JSON batching ($batch) in one tenant.
    .DESCRIPTION
        Requests are sent in chunks of 20 (the Graph limit). A 200 response for
        the batch does not mean each request succeeded, so one result object is
        returned per request with its own Status and Body. Sub-requests that
        return 429 are retried after the longest Retry-After in that batch.
        If the batch call itself returns 401, the token is refreshed once and
        the batch is sent again. Request IDs must be unique: IDs you do not set
        are generated so they never collide with yours. dependsOn is passed
        through, and every request a dependsOn names must be in the same chunk
        of 20. If any request is not a GET, the batch honours -WhatIf and -Confirm.
    .PARAMETER TenantId
        Customer tenant GUID or verified domain. Mandatory unless -PartnerTenant is used.
    .PARAMETER PartnerTenant
        Use your own partner tenant deliberately.
    .PARAMETER Request
        Hashtables or objects with method, url (relative, for example /users/{id})
        and optional id, body, headers and dependsOn.
    .PARAMETER ApiVersion
        v1.0 (default) or beta.
    .PARAMETER MaxRetries
        Retries for throttled sub-requests and for the batch call itself.
    .PARAMETER UserPrincipalName
        Technician whose refresh token is used.
    .EXAMPLE
        $requests = $userIds | ForEach-Object { @{ method = 'GET'; url = "/users/$_/authentication/methods" } }
        Invoke-MspGraphBatch -TenantId $tid -Request $requests
    #>
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium', DefaultParameterSetName = 'Tenant')]
    [OutputType('MspGdap.BatchResult')]
    param(
        [Parameter(Mandatory, ParameterSetName = 'Tenant')]
        [ValidateNotNullOrEmpty()]
        [string]$TenantId,

        [Parameter(Mandatory, ParameterSetName = 'Partner')]
        [switch]$PartnerTenant,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [object[]]$Request,

        [ValidateSet('v1.0', 'beta')]
        [string]$ApiVersion = 'v1.0',

        [ValidateRange(0, 10)]
        [int]$MaxRetries = 5,

        [string]$UserPrincipalName
    )

    try {
        $config = Get-MspConfigurationInternal -RequireComplete
        $tenant = Resolve-MspTargetTenant -TenantId $TenantId -PartnerTenant:$PartnerTenant -Configuration $config

        $normalised = [System.Collections.Generic.List[object]]::new()
        $readField = {
            param($Object, $Name)
            if ($Object -is [System.Collections.IDictionary]) { $Object[$Name] } elseif ($Object.PSObject.Properties[$Name]) { $Object.$Name } else { $null }
        }
        $usedIds = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($item in $Request) {
            $given = & $readField $item 'id'
            if ($given) {
                if (-not $usedIds.Add([string]$given)) {
                    throw (New-MspErrorRecord -Message "Batch request id '$given' is used more than once. Every id must be unique." -ErrorId 'MspGdap.GraphBatch.DuplicateId' -Category InvalidArgument)
                }
            }
        }
        $index = 0
        $generated = 0
        foreach ($item in $Request) {
            $index++
            $get = { param($name) & $readField $item $name }
            $method = ([string](& $get 'method')).ToUpperInvariant()
            if (-not $method) { $method = 'GET' }
            $url = [string](& $get 'url')
            if (-not $url) {
                throw (New-MspErrorRecord -Message "Batch request $index has no url." -ErrorId 'MspGdap.GraphBatch.InvalidRequest' -Category InvalidArgument)
            }
            if ($url -match '^[a-zA-Z][a-zA-Z0-9+.-]*://') {
                throw (New-MspErrorRecord -Message "Batch request $index must use a relative url such as /users." -ErrorId 'MspGdap.GraphBatch.InvalidRequest' -Category InvalidArgument)
            }
            $id = [string](& $get 'id')
            if (-not $id) {
                do { $generated++; $id = [string]$generated } while (-not $usedIds.Add($id))
            }
            $entry = [ordered]@{
                id     = $id
                method = $method
                url    = '/' + $url.TrimStart('/')
            }
            $dependsOn = & $get 'dependsOn'
            if ($dependsOn) { $entry.dependsOn = @($dependsOn | ForEach-Object { [string]$_ }) }
            $body = & $get 'body'
            $headers = & $get 'headers'
            if ($null -ne $body) {
                $entry.body = $body
                if (-not $headers) { $headers = @{ 'Content-Type' = 'application/json' } }
            }
            if ($headers) { $entry.headers = $headers }
            $normalised.Add([pscustomobject]$entry)
        }

        $writes = @($normalised | Where-Object { $_.method -ne 'GET' })
        if ($writes.Count -gt 0) {
            $label = if ($PartnerTenant) { "partner tenant $tenant" } else { "tenant $tenant" }
            if (-not $PSCmdlet.ShouldProcess($label, "Graph batch with $($normalised.Count) requests ($($writes.Count) writes)")) {
                return
            }
        }

        $batchUri = Resolve-MspGraphUri -Uri '$batch' -ApiVersion $ApiVersion
        $tokenParams = @{ Resource = 'Graph' }
        if ($UserPrincipalName) { $tokenParams.UserPrincipalName = $UserPrincipalName }
        if ($PartnerTenant) { $tokenParams.PartnerTenant = $true } else { $tokenParams.TenantId = $tenant }

        for ($offset = 0; $offset -lt $normalised.Count; $offset += 20) {
            $pending = @($normalised[$offset..([math]::Min($offset + 19, $normalised.Count - 1))])
            $chunkIds = @($pending | ForEach-Object { $_.id })
            foreach ($entry in $pending) {
                if ($entry.PSObject.Properties['dependsOn']) {
                    $outside = @($entry.dependsOn | Where-Object { $chunkIds -notcontains $_ })
                    if ($outside.Count -gt 0) {
                        throw (New-MspErrorRecord -Message "Batch request '$($entry.id)' depends on '$($outside -join ', ')', which is not in the same batch of 20. Keep dependent requests together." -ErrorId 'MspGdap.GraphBatch.InvalidDependency' -Category InvalidArgument)
                    }
                }
            }
            $attempt = 0
            while ($pending.Count -gt 0) {
                $refreshed = $false
                $response = $null
                while ($true) {
                    $headers = Get-MspAuthHeader @tokenParams -ForceRefresh:$refreshed
                    try {
                        $response = Invoke-MspRestWithRetry -Uri $batchUri -Method POST -Headers $headers -Body @{ requests = $pending } -MaxRetries $MaxRetries -ServiceName 'Graph'
                        break
                    }
                    catch {
                        if ($_.Exception.Data['StatusCode'] -eq 401 -and -not $refreshed) {
                            Write-Verbose 'Graph returned 401 for the batch. Refreshing the token once and retrying.'
                            $refreshed = $true
                            continue
                        }
                        throw
                    }
                    finally {
                        $headers = $null
                    }
                }
                $retry = [System.Collections.Generic.List[object]]::new()
                $wait = 0
                foreach ($sub in @($response.responses)) {
                    $original = $pending | Where-Object { $_.id -eq [string]$sub.id } | Select-Object -First 1
                    if ([int]$sub.status -eq 429 -and $attempt -lt $MaxRetries -and $original) {
                        $retry.Add($original)
                        $after = 0
                        if ($sub.PSObject.Properties['headers'] -and $sub.headers -and $sub.headers.PSObject.Properties['Retry-After']) {
                            $after = [int]$sub.headers.'Retry-After'
                        }
                        $wait = [math]::Max($wait, $after)
                        continue
                    }
                    [pscustomobject]@{
                        PSTypeName = 'MspGdap.BatchResult'
                        Id         = [string]$sub.id
                        Status     = [int]$sub.status
                        Success    = [int]$sub.status -ge 200 -and [int]$sub.status -lt 300
                        Method     = if ($original) { $original.method } else { $null }
                        Url        = if ($original) { $original.url } else { $null }
                        Body       = if ($sub.PSObject.Properties['body']) { $sub.body } else { $null }
                    }
                }
                $pending = $retry.ToArray()
                if ($pending.Count -gt 0) {
                    $attempt++
                    if ($wait -le 0) { $wait = [math]::Pow(2, $attempt) }
                    $wait = [math]::Min($wait, 60)
                    Write-Verbose "$($pending.Count) batch request(s) were throttled. Waiting $wait seconds."
                    Start-Sleep -Seconds $wait
                }
            }
        }
    }
    catch {
        $PSCmdlet.ThrowTerminatingError($_)
    }
}
