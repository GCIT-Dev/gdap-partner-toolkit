function Invoke-MspTokenRequest {
    <#
    .SYNOPSIS
        Sends a form-encoded POST to the Microsoft identity platform v2.0 token
        endpoint of one tenant and maps failures to clear errors.
    .DESCRIPTION
        Retries transient failures (HTTP 429, 500, 502, 503, 504 and
        temporarily_unavailable) with backoff. Any other failure becomes a
        terminating error with ErrorId MspGdap.TokenRequest.AADSTS<code> and
        guidance for common codes: 50076 and 50079 (MFA), 65001 (consent),
        700082 (expired refresh token), 53003 (Conditional Access), 50020 (no
        GDAP access) and more.

        The request body is never written to verbose output or to errors,
        because it contains the refresh token or authorisation code.
    .PARAMETER TenantId
        Tenant GUID whose token endpoint is called. For customer tokens this is
        the customer tenant, not the partner tenant.
    .PARAMETER Body
        Form fields (client_id, grant_type, scope, credentials and the grant).
    .PARAMETER MaxRetries
        Retries for transient failures.
    .EXAMPLE
        Invoke-MspTokenRequest -TenantId $customerTenantId -Body $body
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory)]
        [ValidatePattern('^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$|^organizations$')]
        [string]$TenantId,

        [Parameter(Mandatory)]
        [System.Collections.IDictionary]$Body,

        [ValidateRange(0, 5)]
        [int]$MaxRetries = 2
    )

    $uri = '{0}/{1}/oauth2/v2.0/token' -f $script:MspConstants.Authority, $TenantId.ToLowerInvariant()
    Write-Verbose ("Token request: tenant {0}, grant_type {1}, scope '{2}'." -f $TenantId, $Body['grant_type'], $Body['scope'])

    for ($attempt = 0; $attempt -le $MaxRetries; $attempt++) {
        try {
            return Invoke-RestMethod -Uri $uri -Method Post -Body $Body -ContentType 'application/x-www-form-urlencoded' -ErrorAction Stop
        }
        catch {
            $info = Get-MspHttpErrorInfo -ErrorRecord $_
            $json = $info.Json
            $oauthError = if ($json -and $json.PSObject.Properties['error']) { [string]$json.error } else { $null }
            $description = if ($json -and $json.PSObject.Properties['error_description']) { [string]$json.error_description } else { $null }

            $transient = $info.StatusCode -in 429, 500, 502, 503, 504 -or $oauthError -eq 'temporarily_unavailable'
            if ($transient -and $attempt -lt $MaxRetries) {
                $delay = if ($info.RetryAfterSeconds) { [math]::Min($info.RetryAfterSeconds, 60) } else { [math]::Pow(2, $attempt + 1) }
                Write-Verbose "Token endpoint returned a transient error (HTTP $($info.StatusCode)). Retrying in $delay seconds."
                Start-Sleep -Seconds $delay
                continue
            }

            $code = $null
            if ($json -and $json.PSObject.Properties['error_codes'] -and $json.error_codes) {
                $code = [string]@($json.error_codes)[0]
            }
            elseif ($description -match 'AADSTS(\d+)') {
                $code = $Matches[1]
            }
            $guidance = Get-MspAadstsGuidance -Code $code -OAuthError $oauthError

            # Keep only the first line of Microsoft's description. It never contains token values.
            $firstLine = if ($description) { ($description -split "`r?`n")[0].Trim() } else { $_.Exception.Message }
            $traceId = if ($json -and $json.PSObject.Properties['trace_id']) { [string]$json.trace_id } else { $null }
            $correlationId = if ($json -and $json.PSObject.Properties['correlation_id']) { [string]$json.correlation_id } else { $null }

            $codeLabel = if ($code) { "AADSTS$code" } else { 'NoCode' }
            $message = "Token request for tenant $TenantId failed ($codeLabel): $($guidance.Guidance) Microsoft said: $firstLine"
            if ($correlationId) { $message += " (correlation ID $correlationId)" }

            $data = @{
                StatusCode    = $info.StatusCode
                AadstsCode    = $code
                OAuthError    = $oauthError
                Reason        = $guidance.Category
                TraceId       = $traceId
                CorrelationId = $correlationId
                TenantId      = $TenantId
            }
            throw (New-MspErrorRecord -Message $message -ErrorId "MspGdap.TokenRequest.$codeLabel" -Category AuthenticationError -TargetObject $TenantId -Data $data)
        }
    }
}
