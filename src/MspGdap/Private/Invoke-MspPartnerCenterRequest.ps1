function Invoke-MspPartnerCenterRequest {
    <#
    .SYNOPSIS
        Sends one App+User request to the Partner Center REST API in the partner tenant.
    .DESCRIPTION
        - Token: Get-MspAuthHeader -PartnerTenant -Resource https://api.partnercenter.microsoft.com
        - Always sends ValidateMfa: true and reports isMfaCompliant (header or body) when Partner Center returns it.
          Since 1 April 2026 Partner Center rejects App+User calls without an MFA claim (401 MFA required).
        - Sends one ms-requestid (Partner Center's documented idempotency key) and one ms-correlationid per
          logical call, and reuses both on every retry, so a retried POST is not processed twice.
        - Retries 429, 500, 502, 503 and 504 with Retry-After or exponential backoff.
        - HTTP errors do not throw. The result carries StatusCode, Body, Error and MfaRequired.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][ValidateSet('GET', 'POST', 'PATCH', 'PUT', 'DELETE')][string]$Method,
        [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$Path,
        [object]$Body,
        [ValidateRange(0, 6)][int]$MaxRetry = 3
    )
    # Parameters: -Method (HTTP method), -Path (relative to https://api.partnercenter.microsoft.com/v1/),
    # -Body (object sent as JSON), -MaxRetry (retries for 429 and 5xx).
    $uri = 'https://api.partnercenter.microsoft.com/v1/' + $Path.TrimStart('/')
    $headers = @{}
    $auth = Get-MspAuthHeader -PartnerTenant -Resource 'https://api.partnercenter.microsoft.com'
    foreach ($key in @($auth.Keys)) { $headers[$key] = $auth[$key] }
    $headers['ValidateMfa'] = 'true'
    $headers['Accept'] = 'application/json'
    $headers['ms-requestid'] = [guid]::NewGuid().ToString()
    $headers['ms-correlationid'] = [guid]::NewGuid().ToString()

    $request = @{ Uri = $uri; Method = $Method; Headers = $headers; UseBasicParsing = $true; ErrorAction = 'Stop' }
    if ($PSBoundParameters.ContainsKey('Body') -and $null -ne $Body) {
        $request['Body'] = if ($Body -is [string]) { $Body } else { ConvertTo-Json -InputObject $Body -Depth 20 -Compress }
        $request['ContentType'] = 'application/json'
    }

    $attempt = 0
    while ($true) {
        $attempt++
        $status = $null; $content = $null; $responseHeaders = $null; $errorText = $null; $retryAfter = $null
        try {
            $response = Invoke-WebRequest @request
            $status = [int]$response.StatusCode
            $content = $response.Content
            $responseHeaders = $response.Headers
        }
        catch {
            $status = Get-MspHttpStatusCode -ErrorRecord $_
            if ($null -eq $status) { throw }
            $errorText = if ($_.ErrorDetails -and $_.ErrorDetails.Message) { $_.ErrorDetails.Message } else { $_.Exception.Message }
            $content = $errorText
            try {
                $resp = $_.Exception.Response
                if ($resp -and $resp.Headers) {
                    if ($resp.Headers.RetryAfter -and $resp.Headers.RetryAfter.Delta) { $retryAfter = [int]$resp.Headers.RetryAfter.Delta.TotalSeconds }
                    elseif ($resp.Headers['Retry-After']) { $retryAfter = [int]$resp.Headers['Retry-After'] }
                }
            }
            catch { $retryAfter = $null }
        }

        if ($status -in @(429, 500, 502, 503, 504) -and $attempt -le $MaxRetry) {
            $delay = if ($retryAfter -and $retryAfter -gt 0) { [Math]::Min($retryAfter, 60) } else { [Math]::Min([int][Math]::Pow(2, $attempt), 30) }
            Write-Verbose ("Partner Center returned {0}. Retrying in {1} s (attempt {2} of {3})." -f $status, $delay, $attempt, $MaxRetry)
            Start-Sleep -Seconds $delay
            continue
        }
        break
    }

    $parsed = $null
    if ($content) { try { $parsed = $content | ConvertFrom-Json -ErrorAction Stop } catch { $parsed = $null } }

    $mfaCompliant = $null
    if ($responseHeaders) {
        foreach ($key in @($responseHeaders.Keys)) {
            if ($key -ieq 'isMfaCompliant') {
                $mfaCompliant = [string](@($responseHeaders[$key])[0]) -ieq 'true'
            }
        }
    }
    if ($null -eq $mfaCompliant -and $parsed -and $parsed.PSObject.Properties['isMfaCompliant']) {
        $mfaCompliant = [bool]$parsed.isMfaCompliant
    }

    $mfaRequired = ($status -eq 401 -and [string]$content -match 'MFA')
    if ($mfaRequired) {
        $errorText = 'Partner Center rejected the request: MFA required. Register the technician token again from an MFA sign-in (Register-MspPartnerToken).'
    }

    $headers = $null
    $request = $null
    [pscustomobject]@{
        StatusCode     = $status
        Success        = ($status -ge 200 -and $status -lt 300)
        Body           = $parsed
        Error          = $errorText
        IsMfaCompliant = $mfaCompliant
        MfaRequired    = $mfaRequired
    }
}
