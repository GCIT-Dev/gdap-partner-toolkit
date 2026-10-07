function Invoke-MspRestWithRetry {
    <#
    .SYNOPSIS
        Invoke-RestMethod with throttling and transient error handling.
    .DESCRIPTION
        Retries HTTP 429 for every method. HTTP 503 and 504 are retried only for
        idempotent methods (GET, PUT, DELETE), because a 503 or 504 on a POST or
        PATCH can mean the change was made, and repeating it could create a
        duplicate. Waits for the Retry-After header when present, otherwise uses
        exponential backoff with jitter, capped at MaxBackoffSeconds. Other failures become a terminating error that
        carries StatusCode, the service error code and request ID in
        Exception.Data. Request headers (which hold the bearer token) are
        never included in messages, and the original HttpResponseException is
        not attached because its request message carries the Authorization
        header.
    .PARAMETER Uri
        Absolute request URI.
    .PARAMETER Method
        HTTP method.
    .PARAMETER Headers
        Request headers including Authorization.
    .PARAMETER Body
        Request body (string or object). Objects are sent as JSON.
    .PARAMETER ContentType
        Content type for the body.
    .PARAMETER MaxRetries
        Maximum retries for retryable responses.
    .PARAMETER MaxBackoffSeconds
        Upper bound for a single wait.
    .PARAMETER ServiceName
        Label used in error messages, for example Graph.
    .EXAMPLE
        Invoke-MspRestWithRetry -Uri 'https://graph.microsoft.com/v1.0/organization' -Method GET -Headers $headers
    #>
    [CmdletBinding()]
    [OutputType([object])]
    param(
        [Parameter(Mandatory)]
        [uri]$Uri,

        [ValidateSet('GET', 'POST', 'PATCH', 'PUT', 'DELETE')]
        [string]$Method = 'GET',

        [System.Collections.IDictionary]$Headers,

        [object]$Body,

        [string]$ContentType = 'application/json',

        [ValidateRange(0, 10)]
        [int]$MaxRetries = 5,

        [ValidateRange(1, 300)]
        [int]$MaxBackoffSeconds = 60,

        [string]$ServiceName = 'HTTP'
    )

    $params = @{
        Uri         = $Uri
        Method      = $Method
        ErrorAction = 'Stop'
    }
    if ($Headers) { $params.Headers = $Headers }
    if ($PSBoundParameters.ContainsKey('Body') -and $null -ne $Body) {
        $params.Body = if ($Body -is [string]) { $Body } else { ConvertTo-Json -InputObject $Body -Depth 20 -Compress }
        $params.ContentType = $ContentType
    }

    $retryable = if ($Method -in 'GET', 'PUT', 'DELETE') { @(429, 503, 504) } else { @(429) }
    for ($attempt = 0; $attempt -le $MaxRetries; $attempt++) {
        try {
            return Invoke-RestMethod @params
        }
        catch {
            $info = Get-MspHttpErrorInfo -ErrorRecord $_
            if ($info.StatusCode -in $retryable -and $attempt -lt $MaxRetries) {
                $delay = if ($info.RetryAfterSeconds) {
                    [math]::Min([double]$info.RetryAfterSeconds, $MaxBackoffSeconds)
                }
                else {
                    [math]::Min([math]::Pow(2, $attempt) + (Get-Random -Minimum 0.0 -Maximum 1.0), $MaxBackoffSeconds)
                }
                Write-Verbose ("{0} returned HTTP {1} for {2} {3}. Waiting {4:N1} seconds (retry {5} of {6})." -f $ServiceName, $info.StatusCode, $Method, $Uri.AbsolutePath, $delay, ($attempt + 1), $MaxRetries)
                Start-Sleep -Milliseconds ([int]($delay * 1000))
                continue
            }

            $serviceCode = $null
            $serviceMessage = $null
            $requestId = $null
            if ($info.Json -and $info.Json.PSObject.Properties['error'] -and $info.Json.error -is [pscustomobject]) {
                $err = $info.Json.error
                if ($err.PSObject.Properties['code']) { $serviceCode = [string]$err.code }
                if ($err.PSObject.Properties['message']) { $serviceMessage = [string]$err.message }
                if ($err.PSObject.Properties['innerError'] -and $err.innerError -and $err.innerError.PSObject.Properties['request-id']) {
                    $requestId = [string]$err.innerError.'request-id'
                }
            }
            if (-not $serviceMessage) { $serviceMessage = $_.Exception.Message }

            $statusLabel = if ($info.StatusCode) { "HTTP $($info.StatusCode)" } else { 'no HTTP status' }
            $codeLabel = if ($serviceCode) { " $serviceCode" } else { '' }
            $message = "$ServiceName $Method $($Uri.AbsolutePath) failed ($statusLabel$codeLabel): $serviceMessage"
            if ($requestId) { $message += " (request ID $requestId)" }

            $category = switch ($info.StatusCode) {
                401 { [System.Management.Automation.ErrorCategory]::AuthenticationError }
                403 { [System.Management.Automation.ErrorCategory]::PermissionDenied }
                404 { [System.Management.Automation.ErrorCategory]::ObjectNotFound }
                409 { [System.Management.Automation.ErrorCategory]::ResourceExists }
                default { [System.Management.Automation.ErrorCategory]::InvalidOperation }
            }
            $idSuffix = if ($info.StatusCode) { $info.StatusCode } else { 'Error' }
            $data = @{
                StatusCode  = $info.StatusCode
                ServiceCode = $serviceCode
                RequestId   = $requestId
                Method      = $Method
                Path        = $Uri.AbsolutePath
                Json        = $info.Json
            }
            throw (New-MspErrorRecord -Message $message -ErrorId "MspGdap.$ServiceName.Http$idSuffix" -Category $category -TargetObject $Uri.AbsolutePath -Data $data)
        }
    }
}
