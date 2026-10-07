function Get-MspHttpErrorInfo {
    <#
    .SYNOPSIS
        Extracts status code, Retry-After and the JSON error body from an
        Invoke-RestMethod error.
    .DESCRIPTION
        Works with the HttpResponseException thrown by PowerShell 7 and with
        MspGdap error records that already carry StatusCode in Exception.Data.
        Never returns request headers, so bearer tokens cannot leak through it.
    .PARAMETER ErrorRecord
        The caught error record.
    .EXAMPLE
        catch { $info = Get-MspHttpErrorInfo -ErrorRecord $_ }
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory)]
        [System.Management.Automation.ErrorRecord]$ErrorRecord
    )
    $exception = $ErrorRecord.Exception
    $statusCode = $null
    $retryAfter = $null

    if ($exception.Data.Contains('StatusCode')) {
        $statusCode = [int]$exception.Data['StatusCode']
    }

    $response = $null
    if ($exception.PSObject.Properties['Response']) {
        $response = $exception.Response
    }
    if ($response) {
        if ($null -eq $statusCode -and $response.PSObject.Properties['StatusCode']) {
            $statusCode = [int]$response.StatusCode
        }
        if ($response.PSObject.Properties['Headers'] -and $response.Headers -and $response.Headers.PSObject.Properties['RetryAfter']) {
            $retry = $response.Headers.RetryAfter
            if ($retry -and $retry.Delta) {
                $retryAfter = [int][math]::Ceiling($retry.Delta.TotalSeconds)
            }
            elseif ($retry -and $retry.Date) {
                $retryAfter = [int][math]::Max(0, [math]::Ceiling(([DateTimeOffset]$retry.Date - [DateTimeOffset]::UtcNow).TotalSeconds))
            }
        }
    }

    $body = $null
    if ($ErrorRecord.ErrorDetails -and $ErrorRecord.ErrorDetails.Message) {
        $body = $ErrorRecord.ErrorDetails.Message
    }
    $json = $null
    if ($body) {
        try { $json = $body | ConvertFrom-Json -ErrorAction Stop } catch { $json = $null }
    }

    [pscustomobject]@{
        StatusCode        = $statusCode
        RetryAfterSeconds = $retryAfter
        Body              = $body
        Json              = $json
    }
}
