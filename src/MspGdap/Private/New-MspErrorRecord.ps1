function New-MspErrorRecord {
    <#
    .SYNOPSIS
        Builds an ErrorRecord with a stable MspGdap error ID and structured data.
    .DESCRIPTION
        Messages and data must never contain token or secret values. Callers
        pass only identifiers, status codes and Microsoft's error codes.
    .PARAMETER Message
        Human readable message.
    .PARAMETER ErrorId
        Stable identifier, for example MspGdap.TokenRequest.AADSTS700082.
    .PARAMETER Category
        PowerShell error category.
    .PARAMETER TargetObject
        Usually the tenant ID or URI involved.
    .PARAMETER Data
        Extra non-secret values copied to Exception.Data.
    .PARAMETER InnerException
        Optional underlying exception.
    .EXAMPLE
        throw (New-MspErrorRecord -Message 'No tenant' -ErrorId 'MspGdap.Tenant.Missing' -Category InvalidArgument)
    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Creates an in-memory object only.')]
    [CmdletBinding()]
    [OutputType([System.Management.Automation.ErrorRecord])]
    param(
        [Parameter(Mandatory)]
        [string]$Message,

        [Parameter(Mandatory)]
        [string]$ErrorId,

        [System.Management.Automation.ErrorCategory]$Category = [System.Management.Automation.ErrorCategory]::NotSpecified,

        [object]$TargetObject,

        [System.Collections.IDictionary]$Data,

        [System.Exception]$InnerException
    )
    $exception = if ($InnerException) {
        [System.InvalidOperationException]::new($Message, $InnerException)
    }
    else {
        [System.InvalidOperationException]::new($Message)
    }
    if ($Data) {
        foreach ($key in $Data.Keys) {
            $exception.Data[$key] = $Data[$key]
        }
    }
    [System.Management.Automation.ErrorRecord]::new($exception, $ErrorId, $Category, $TargetObject)
}
