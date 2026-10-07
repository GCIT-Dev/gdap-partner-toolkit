function Test-MspConsentMissingError {
    <#
    .SYNOPSIS
        True when an error means the partner app has no consent in the customer yet (AADSTS65001).
    .DESCRIPTION
        Microsoft Entra ID returns AADSTS65001 ("The user or administrator has not consented to use the
        application") when the partner app requests a token in a customer before it has been pre-consented.
        The consent functions read the customer with that same app, so this error is the expected state of a
        customer that has never been consented, not a failure to read it.
    .PARAMETER ErrorRecord
        The error from the failed call.
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param([Parameter(Mandatory)][System.Management.Automation.ErrorRecord]$ErrorRecord)
    $text = @($ErrorRecord.Exception.Message, [string]$ErrorRecord.ErrorDetails, [string]$ErrorRecord.FullyQualifiedErrorId) -join ' '
    $inner = $ErrorRecord.Exception.InnerException
    while ($inner) { $text += ' ' + $inner.Message; $inner = $inner.InnerException }
    return ($text -match 'AADSTS65001\b')
}
