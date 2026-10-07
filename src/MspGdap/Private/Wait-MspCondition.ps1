function Wait-MspCondition {
    <#
    .SYNOPSIS
        Re-runs a readback until it returns a truthy value or the attempts run out. Used to confirm writes,
        which can take a few seconds to replicate in Microsoft Entra ID.
    .DESCRIPTION
        Attempts = 1 + ceiling(TimeoutSeconds / IntervalSeconds), with Start-Sleep between attempts.
        Counting attempts (not wall-clock time) keeps the behaviour predictable and testable.
    .OUTPUTS
        [pscustomobject] Satisfied (bool), Value (last value), Attempts, LastError (message or $null).
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory)][scriptblock]$Condition,
        [ValidateRange(0, 3600)][int]$TimeoutSeconds = 60,
        [ValidateRange(1, 300)][int]$IntervalSeconds = 5
    )
    $maxAttempts = 1 + [int][Math]::Ceiling($TimeoutSeconds / [double]$IntervalSeconds)
    $value = $null
    $lastError = $null
    for ($attempt = 1; $attempt -le $maxAttempts; $attempt++) {
        try {
            $value = & $Condition
            $lastError = $null
        }
        catch {
            $value = $null
            $lastError = $_.Exception.Message
        }
        if ($value) { return [pscustomobject]@{ Satisfied = $true; Value = $value; Attempts = $attempt; LastError = $null } }
        if ($attempt -lt $maxAttempts) { Start-Sleep -Seconds $IntervalSeconds }
    }
    [pscustomobject]@{ Satisfied = $false; Value = $value; Attempts = $maxAttempts; LastError = $lastError }
}
