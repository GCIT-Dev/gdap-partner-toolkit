function New-MspStepResult {
    <#
    .SYNOPSIS
        Creates one step result for the structured output of MspGdap write and test commands.
    .DESCRIPTION
        Status values:
          Passed   the check succeeded or the target was already in the desired state
          Changed  a change was made AND confirmed by readback
          WhatIf   a change would have been made (-WhatIf, or declined at -Confirm)
          Skipped  the step did not apply
          Warning  informational, does not fail the operation
          Failed   the step failed, or a change could not be confirmed by readback
          Unknown  the state could not be determined (treat as a failure)
    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Creates an in-memory object only.')]
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory)][string]$Step,
        [Parameter(Mandatory)][ValidateSet('Passed', 'Changed', 'WhatIf', 'Skipped', 'Warning', 'Failed', 'Unknown')][string]$Status,
        [string]$Detail = '',
        [string]$Target,
        [object]$Data
    )
    [pscustomobject]@{
        PSTypeName   = 'MspGdap.StepResult'
        Step         = $Step
        Status       = $Status
        Detail       = $Detail
        Target       = $Target
        Data         = $Data
        TimestampUtc = [datetime]::UtcNow
    }
}
