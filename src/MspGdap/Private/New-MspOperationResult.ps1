function New-MspOperationResult {
    <#
    .SYNOPSIS
        Wraps step results into one operation result per target.
    .DESCRIPTION
        Success is true only when there is at least one step and no step is Failed, Unknown or WhatIf.
        Outcome is Failed when any step is Failed or Unknown, WhatIf when any step is WhatIf, otherwise Succeeded.
        A failed step can never produce Success = true.

        With -Cmdlet (the calling command's $PSCmdlet), a Failed outcome is written to the caller's output
        first and then raised as a non-terminating error (ErrorId MspGdap.<Operation>.Failed, TargetObject =
        the result). Automation can therefore rely on $?, -ErrorAction Stop and try/catch as well as on the
        Success property. Nothing is returned in that case because the result was already written.
    .PARAMETER Operation
        Command name, for example Grant-MspPartnerAppConsent.
    .PARAMETER TenantId
        Tenant the operation acted on.
    .PARAMETER Target
        Object the operation acted on.
    .PARAMETER Steps
        Step results from New-MspStepResult.
    .PARAMETER Property
        Extra command-specific properties.
    .PARAMETER Cmdlet
        The calling command's $PSCmdlet. Enables the error record on failure.
    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Creates an in-memory object only.')]
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory)][string]$Operation,
        [string]$TenantId,
        [string]$Target,
        [AllowEmptyCollection()][object[]]$Steps = @(),
        [System.Collections.IDictionary]$Property,
        [System.Management.Automation.PSCmdlet]$Cmdlet
    )
    $stepList = @($Steps | Where-Object { $null -ne $_ })
    $bad = @($stepList | Where-Object { $_.Status -in @('Failed', 'Unknown') })
    $whatIf = @($stepList | Where-Object { $_.Status -eq 'WhatIf' })
    $outcome = if ($bad.Count -gt 0) { 'Failed' } elseif ($whatIf.Count -gt 0) { 'WhatIf' } else { 'Succeeded' }
    $success = ($stepList.Count -gt 0) -and ($outcome -eq 'Succeeded')

    $result = [ordered]@{
        PSTypeName = 'MspGdap.OperationResult'
        Operation  = $Operation
        TenantId   = $TenantId
        Target     = $Target
        Success    = $success
        Outcome    = $outcome
        Steps      = $stepList
    }
    if ($Property) {
        foreach ($key in $Property.Keys) { $result[$key] = $Property[$key] }
    }
    $object = [pscustomobject]$result

    if ($Cmdlet -and $outcome -eq 'Failed') {
        $Cmdlet.WriteObject($object)
        $first = $bad | Select-Object -First 1
        $message = "{0} failed for {1}: step '{2}' is {3}. {4}" -f $Operation, $(if ($TenantId) { $TenantId } else { $Target }), $first.Step, $first.Status, $first.Detail
        $exception = [System.InvalidOperationException]::new($message.Trim())
        $exception.Data['Operation'] = $Operation
        $exception.Data['TenantId'] = $TenantId
        $record = [System.Management.Automation.ErrorRecord]::new($exception, "MspGdap.$Operation.Failed", [System.Management.Automation.ErrorCategory]::InvalidResult, $object)
        $Cmdlet.WriteError($record)
        return
    }
    $object
}
