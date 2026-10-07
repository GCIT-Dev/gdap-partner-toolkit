function Remove-MspPartnerAppConsent {
    <#
    .SYNOPSIS
        Removes the partner app's consent from customer tenants through Partner Center, and confirms it through Graph.
    .DESCRIPTION
        Uses Partner Center DELETE /v1/customers/{customerId}/applicationconsents/{applicationId} with the
        technician's MFA-backed token (ValidateMfa: true), then reads the customer through Microsoft Graph until
        the partner app has no tenant-wide delegated grants left (or no service principal at all).
        Use it when you offboard a customer, or to roll back Grant-MspPartnerAppConsent.
        The technician needs Application Administrator or Cloud Application Administrator in the customer
        through GDAP, and membership of AdminAgents. Your partner tenant is refused as a target.
        A Failed result is also written as a non-terminating error (MspGdap.Remove-MspPartnerAppConsent.Failed).
        It does not touch the separate automation app or any GDAP relationship.
    .PARAMETER TenantId
        One or more customer tenant IDs or verified domains. Alias: CustomerTenantId.
    .PARAMETER AppId
        Partner app ID. Must be the configured app, because the Partner Center token is minted for it.
    .PARAMETER ReadbackTimeoutSeconds
        How long to keep reading the customer back before reporting the removal. Default 60.
    .EXAMPLE
        Remove-MspPartnerAppConsent -TenantId 'fabrikam.onmicrosoft.com' -WhatIf
    .OUTPUTS
        One MspGdap.OperationResult per customer.
    #>
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [Alias('CustomerTenantId')][ValidateNotNullOrEmpty()][string[]]$TenantId,
        [ValidatePattern('^[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}$')][string]$AppId,
        [ValidateRange(0, 600)][int]$ReadbackTimeoutSeconds = 60
    )
    begin {
        try {
            $resolvedAppId = Resolve-MspPartnerAppId
            if ($AppId -and $AppId.ToLowerInvariant() -ne $resolvedAppId) {
                throw (New-MspErrorRecord -Message "AppId $AppId is not the configured partner app ($resolvedAppId). Partner Center tokens are always issued to the configured app." -ErrorId 'MspGdap.Consent.AppIdMismatch' -Category InvalidArgument -TargetObject $AppId)
            }
        }
        catch {
            $PSCmdlet.ThrowTerminatingError($_)
        }
        $stopAll = $false
    }
    process {
        foreach ($id in $TenantId) {
            $steps = New-Object System.Collections.Generic.List[object]
            $info = [ordered]@{ AppId = $resolvedAppId; IsMfaCompliant = $null }
            $tenant = $id
            $emit = { New-MspOperationResult -Cmdlet $PSCmdlet -Operation 'Remove-MspPartnerAppConsent' -TenantId $tenant -Target $resolvedAppId -Steps $steps.ToArray() -Property $info }
            try { $tenant = Resolve-MspCustomerTenant -TenantId $id }
            catch {
                $steps.Add((New-MspStepResult -Step 'Resolve tenant' -Status Failed -Detail $_.Exception.Message))
                & $emit; continue
            }
            if ($stopAll) {
                $steps.Add((New-MspStepResult -Step 'Partner Center MFA' -Status Failed -Detail 'Skipped: an earlier Partner Center call was rejected with MFA required.'))
                & $emit; continue
            }
            if (-not $PSCmdlet.ShouldProcess("customer $tenant", "Remove the consent of partner app $resolvedAppId (Partner Center DELETE)")) {
                $steps.Add((New-MspStepResult -Step 'Remove consent' -Status WhatIf -Detail 'Would DELETE the application consent.'))
                & $emit; continue
            }

            $response = Invoke-MspPartnerCenterRequest -Method DELETE -Path "customers/$tenant/applicationconsents/$resolvedAppId"
            if ($null -ne $response.IsMfaCompliant) { $info.IsMfaCompliant = $response.IsMfaCompliant }
            if ($response.MfaRequired) {
                $stopAll = $true
                $steps.Add((New-MspStepResult -Step 'Remove consent' -Status Failed -Detail $response.Error))
                & $emit; continue
            }
            $notFound = $response.StatusCode -eq 404
            if (-not $response.Success -and -not $notFound) {
                $steps.Add((New-MspStepResult -Step 'Remove consent' -Status Failed -Detail "Partner Center returned $($response.StatusCode). Nothing was confirmed removed. $($response.Error)"))
                & $emit; continue
            }
            $detail = if ($notFound) { 'Partner Center has no consent for this app in this customer (404). Checking Graph.' } else { "Partner Center accepted the DELETE ($($response.StatusCode))." }
            $steps.Add((New-MspStepResult -Step 'Remove consent' -Status Passed -Detail $detail))

            $check = Wait-MspCondition -TimeoutSeconds $ReadbackTimeoutSeconds -Condition {
                $now = Get-MspCustomerConsentState -TenantId $tenant -AppId $resolvedAppId -Grant @()
                if (-not $now.AppServicePrincipal -or @($now.OtherGrants).Count -eq 0) { $now }
            }
            if ($check.Satisfied) {
                $status = if ($notFound) { 'Passed' } else { 'Changed' }
                $text = if ($check.Value.AppServicePrincipal) { 'No tenant-wide delegated grants remain for the partner app.' } else { 'The partner app has no service principal in this customer.' }
                $steps.Add((New-MspStepResult -Step 'Confirm through Graph' -Status $status -Detail $text))
            }
            elseif ($check.LastError) {
                $steps.Add((New-MspStepResult -Step 'Confirm through Graph' -Status Unknown -Detail "Could not read the customer through Graph. $($check.LastError)"))
            }
            else {
                $hint = if ($notFound) { ' The consent was probably not created through Partner Center. A customer admin can delete the app under Enterprise apps.' } else { '' }
                $steps.Add((New-MspStepResult -Step 'Confirm through Graph' -Status Failed -Detail "The partner app still holds delegated grants in this customer.$hint"))
            }
            & $emit
        }
    }
}
