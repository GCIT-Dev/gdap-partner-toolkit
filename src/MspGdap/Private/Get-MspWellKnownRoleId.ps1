function Get-MspWellKnownRoleId {
    <#
    .SYNOPSIS
        Returns the built-in Microsoft Entra role template IDs that MspGdap treats specially.
    .DESCRIPTION
        These are public, built-in role template IDs (the same in every tenant). They are kept in one
        place so the guardrails in the GDAP and Exchange commands cannot drift apart.
        Exchange app-only roles are the roles Microsoft Learn lists for Exchange Online app-only
        authentication (app-only-auth-powershell-v2).
    .PARAMETER Name
        GlobalAdministrator, PrivilegedRoleAdministrator, ExchangeAdministrator,
        ExchangeRecipientAdministrator, or ExchangeAppOnlySupported (returns every supported role ID).
    .EXAMPLE
        Get-MspWellKnownRoleId -Name GlobalAdministrator
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory)]
        [ValidateSet('GlobalAdministrator', 'PrivilegedRoleAdministrator', 'ExchangeAdministrator', 'ExchangeRecipientAdministrator', 'ExchangeAppOnlySupported')]
        [string]$Name
    )
    $roles = [ordered]@{
        GlobalAdministrator            = '62e90394-69f5-4237-9190-012177145e10' # roleTemplateId, public-id
        PrivilegedRoleAdministrator    = 'e8611ab8-c189-46e8-94e1-60213ab1f814' # roleTemplateId, public-id
        ExchangeAdministrator          = '29232cdf-9323-42fd-ade2-1d097af3e4de' # roleTemplateId, public-id
        ExchangeRecipientAdministrator = '31392ffb-586c-42d1-9346-e59415a2cc4e' # roleTemplateId, public-id
        ComplianceAdministrator        = '17315797-102d-40b4-93e0-432062caca18' # roleTemplateId, public-id
        GlobalReader                   = 'f2ef992c-3afb-46b9-b7cf-a126ee74c451' # roleTemplateId, public-id
        HelpdeskAdministrator          = '729827e3-9c14-49f7-bb1b-9608f156bbb8' # roleTemplateId, public-id
        SecurityAdministrator          = '194ae4cb-b126-40b2-bd5b-6091b380977d' # roleTemplateId, public-id
        SecurityReader                 = '5d6b6bb7-de71-4623-b4af-96380a352509' # roleTemplateId, public-id
    }
    if ($Name -eq 'ExchangeAppOnlySupported') {
        foreach ($key in 'ComplianceAdministrator', 'ExchangeAdministrator', 'ExchangeRecipientAdministrator', 'GlobalAdministrator', 'GlobalReader', 'HelpdeskAdministrator', 'SecurityAdministrator', 'SecurityReader') {
            $roles[$key]
        }
        return
    }
    $roles[$Name]
}
