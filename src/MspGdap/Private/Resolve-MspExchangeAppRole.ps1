function Resolve-MspExchangeAppRole {
    <#
    .SYNOPSIS
        Resolves the directory roles to give an Exchange automation app, with guardrails.
    .DESCRIPTION
        An app-only service principal holds its directory roles permanently, outside GDAP, so the rules are
        stricter than for technician roles:
          - every role must be one Microsoft Learn lists for Exchange Online app-only authentication
            (Compliance Administrator, Exchange Administrator, Exchange Recipient Administrator, Global
            Administrator, Global Reader, Helpdesk Administrator, Security Administrator, Security Reader).
            Any other role would be assigned but do nothing for Exchange, so it is refused.
          - allowed without an extra switch: Exchange Administrator, Exchange Recipient Administrator,
            Global Reader and Security Reader
          - Compliance Administrator, Security Administrator and Helpdesk Administrator need -AllowPrivilegedRole
          - Global Administrator needs -AllowGlobalAdministrator
          - role template IDs that are not in the role catalogue are refused
    .PARAMETER Role
        Role names or template IDs.
    .PARAMETER AllowGlobalAdministrator
        Permit Global Administrator.
    .PARAMETER AllowPrivilegedRole
        Permit Compliance Administrator, Security Administrator and Helpdesk Administrator.
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory)][string[]]$Role,
        [switch]$AllowGlobalAdministrator,
        [switch]$AllowPrivilegedRole
    )
    $resolved = @(Resolve-MspGdapRole -Role $Role)
    $supported = @(Get-MspWellKnownRoleId -Name ExchangeAppOnlySupported)
    $globalAdministrator = Get-MspWellKnownRoleId -Name GlobalAdministrator
    $standard = @(
        (Get-MspWellKnownRoleId -Name ExchangeAdministrator)
        (Get-MspWellKnownRoleId -Name ExchangeRecipientAdministrator)
    )
    $readOnly = @($resolved | Where-Object { $_.DisplayName -in @('Global Reader', 'Security Reader') } | ForEach-Object { $_.RoleTemplateId })

    foreach ($entry in $resolved) {
        if ($supported -notcontains $entry.RoleTemplateId) {
            throw "$($entry.DisplayName) does not work for Exchange Online app-only access. Microsoft supports Exchange Administrator, Exchange Recipient Administrator, Compliance Administrator, Security Administrator, Security Reader, Helpdesk Administrator, Global Reader and Global Administrator."
        }
        if ($entry.RoleTemplateId -eq $globalAdministrator) {
            if (-not $AllowGlobalAdministrator) {
                throw 'Global Administrator is not allowed for the automation app. Microsoft says Exchange Administrator covers any Exchange Online PowerShell task. Use -AllowGlobalAdministrator only if you must.'
            }
            continue
        }
        if ($standard -notcontains $entry.RoleTemplateId -and $readOnly -notcontains $entry.RoleTemplateId -and -not $AllowPrivilegedRole) {
            throw "$($entry.DisplayName) gives the automation app standing access beyond Exchange. Pass -AllowPrivilegedRole if the job really needs it."
        }
    }
    foreach ($entry in $resolved) { $entry }
}
