#Requires -Version 7.4
#Requires -Modules MspGdap

<#
.SYNOPSIS
    Asks technicians for IT Glue quick notes through a Power Automate approval, and appends the answers to IT Glue.
.DESCRIPTION
    Two modes, one for each function in the original solution:

    -RequestNotes picks random IT Glue organisations that have no quick notes and sends one to each
    technician through a Power Automate flow that starts an approval. Technicians come from a
    security group in your partner tenant (read through Microsoft Graph with MspGdap) or from
    -Technician. The flow's HTTP trigger URL contains a signature, so it is read from a
    SecretManagement vault like the IT Glue API key, never pasted into the script.

    -AppendNote adds the approver's comments to one organisation's quick notes. The note is HTML
    encoded before it is stored, so text typed into an approval can't inject markup into IT Glue.

    Both modes only report unless -Apply is set. Combine -Apply with -WhatIf to preview.

    To schedule it, host the script on Azure Functions 4.x with PowerShell 7.6 (Functions 1.x and
    its experimental PowerShell support are retired). Give the function app a managed identity
    with Key Vault Secrets User on a vault registered with SecretManagement, use a timer trigger for
    -RequestNotes, and protect the HTTP trigger for -AppendNote with Microsoft Entra
    authentication rather than a function key. See docs/08.

    This script works with your partner tenant and IT Glue only, so it takes no -TenantId or
    -AllCustomers. CustomerName in each row is the IT Glue organisation.
.PARAMETER RequestNotes
    Send quick note requests to technicians.
.PARAMETER TechnicianGroupId
    Object ID of a security group in your partner tenant whose enabled members receive requests.
.PARAMETER Technician
    Email addresses of technicians who receive requests, instead of -TechnicianGroupId.
.PARAMETER FlowUrlSecretName
    Name of the secret that holds the Power Automate HTTP trigger URL.
.PARAMETER AppendNote
    Append a note to one organisation's quick notes.
.PARAMETER OrganizationId
    IT Glue organisation ID to append the note to.
.PARAMETER Note
    Note text from the approval comments.
.PARAMETER Responder
    Name of the technician who wrote the note. Added after the note.
.PARAMETER ITGlueApiKeySecretName
    Name of the secret that holds the IT Glue API key.
.PARAMETER VaultName
    SecretManagement vault that holds the secrets. Defaults to your default vault.
.PARAMETER ITGlueBaseUri
    IT Glue API address for your data centre.
.PARAMETER Apply
    Send the requests or save the note. Without this switch the script only reports.
.PARAMETER OutputPath
    Optional path of a CSV file for the results.
.EXAMPLE
    ./itglue-quick-notes.ps1 -RequestNotes -TechnicianGroupId '00000000-0000-0000-0000-000000000000' -FlowUrlSecretName 'QuickNotesFlowUrl' -ITGlueApiKeySecretName 'ITGlueApiKey' -Apply

    Sends one quick note request per technician in the group. In an Azure Functions 4.x timer
    trigger, run.ps1 can call this line after param($Timer).
.EXAMPLE
    ./itglue-quick-notes.ps1 -AppendNote -OrganizationId 123456 -Note 'Owner prefers calls after 2 pm.' -Responder 'Jane Citizen' -ITGlueApiKeySecretName 'ITGlueApiKey' -Apply -WhatIf

    Shows the quick notes that would be saved. In an HTTP trigger, pass $Request.Body.id and
    $Request.Body.notes, then return the result with Push-OutputBinding.
.NOTES
    Replaces the original 2019 method: Azure Functions 1.x experimental PowerShell ($req and
    Out-File $res), the IT Glue API key pasted into both functions, and an anonymous Microsoft Flow
    HTTP trigger URL in the script.
    Required GDAP roles: none. The script works in your partner tenant, where the technician needs
    to read the group's members.
    Required partner app permissions: Microsoft Graph GroupMember.Read.All (delegated) for
    -TechnicianGroupId, which Directory.ReadWrite.All in partner-app.full.json covers. None for
    -Technician or -AppendNote.
.LINK
    https://gcit.com.au/knowledge-base/automatically-request-it-glue-quick-notes-from-your-team/
.LINK
    ../../docs/08-unattended-automation.md
.LINK
    https://learn.microsoft.com/en-us/azure/azure-functions/functions-reference-powershell
#>
[CmdletBinding(DefaultParameterSetName = 'RequestByGroup', SupportsShouldProcess, ConfirmImpact = 'Medium')]
[OutputType([pscustomobject])]
param(
    [Parameter(Mandatory, ParameterSetName = 'RequestByGroup')]
    [Parameter(Mandatory, ParameterSetName = 'RequestByList')]
    [switch]$RequestNotes,

    [Parameter(Mandatory, ParameterSetName = 'RequestByGroup')]
    [ValidatePattern('^[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}$')]
    [string]$TechnicianGroupId,

    [Parameter(Mandatory, ParameterSetName = 'RequestByList')]
    [ValidatePattern('^[^@\s]+@[^@\s]+\.[^@\s]+$')]
    [string[]]$Technician,

    [Parameter(Mandatory, ParameterSetName = 'RequestByGroup')]
    [Parameter(Mandatory, ParameterSetName = 'RequestByList')]
    [ValidateNotNullOrEmpty()]
    [string]$FlowUrlSecretName,

    [Parameter(Mandatory, ParameterSetName = 'Append')]
    [switch]$AppendNote,

    [Parameter(Mandatory, ParameterSetName = 'Append')]
    [ValidateRange('Positive')]
    [long]$OrganizationId,

    [Parameter(Mandatory, ParameterSetName = 'Append')]
    [ValidateLength(1, 10000)]
    [string]$Note,

    [Parameter(ParameterSetName = 'Append')]
    [ValidateLength(0, 200)]
    [string]$Responder,

    [Parameter(Mandatory)]
    [ValidateNotNullOrEmpty()]
    [string]$ITGlueApiKeySecretName,

    [string]$VaultName,

    [ValidateSet('https://api.itglue.com', 'https://api.eu.itglue.com', 'https://api.au.itglue.com')]
    [string]$ITGlueBaseUri = 'https://api.itglue.com',

    [switch]$Apply,

    [string]$OutputPath
)

$ErrorActionPreference = 'Stop'
$results = [System.Collections.Generic.List[object]]::new()

function ConvertTo-ResultRow {
    param([System.Collections.IDictionary]$Values = @{})
    $row = [ordered]@{}
    foreach ($column in 'CustomerTenantId', 'CustomerName', 'Status', 'ITGlueOrganisationId', 'Technician', 'Action', 'Detail') {
        $row[$column] = if ($Values.Contains($column)) { $Values[$column] } else { $null }
    }
    [pscustomobject]$row
}

function Get-SecretText {
    param([Parameter(Mandatory)][string]$Name, [string]$Vault)
    $secretParams = @{ Name = $Name }
    if ($Vault) { $secretParams.Vault = $Vault }
    $secret = Get-Secret @secretParams
    if ($secret -is [securestring]) { [System.Net.NetworkCredential]::new('', $secret).Password } else { [string]$secret }
}

function Get-ITGlueCollection {
    # Follows JSON:API links.next, but only within the IT Glue API host, so the key never goes elsewhere.
    param([Parameter(Mandatory)][string]$Path, [Parameter(Mandatory)][hashtable]$Headers, [Parameter(Mandatory)][string]$BaseUri)
    $next = "$BaseUri/$Path"
    while ($next) {
        if (-not $next.StartsWith("$BaseUri/", [System.StringComparison]::OrdinalIgnoreCase)) {
            throw "Refusing to follow an IT Glue link outside $BaseUri."
        }
        $page = Invoke-RestMethod -Method GET -Uri $next -Headers $Headers -ContentType 'application/vnd.api+json'
        foreach ($item in @($page.data)) { $item }
        $next = if ($page.links -and $page.links.next) { [string]$page.links.next } else { $null }
    }
}

$itGlueHeaders = @{ 'x-api-key' = (Get-SecretText -Name $ITGlueApiKeySecretName -Vault $VaultName); Accept = 'application/vnd.api+json' }

try {
    if ($AppendNote -and -not $RequestNotes) {
        $values = @{ ITGlueOrganisationId = [string]$OrganizationId; Status = 'Succeeded' }
        try {
            $organisation = (Invoke-RestMethod -Method GET -Uri "$ITGlueBaseUri/organizations/$OrganizationId" -Headers $itGlueHeaders -ContentType 'application/vnd.api+json').data
            $values.CustomerName = [string]$organisation.attributes.name
            $text = [System.Net.WebUtility]::HtmlEncode($Note.Trim())
            if ($Responder) { $text = "$text - $([System.Net.WebUtility]::HtmlEncode($Responder.Trim()))" }
            $existing = [string]$organisation.attributes.'quick-notes'
            $combined = if ($existing) { "$existing<br><br>$text" } else { $text }
            $values.Action = 'ReportOnly'
            $values.Detail = $text
            if ($Apply) {
                if ($PSCmdlet.ShouldProcess("IT Glue organisation $($values.CustomerName)", 'Append quick note')) {
                    $body = @{ data = @{ type = 'organizations'; attributes = @{ 'quick-notes' = $combined } } } | ConvertTo-Json -Depth 5
                    $null = Invoke-RestMethod -Method PATCH -Uri "$ITGlueBaseUri/organizations/$OrganizationId" -Headers $itGlueHeaders -ContentType 'application/vnd.api+json' -Body $body
                    $values.Action = 'Appended'
                }
                else {
                    $values.Action = 'WhatIf'
                }
            }
        }
        catch {
            $values.Status = 'Failed'
            $values.Detail = $_.Exception.Message
        }
        $row = ConvertTo-ResultRow -Values $values
        $results.Add($row)
        $row
    }
    else {
        if ($TechnicianGroupId) {
            $members = @(Invoke-MspGraphRequest -PartnerTenant -Method GET -Uri "v1.0/groups/$TechnicianGroupId/members/microsoft.graph.user?`$select=mail,userPrincipalName,accountEnabled")
            $Technician = @($members | Where-Object { $_.accountEnabled } | ForEach-Object { if ($_.mail) { [string]$_.mail } else { [string]$_.userPrincipalName } })
        }
        if ($Technician.Count -eq 0) { throw 'No technicians to send requests to.' }

        $withoutNotes = @(Get-ITGlueCollection -Path 'organizations?page[size]=1000' -Headers $itGlueHeaders -BaseUri $ITGlueBaseUri |
                Where-Object { -not $_.attributes.'quick-notes' })
        if ($withoutNotes.Count -eq 0) {
            $row = ConvertTo-ResultRow -Values @{ Status = 'Succeeded'; Action = 'None'; Detail = 'Every organisation already has quick notes.' }
            $results.Add($row)
            $row
        }
        else {
            $picked = @(Get-Random -InputObject $withoutNotes -Count ([math]::Min($Technician.Count, $withoutNotes.Count)))
            $flowUri = $null
            if ($Apply) {
                $flowUri = Get-SecretText -Name $FlowUrlSecretName -Vault $VaultName
                if ($flowUri -notmatch '^https://') { throw 'The flow URL secret must be an https URL.' }
            }
            for ($i = 0; $i -lt $picked.Count; $i++) {
                $organisation = $picked[$i]
                $values = @{ CustomerName = [string]$organisation.attributes.name; ITGlueOrganisationId = [string]$organisation.id; Technician = $Technician[$i]; Status = 'Succeeded'; Action = 'ReportOnly' }
                try {
                    if ($Apply) {
                        if ($PSCmdlet.ShouldProcess("technician $($Technician[$i])", "Request quick notes for $($values.CustomerName)")) {
                            $body = @{ name = $values.CustomerName; id = $values.ITGlueOrganisationId; approver = $Technician[$i] } | ConvertTo-Json
                            $null = Invoke-RestMethod -Method POST -Uri $flowUri -ContentType 'application/json' -Body $body
                            $values.Action = 'Requested'
                        }
                        else {
                            $values.Action = 'WhatIf'
                        }
                    }
                }
                catch {
                    $values.Status = 'Failed'
                    $values.Detail = $_.Exception.Message
                }
                $row = ConvertTo-ResultRow -Values $values
                $results.Add($row)
                $row
            }
        }
    }
}
finally {
    $itGlueHeaders = $null
    $flowUri = $null
}

if ($OutputPath -and $results.Count -gt 0) {
    $results | Export-Csv -LiteralPath $OutputPath -NoTypeInformation -Encoding utf8
    Write-Verbose "Saved $($results.Count) rows to $OutputPath"
}
