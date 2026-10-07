function Get-MspHttpStatusCode {
    <#
    .SYNOPSIS
        Extracts an HTTP status code from an ErrorRecord, whatever produced it.
    .DESCRIPTION
        Works with Invoke-RestMethod and Invoke-WebRequest errors (PowerShell 7 and Windows PowerShell 5.1)
        and with errors raised by Invoke-MspGraphRequest that carry a StatusCode property or Data entry.
        Returns $null when no status code can be found.
    #>
    [CmdletBinding()]
    [OutputType([int])]
    param([Parameter(Mandatory)][AllowNull()][object]$ErrorRecord)

    if ($null -eq $ErrorRecord) { return $null }
    $exception = if ($ErrorRecord -is [System.Management.Automation.ErrorRecord]) { $ErrorRecord.Exception } else { $ErrorRecord }

    $candidates = New-Object System.Collections.Generic.List[object]
    $current = $exception
    $depth = 0
    while ($null -ne $current -and $depth -lt 4) {
        foreach ($name in @('StatusCode', 'HttpStatusCode')) {
            $prop = $current.PSObject.Properties[$name]
            if ($prop) { $candidates.Add($prop.Value) }
        }
        $responseProp = $current.PSObject.Properties['Response']
        if ($responseProp -and $null -ne $responseProp.Value) {
            $statusProp = $responseProp.Value.PSObject.Properties['StatusCode']
            if ($statusProp) { $candidates.Add($statusProp.Value) }
        }
        if ($current -is [System.Exception] -and $current.Data -and $current.Data.Contains('StatusCode')) {
            $candidates.Add($current.Data['StatusCode'])
        }
        $current = if ($current -is [System.Exception]) { $current.InnerException } else { $null }
        $depth++
    }
    if ($ErrorRecord -is [System.Management.Automation.ErrorRecord] -and $ErrorRecord.TargetObject) {
        $targetProp = $ErrorRecord.TargetObject.PSObject.Properties['StatusCode']
        if ($targetProp) { $candidates.Add($targetProp.Value) }
    }

    foreach ($candidate in $candidates) {
        if ($null -eq $candidate) { continue }
        $number = 0
        if ($candidate -is [System.Enum]) { return [int]$candidate }
        if ([int]::TryParse([string]$candidate, [ref]$number) -and $number -ge 100 -and $number -le 599) { return $number }
    }

    $text = [string]$ErrorRecord
    if ($ErrorRecord -is [System.Management.Automation.ErrorRecord] -and $ErrorRecord.ErrorDetails) {
        $text = $text + ' ' + $ErrorRecord.ErrorDetails.Message
    }
    $patterns = @(
        'Response status code does not indicate success:\s*(\d{3})',
        '\((\d{3})\)\s',
        '\bstatus(?:\s*code)?\s*[:=]?\s*(\d{3})\b',
        '\bHTTP(?:/\d(?:\.\d)?)?\s+(\d{3})\b'
    )
    foreach ($pattern in $patterns) {
        $match = [regex]::Match($text, $pattern, 'IgnoreCase')
        if ($match.Success) { return [int]$match.Groups[1].Value }
    }
    return $null
}
