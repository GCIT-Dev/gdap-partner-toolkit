function ConvertFrom-MspQueryString {
    <#
    .SYNOPSIS
        Parses a URL query string into a case-sensitive hashtable.
    .PARAMETER Query
        The query string, with or without the leading question mark.
    .EXAMPLE
        ConvertFrom-MspQueryString -Query '?code=abc&state=xyz'
    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseLiteralInitializerForHashtable', '', Justification = 'OAuth query parameter names are case-sensitive, so an ordinal comparer is intended.')]
    [CmdletBinding()]
    [OutputType([hashtable])]
    param(
        [AllowEmptyString()]
        [AllowNull()]
        [string]$Query
    )
    $result = [hashtable]::new([System.StringComparer]::Ordinal)
    if ([string]::IsNullOrEmpty($Query)) { return $result }
    foreach ($pair in $Query.TrimStart('?').Split('&')) {
        if ([string]::IsNullOrEmpty($pair)) { continue }
        $parts = $pair.Split('=', 2)
        $key = [System.Uri]::UnescapeDataString($parts[0].Replace('+', ' '))
        $value = if ($parts.Count -gt 1) { [System.Uri]::UnescapeDataString($parts[1].Replace('+', ' ')) } else { '' }
        $result[$key] = $value
    }
    $result
}
