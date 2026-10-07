function Send-MspListenerResponse {
    <#
    .SYNOPSIS
        Writes a small static HTML page to a loopback listener request and closes it.
    .DESCRIPTION
        The page text is fixed by the caller and HTML encoded. Query values from
        the callback are never reflected into the page.
    .PARAMETER Context
        The HttpListenerContext.
    .PARAMETER StatusCode
        HTTP status code.
    .PARAMETER Message
        Plain text to show.
    .EXAMPLE
        Send-MspListenerResponse -Context $context -StatusCode 200 -Message 'Sign-in complete.'
    #>
    [CmdletBinding()]
    [OutputType([void])]
    param(
        [Parameter(Mandatory)]
        [System.Net.HttpListenerContext]$Context,

        [Parameter(Mandatory)]
        [int]$StatusCode,

        [Parameter(Mandatory)]
        [string]$Message
    )
    try {
        $encoded = [System.Net.WebUtility]::HtmlEncode($Message)
        $html = "<!DOCTYPE html><html lang=`"en`"><head><meta charset=`"utf-8`"><title>MspGdap sign-in</title></head><body style=`"font-family:system-ui,sans-serif;margin:3rem`"><p>$encoded</p></body></html>"
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($html)
        $response = $Context.Response
        $response.StatusCode = $StatusCode
        $response.ContentType = 'text/html; charset=utf-8'
        $response.Headers['Cache-Control'] = 'no-store'
        $response.Headers['Referrer-Policy'] = 'no-referrer'
        $response.ContentLength64 = $bytes.Length
        $response.OutputStream.Write($bytes, 0, $bytes.Length)
    }
    catch {
        Write-Verbose "Could not write the loopback response: $($_.Exception.Message)"
    }
    finally {
        try { $Context.Response.Close() } catch { Write-Verbose 'Loopback response already closed.' }
    }
}
