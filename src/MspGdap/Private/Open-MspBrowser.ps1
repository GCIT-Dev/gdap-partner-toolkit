function Open-MspBrowser {
    <#
    .SYNOPSIS
        Opens a URL in the default browser, or prints it if that fails.
    .PARAMETER Url
        The URL to open.
    .EXAMPLE
        Open-MspBrowser -Url $authorizationUrl
    #>
    [CmdletBinding()]
    [OutputType([void])]
    param(
        [Parameter(Mandatory)]
        [string]$Url
    )
    try {
        if ($IsWindows) {
            Start-Process -FilePath $Url -ErrorAction Stop
        }
        elseif ($IsMacOS) {
            Start-Process -FilePath 'open' -ArgumentList $Url -ErrorAction Stop
        }
        else {
            Start-Process -FilePath 'xdg-open' -ArgumentList $Url -ErrorAction Stop
        }
        Write-Information -MessageData 'A browser window has opened for sign-in. Complete sign-in and MFA with your partner admin account.' -InformationAction Continue
    }
    catch {
        Write-Information -MessageData "Could not open a browser. Open this URL to sign in:`n$Url" -InformationAction Continue
    }
}
