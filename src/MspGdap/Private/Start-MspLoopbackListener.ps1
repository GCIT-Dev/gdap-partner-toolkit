function Start-MspLoopbackListener {
    <#
    .SYNOPSIS
        Runs a one-shot HTTP listener on http://localhost:<port>/ to receive
        the authorisation code redirect.
    .DESCRIPTION
        Starts System.Net.HttpListener on the loopback host only, opens the
        browser at the authorisation URL (or prints it with -NoBrowser), and
        waits for the redirect. On Windows the http://localhost prefix can also
        accept connections on other interfaces, so any request whose remote
        address is not a loopback address gets a 403 and is ignored without
        being parsed. Requests to other paths, or without a code or error, get
        a 404 and are ignored. A callback whose state does not match
        exactly is rejected and ignored, so a forged request cannot complete the
        flow. The listener is always stopped and closed when the function
        returns, times out or fails.

        The PKCE verifier stays with the caller. Only the PKCE challenge is in
        the authorisation URL.
    .PARAMETER Port
        Loopback port. Must match the redirect_uri in the authorisation URL.
    .PARAMETER ExpectedState
        The state value sent in the authorisation request.
    .PARAMETER AuthorizationUrl
        The full /authorize URL to open.
    .PARAMETER TimeoutSeconds
        How long to wait for the redirect.
    .PARAMETER NoBrowser
        Print the URL instead of opening a browser.
    .OUTPUTS
        PSCustomObject with Code and RedirectUri.
    .EXAMPLE
        Start-MspLoopbackListener -Port 53100 -ExpectedState $state -AuthorizationUrl $url -TimeoutSeconds 300
    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Private helper. Opens a short-lived loopback listener for one interactive sign-in and always closes it.')]
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory)]
        [ValidateRange(1, 65535)]
        [int]$Port,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$ExpectedState,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$AuthorizationUrl,

        [ValidateRange(1, 1800)]
        [int]$TimeoutSeconds = 300,

        [switch]$NoBrowser
    )

    $listener = [System.Net.HttpListener]::new()
    $listener.Prefixes.Add("http://localhost:$Port/")
    try {
        try {
            $listener.Start()
        }
        catch {
            throw (New-MspErrorRecord -Message "Could not listen on http://localhost:$Port/. Choose another port with Set-MspConfiguration -LoopbackPort, or use 0 for a random free port. Detail: $($_.Exception.Message)" -ErrorId 'MspGdap.Listener.StartFailed' -Category ResourceUnavailable -TargetObject $Port)
        }

        if ($NoBrowser) {
            Write-Information -MessageData "Open this URL in a browser on this computer and sign in with your partner admin account:`n$AuthorizationUrl" -InformationAction Continue
        }
        else {
            Open-MspBrowser -Url $AuthorizationUrl
        }
        Write-Verbose "Waiting up to $TimeoutSeconds seconds for the sign-in redirect on port $Port."

        $deadline = [datetime]::UtcNow.AddSeconds($TimeoutSeconds)
        while ($true) {
            $contextTask = $listener.GetContextAsync()
            while (-not $contextTask.IsCompleted) {
                if ([datetime]::UtcNow -ge $deadline) {
                    throw (New-MspErrorRecord -Message "No sign-in response was received within $TimeoutSeconds seconds. Run Register-MspPartnerToken again." -ErrorId 'MspGdap.Listener.Timeout' -Category OperationTimeout -TargetObject $Port)
                }
                $null = $contextTask.Wait(250)
            }
            $context = $contextTask.GetAwaiter().GetResult()
            $request = $context.Request
            $remote = $request.RemoteEndPoint
            if (-not $remote -or -not [System.Net.IPAddress]::IsLoopback($remote.Address)) {
                Write-Verbose 'Ignored a request to the sign-in listener that did not come from this computer.'
                Send-MspListenerResponse -Context $context -StatusCode 403 -Message 'Forbidden.'
                continue
            }
            $query = ConvertFrom-MspQueryString -Query $request.Url.Query

            $isCallback = $request.HttpMethod -eq 'GET' -and $request.Url.AbsolutePath -eq '/' -and ($query.ContainsKey('code') -or $query.ContainsKey('error'))
            if (-not $isCallback) {
                Send-MspListenerResponse -Context $context -StatusCode 404 -Message 'Not found.'
                continue
            }
            if (-not $query.ContainsKey('state') -or $query['state'] -cne $ExpectedState) {
                Write-Warning 'Ignored a sign-in callback with a missing or unexpected state value.'
                Send-MspListenerResponse -Context $context -StatusCode 400 -Message 'This sign-in response does not match the request in progress. Return to PowerShell.'
                continue
            }
            if ($query.ContainsKey('error')) {
                Send-MspListenerResponse -Context $context -StatusCode 200 -Message 'Sign-in did not complete. Return to PowerShell for details.'
                $errorCode = $query['error'] -replace '[^A-Za-z0-9_]', ''
                $description = if ($query.ContainsKey('error_description')) { ($query['error_description'] -split "`r?`n")[0] } else { '' }
                throw (New-MspErrorRecord -Message "Sign-in failed: $errorCode. $description" -ErrorId "MspGdap.Listener.$errorCode" -Category AuthenticationError -TargetObject $Port)
            }

            Send-MspListenerResponse -Context $context -StatusCode 200 -Message 'Sign-in complete. You can close this tab and return to PowerShell.'
            return [pscustomobject]@{
                Code        = $query['code']
                RedirectUri = "http://localhost:$Port"
            }
        }
    }
    finally {
        if ($listener.IsListening) { $listener.Stop() }
        $listener.Close()
    }
}
