function Get-MspFreeLoopbackPort {
    <#
    .SYNOPSIS
        Returns a free TCP port on the loopback interface.
    .DESCRIPTION
        Entra ID ignores the port when matching a localhost redirect URI, so a
        single registered http://localhost redirect works with any port.
    .EXAMPLE
        $port = Get-MspFreeLoopbackPort
    #>
    [CmdletBinding()]
    [OutputType([int])]
    param()
    $listener = [System.Net.Sockets.TcpListener]::new([System.Net.IPAddress]::Loopback, 0)
    try {
        $listener.Start()
        [int]$listener.LocalEndpoint.Port
    }
    finally {
        $listener.Stop()
    }
}
