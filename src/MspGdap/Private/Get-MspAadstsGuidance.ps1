function Get-MspAadstsGuidance {
    <#
    .SYNOPSIS
        Maps a Microsoft Entra AADSTS error code to plain guidance for an MSP technician.
    .PARAMETER Code
        Numeric AADSTS code (for example 700082).
    .PARAMETER OAuthError
        The OAuth error value (for example invalid_grant) used when the code is unknown.
    .EXAMPLE
        Get-MspAadstsGuidance -Code 65001
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [string]$Code,
        [string]$OAuthError
    )

    $map = @{
        '50076'   = @('MfaRequired', 'Multifactor authentication is required. Register the technician token again with Register-MspPartnerToken and complete MFA.')
        '50079'   = @('MfaRequired', 'The technician account must register for multifactor authentication before it can be used. Complete MFA registration, then run Register-MspPartnerToken again.')
        '50158'   = @('MfaRequired', 'An external security challenge (for example a third-party MFA provider) was not satisfied. Register the technician token again interactively.')
        '65001'   = @('ConsentRequired', 'The partner app has not been consented in this customer tenant, or the requested permission is not consented. Pre-consent the app with Grant-MspPartnerAppConsent, then retry.')
        '65004'   = @('ConsentRequired', 'Consent was declined during sign-in.')
        '700016'  = @('ConsentRequired', 'The partner app was not found in this tenant. It has no service principal there yet. Pre-consent the app with Grant-MspPartnerAppConsent.')
        '500011'  = @('ResourceNotFound', 'The resource (API) has no service principal in this tenant. The customer may not be licensed for that service.')
        '70011'   = @('InvalidScope', 'The requested scope is invalid. Request one resource at a time, for example https://graph.microsoft.com/.default.')
        '700082'  = @('RefreshTokenExpired', 'The stored refresh token expired after 90 days without use. Run Register-MspPartnerToken again.')
        '70008'   = @('RefreshTokenExpired', 'The stored refresh token or authorisation code has expired. Run Register-MspPartnerToken again.')
        '70043'   = @('RefreshTokenExpired', 'The refresh token expired because of a Conditional Access sign-in frequency policy. Run Register-MspPartnerToken again.')
        '50173'   = @('RefreshTokenRevoked', 'The refresh token was revoked (password reset or session revocation). Run Register-MspPartnerToken again.')
        '50133'   = @('RefreshTokenRevoked', 'The session is invalid because of a password change or expiry. Run Register-MspPartnerToken again.')
        '700084'  = @('RefreshTokenExpired', 'The refresh token was issued to a single-page application and has expired. Register the redirect URI on the Web platform, not SPA.')
        '53000'   = @('ConditionalAccess', 'A Conditional Access policy requires a compliant device.')
        '53001'   = @('ConditionalAccess', 'A Conditional Access policy requires a domain-joined device.')
        '53003'   = @('ConditionalAccess', 'Access was blocked by a Conditional Access policy. Check the sign-in logs in the partner or customer tenant.')
        '530003'  = @('ConditionalAccess', 'A Conditional Access policy requires a managed device.')
        '50020'   = @('NoGdapAccess', 'The technician has no access to this customer tenant. Check that an active GDAP relationship exists and that one of the technician''s security groups has an active access assignment.')
        '50177'   = @('MfaRequired', 'An external (federated or third-party) MFA challenge is not supported for this account. Use a cloud-only partner admin account with Microsoft Entra MFA, then run Register-MspPartnerToken again.')
        '90002'   = @('TenantNotFound', 'The tenant was not found. Check the tenant ID or domain.')
        '90072'   = @('NoGdapAccess', 'The technician account does not exist in this tenant and has no GDAP access to it.')
        '7000215' = @('InvalidClientCredential', 'The partner app client secret is invalid or expired. Update it with Set-MspConfiguration -ClientSecret.')
        '7000222' = @('InvalidClientCredential', 'The partner app client secret has expired. Prefer a certificate credential.')
        '700027'  = @('InvalidClientCredential', 'The client assertion signature failed. The certificate is not registered on the partner app, has expired, or the private key does not match.')
        '700024'  = @('InvalidClientCredential', 'The client assertion is outside its valid time range. Check the local clock.')
        '700025'  = @('InvalidClientCredential', 'The app is registered as a public client. Register http://localhost on the Web platform so the confidential client flow works.')
        '50011'   = @('RedirectUriMismatch', 'The redirect URI does not match the partner app registration. Add http://localhost as a Web redirect URI.')
        '9002313' = @('InvalidRequest', 'The token request was malformed.')
    }

    $category = 'TokenRequestFailed'
    $guidance = 'The token request failed.'
    if ($Code -and $map.ContainsKey($Code)) {
        $category = $map[$Code][0]
        $guidance = $map[$Code][1]
    }
    elseif ($OAuthError -eq 'invalid_grant') {
        $category = 'InvalidGrant'
        $guidance = 'The refresh token or authorisation code was rejected. If this persists, run Register-MspPartnerToken again.'
    }
    elseif ($OAuthError -eq 'interaction_required') {
        $category = 'InteractionRequired'
        $guidance = 'Interactive sign-in is required. Run Register-MspPartnerToken again.'
    }
    elseif ($OAuthError -eq 'invalid_client') {
        $category = 'InvalidClientCredential'
        $guidance = 'The partner app credential was rejected. Check the certificate or client secret in Set-MspConfiguration.'
    }

    [pscustomobject]@{
        Code     = $Code
        Category = $category
        Guidance = $guidance
    }
}
