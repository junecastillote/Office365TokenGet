Function New-MSGraphAPIToken {
    <#

    .SYNOPSIS
    Acquire authentication token for MS Graph API
    
    .DESCRIPTION
    If you have a registered app in Azure AD, this function can help you get the authentication token
    from the MS Graph API endpoint. Each token is valid for 60 minutes.
    
    .PARAMETER appID
    This is the registered appID in AzureAD
    
    .PARAMETER appKey
    This is the key of the registered app in AzureAD
    
    .PARAMETER domain
    This is your Office 365 Tenant Domain
    
    .EXAMPLE
    $graphToken = New-MSGraphAPIToken -appID <appID> -appKey <appKey> -domain <tenant domain>

    The above example gets a new token using the appID, appKey and tenant domain combination
    
    .NOTES
    General notes
    #>
    
    param(
    [parameter(mandatory=$true)]
    [string]$appID,
    [parameter(mandatory=$true)]
    [string]$appKey,
    [parameter(mandatory=$true)]
    [string]$domain
    )
    
    $body = @{grant_type="client_credentials";scope="https://graph.microsoft.com/.default";client_id=$appID;client_secret=$appKey}
    $oauth = Invoke-RestMethod -Method Post -Uri https://login.microsoftonline.com/$domain/oauth2/v2.0/token -Body $body
    $token = @{'Authorization'="$($oauth.token_type) $($oauth.access_token)"}    
    Return $token
}

Function New-OutlookRestAPIToken {
    <#
    .SYNOPSIS
    Acquire authentication token for Outlook REST API
    
    .DESCRIPTION
    If you have a registered app in Azure AD, this function can help you get the authentication token
    from the Outlook REST API endpoint. Each token is valid for 60 minutes.
    
    .PARAMETER appID
    This is the registered appID in AzureAD
    
    .PARAMETER appKey
    This is the key of the registered app in AzureAD
    
    .PARAMETER domain
    This is your Office 365 Tenant Domain
    
    .EXAMPLE
    $graphToken = New-OutlookRestAPIToken -appID <appID> -appKey <appKey> -domain <tenant domain>

    The above example gets a new token using the appID, appKey and tenant domain combination
    
    .NOTES
    General notes
    #>
    param(
    [parameter(mandatory=$true)]
    [string]$appID,
    [parameter(mandatory=$true)]
    [string]$appKey,
    [parameter(mandatory=$true)]
    [string]$domain
    )
    
    $body = @{grant_type="client_credentials";scope="https://outlook.office.com/.default";client_id=$appID;client_secret=$appKey}
    $oauth = Invoke-RestMethod -Method Post -Uri https://login.microsoftonline.com/$domain/oauth2/v2.0/token -Body $body
    $token = @{'Authorization'="$($oauth.token_type) $($oauth.access_token)"}    
    Return $token
}