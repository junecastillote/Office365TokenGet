Function New-MSGraphAPIToken {
    <#

    .SYNOPSIS
    Acquire authentication token for MS Graph API
    
    .DESCRIPTION
    If you have a registered app in Azure AD, this function can help you get the authentication token
    from the MS Graph API endpoint. Each token is valid for 60 minutes.
    
    .PARAMETER ClientID
    This is the registered ClientID in AzureAD
    
    .PARAMETER ClientSecret
    This is the key of the registered app in AzureAD
    
    .PARAMETER TenantID
    This is your Office 365 Tenant Domain
    
    .EXAMPLE
    $token = New-MSGraphAPIToken -ClientID <ClientID> -ClientSecret <ClientSecret> -TenantID <TenantID>

    The above example gets a new token using the ClientID, ClientSecret and TenantID combination
    
    .NOTES
    General notes
    #>
    
    param(
    [parameter(mandatory=$true)]
    [string]$ClientID,
    [parameter(mandatory=$true)]
    [string]$ClientSecret,
    [parameter(mandatory=$true)]
    [string]$TenantID
    )
    
    $body = @{grant_type="client_credentials";scope="https://graph.microsoft.com/.default";client_id=$ClientID;client_secret=$ClientSecret}
    $oauth = Invoke-RestMethod -Method Post -Uri https://login.microsoftonline.com/$TenantID/oauth2/v2.0/token -Body $body
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
    
    .PARAMETER ClientID
    This is the registered ClientID in AzureAD
    
    .PARAMETER ClientSecret
    This is the key of the registered app in AzureAD
    
    .PARAMETER TenantID
    This is your Office 365 TenantID
    
    .EXAMPLE
    $token = New-OutlookRestAPIToken -ClientID <ClientID> -ClientSecret <ClientSecret> -TenantID <TenantID>

    The above example gets a new token using the ClientID, ClientSecret and TenantID combination
    
    .NOTES
    General notes
    #>
    param(
    [parameter(mandatory=$true)]
    [string]$ClientID,
    [parameter(mandatory=$true)]
    [string]$ClientSecret,
    [parameter(mandatory=$true)]
    [string]$TenantID
    )
    
    $body = @{grant_type="client_credentials";scope="https://outlook.office.com/.default";client_id=$ClientID;client_secret=$ClientSecret}
    $oauth = Invoke-RestMethod -Method Post -Uri https://login.microsoftonline.com/$TenantID/oauth2/v2.0/token -Body $body
    $token = @{'Authorization'="$($oauth.token_type) $($oauth.access_token)"}
    Return $token
}

Function New-Office365ManagementAPIToken {
    <#

    .SYNOPSIS
    Acquire authentication token for Office 365 Management API
    
    .DESCRIPTION
    If you have a registered app in Azure AD, this function can help you get the authentication token
    from the Office 365 Management API endpoint. Each token is valid for 60 minutes.
    
    .PARAMETER ClientID
    This is the registered ClientID in AzureAD
    
    .PARAMETER ClientSecret
    This is the key of the registered app in AzureAD
    
    .PARAMETER TenantID
    This is your Office 365 Tenant Domain
    
    .EXAMPLE
    $token = New-ManagementAPIToken -ClientID <ClientID> -ClientSecret <ClientSecret> -TenantID <TenantID>

    The above example gets a new token using the ClientID, ClientSecret and TenantID combination
    
    .NOTES
    General notes
    #>
    
    param(
    [parameter(mandatory=$true)]
    [string]$ClientID,
    [parameter(mandatory=$true)]
    [string]$ClientSecret,
    [parameter(mandatory=$true)]
    [string]$TenantID
    )
    
    $body = @{grant_type="client_credentials";resource="https://manage.office.com";client_id=$ClientID;client_secret=$ClientSecret}
	$oauth = Invoke-RestMethod -Method Post -Uri "https://login.microsoftonline.com/$($tenantID)/oauth2/token?api-version=1.0" -Body $body
    $token = @{'Authorization'="$($oauth.token_type) $($oauth.access_token)"}
    Return $token
}