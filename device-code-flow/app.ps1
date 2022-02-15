# A PowerShell application that demonstrates how to use the
# Device Code flow to make an API call to Microsoft Graph.

# 'Application (client) ID' of app registration in Azure portal - this value is a GUID
$ClientId = ""

# 'Directory (tenant) ID' of app registration in Azure portal - this value is a GUID
$TenantId = ""

# Request the token from MSAL using the configured ClientId and TenantId for the "User.Read" scope
# The -Device flag is used to initiate a Device Code flow
# $TokenRequest will contain the response of this request
$TokenRequest = Get-MsalToken -ClientId $ClientId -Authority "https://login.microsoftonline.com/$TenantId" -Scopes "User.Read" -Device

# Configure $GraphRequestParams with the AccessToken received from MSAL as the Bearer token for Graph
$GraphRequestParams = @{
    Method  = "GET"
    Uri     = "https://graph.microsoft.com/v1.0/me"
    Headers = @{
        "Authorization" = "Bearer " + $TokenRequest.AccessToken
    }
}

# Send a request to the Graph API with the token to retrieve the values from /me
$GraphRequest = Invoke-RestMethod @GraphRequestParams

# Display the response to the console
Write-Output $GraphRequest
