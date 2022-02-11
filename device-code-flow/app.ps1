# 'Application (client) ID' of app registration in Azure portal - this value is a GUID
$ClientId = ""

# 'Directory (tenant) ID' of app registration in Azure portal - this value is a GUID
$TenantId = ""

# Configure $DeviceCodeRequestParams for the desired $TenantId and $ClientId
$DeviceCodeRequestParams = @{
    Method = "POST"
    Uri    = "https://login.microsoftonline.com/$TenantId/oauth2/devicecode"
    Body   = @{
        resource  = "https://graph.microsoft.com"
        client_id = $ClientId
    }
}

# Initiate a device code flow using the values specified in $DeviceCodeRequestParams
$DeviceCodeRequest = Invoke-RestMethod @DeviceCodeRequestParams

# Display the message instructing the user to enter a code at the specified URL on their device
Write-Host $DeviceCodeRequest.message

# Configure $TokenRequestParams for the desired $TenantId
$TokenRequestParams = @{
    Method = "POST"
    Uri    = "https://login.microsoftonline.com/$TenantId/oauth2/token"
    Body   = @{
        grant_type = "urn:ietf:params:oauth:grant-type:device_code"
        code       = $DeviceCodeRequest.device_code
        client_id  = $ClientId
    }
}

# Poll to check if the user has successfully authenticated.  If the authentication is still pending, suppress the error.
while ([string]::IsNullOrEmpty($TokenRequest.access_token)) {
    $TokenRequest = try {
        Invoke-RestMethod @TokenRequestParams -ErrorAction Stop
    }
    catch {
        $Message = $_.ErrorDetails.Message | ConvertFrom-Json
        if ($Message.error -ne "authorization_pending") {
            throw
        }
    }
    Start-Sleep -Seconds 1
}

# Configure $GraphRequestParams with the access_token
$GraphRequestParams = @{
    Method  = "GET"
    Uri     = "https://graph.microsoft.com/v1.0/me"
    Headers = @{
        "Authorization" = "Bearer " + $TokenRequest.access_token
    }
}

# Send a request to the Graph API with the token to retrieve the values from /me
$GraphRequest = Invoke-RestMethod @GraphRequestParams

# Display the response to the console
Write-Output $GraphRequest
