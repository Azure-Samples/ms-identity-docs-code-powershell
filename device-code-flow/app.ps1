# A PowerShell application that demonstrates how to use the
# Device Code flow to make an API call to Microsoft Graph.

# Import the MSAL module
Import-Module -Name Microsoft.Identity.Client

# DeviceCodeHelper, written in C#, will display a message on the console
# instructing the user how to authenticate via their device.
# AcquireTokenWithDeviceCode() will poll the server after firing the
# device code callback to look for a successful login with the provided code.
Add-Type -TypeDefinition @"
using System;
using System.Threading.Tasks;
using System.Collections.Generic;
using Microsoft.Identity.Client;

public static class DeviceCodeHelper
{
    public static Func<DeviceCodeResult,Task> GetDeviceCodeResultCallback()
    {
        return deviceCodeResult =>
        {
            Console.WriteLine(deviceCodeResult.Message);
            return Task.FromResult(0);
        };
    }
}
"@ -ReferencedAssemblies Microsoft.Identity.Client

# 'Application (client) ID' of app registration in Azure portal - this value is a GUID
$ClientId = ""

# 'Directory (tenant) ID' of app registration in Azure portal - this value is a GUID
$TenantId = ""

# Scope permission for Graph
[string[]] $Scope = "User.Read"

# The Device Code flow requires a Public Client Application
# Build a PublicClientApplication with the $ClientId and $TenantId
$publicClient = [Microsoft.Identity.Client.PublicClientApplicationBuilder]::Create($ClientId).WithTenantId($TenantId)

# Acquire an AccessToken for the specified scope
$TokenRequest = $publicClient.Build().AcquireTokenWithDeviceCode($Scope, [DeviceCodeHelper]::GetDeviceCodeResultCallback()).ExecuteAsync().Result

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
