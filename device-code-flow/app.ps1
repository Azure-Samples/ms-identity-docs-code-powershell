# A PowerShell application that demonstrates how to use the
# Device Code flow to make an API call to Microsoft Graph.

# Import the MSAL module
Import-Module -Name Microsoft.Identity.Client

# Assemblies required by DeviceCodeHelper
[System.Collections.Generic.List[string]] $RequiredAssemblies = New-Object System.Collections.Generic.List[string]

$RequiredAssemblies.Add("System.Console.dll")
$RequiredAssemblies.Add("Microsoft.Identity.Client")

# DeviceCodeHelper, written in C#, will display a message on the console
# instructing the user how to authenticate via their device.
# AcquireTokenWithDeviceCode() will poll the server after firing the
# device code callback to look for a successful login with the provided code.

Add-Type -TypeDefinition @"
using System;
using System.Threading.Tasks;
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
"@ -ReferencedAssemblies $RequiredAssemblies -IgnoreWarnings -WarningAction SilentlyContinue

# 'Application (client) ID' of app registration in the Microsoft Entra admin center - this value is a GUID
$ClientId = "Enter_the_Application_Id_Here"

# 'Directory (tenant) ID' of app registration in the Microsoft Entra admin center - this value is a GUID
$TenantId = "Enter_the_Tenant_ID_Here"

# The Device Code flow requires a Public Client Application
# Build a PublicClientApplication with the $ClientId and $TenantId
$publicClient = [Microsoft.Identity.Client.PublicClientApplicationBuilder]::Create($ClientId).WithTenantId($TenantId).Build()

Write-Output "Checking cache for existing accounts."
# Look for cached access tokens.  This will attempt to utilize existing access tokens if possible.
# This sample will always result in a cache miss but demonstrates the proper usage pattern.
[Microsoft.Identity.Client.IAccount[]] $Accounts = $publicClient.GetAccountsAsync().GetAwaiter().GetResult()
if ($Accounts.Count) {
    Write-Output "Found an account, using the first one."
    [Microsoft.Identity.Client.IAccount] $Account = $publicClient.GetAccountsAsync().GetAwaiter().GetResult() | Select-Object -First 1
    $TokenResponse = $publicClient.AcquireTokenSilent([string[]]::"User.Read", $Account).ExecuteAsync().Result
} else {
    # No usable cached access token was found for this scope & account.  An interactive user flow will be required.
    Write-Output "No cached accounts found."
}

if ([string]::IsNullOrWhitespace($TokenResponse.AccessToken))   {
    Write-Output "Initiating a Device Code Flow."
    # Attempt to acquire an access token for the User.Read scope
    $TokenResponse = $publicClient.AcquireTokenWithDeviceCode([string[]]::"User.Read",[DeviceCodeHelper]::GetDeviceCodeResultCallback()).ExecuteAsync().Result
}

# Configure $GraphRequestParams with the AccessToken received from MSAL as the Bearer token for Graph
$GraphRequestParams = @{
    Method         = "GET"
    Uri            = "https://graph.microsoft.com/v1.0/me"
    Authentication = "Bearer"
    Token          = (ConvertTo-SecureString -String $TokenResponse.AccessToken -AsPlainText -Force)
}

# Send a request to the Graph API with the token to retrieve the values from /me
$GraphResponse = Invoke-RestMethod @GraphRequestParams

# Display the response to the console
Write-Output $GraphResponse
