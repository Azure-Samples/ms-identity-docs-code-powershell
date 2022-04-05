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

# Include the MsalCacheHelper used for token caching
if (-not ('Microsoft.Identity.Client.Extensions.Msal.MsalCacheHelper' -as [type])) {
    try {
        Add-Type -Path (Join-Path (Split-Path $PSCommandPath) 'Microsoft.Identity.Client.Extensions.Msal.2.19.6\lib\netcoreapp2.1\Microsoft.Identity.Client.Extensions.Msal.dll')
    }
    catch {
        Write-Error $_
        return
    }
}

# 'Application (client) ID' of app registration in Azure portal - this value is a GUID
$ClientId = ""

# 'Directory (tenant) ID' of app registration in Azure portal - this value is a GUID
$TenantId = ""

# Set to 0 by default.  Set this value to 1 to clear the cache.
$clearCache = 1

# The Device Code flow requires a Public Client Application
# Build a PublicClientApplication with the $ClientId and $TenantId
$publicClient = [Microsoft.Identity.Client.PublicClientApplicationBuilder]::Create($ClientId).WithTenantId($TenantId).Build()

# Setup caching.  In this sample a local file named "msal-cache.bin" is where cached tokens will be stored.
$cacheDir = Split-Path $PSCommandPath
$storageBuilder = New-Object Microsoft.Identity.Client.Extensions.Msal.StorageCreationPropertiesBuilder("msal-cache.bin", $cacheDir, $ClientId)
$cacheHelper = [Microsoft.Identity.Client.Extensions.Msal.MsalCacheHelper]::CreateAsync($storageBuilder.Build()).GetAwaiter().GetResult()
$cacheHelper.RegisterCache($publicClient.UserTokenCache)

if ($clearCache -eq 1)  {
    Write-Output "Clearing the cache."
    $cacheHelper.Clear()
}

Write-Output "Checking cache for existing accounts."
[Microsoft.Identity.Client.IAccount[]] $Accounts = $publicClient.GetAccountsAsync().GetAwaiter().GetResult()
if ($Accounts.Count) {
    Write-Output "Found an account, using the first one."
    [Microsoft.Identity.Client.IAccount] $Account = $publicClient.GetAccountsAsync().GetAwaiter().GetResult() | Select-Object -First 1
    $TokenResponse = $publicClient.AcquireTokenSilent([string[]]::"User.Read", $Account).ExecuteAsync().Result
} else {
    Write-Output "No cached acounts found."
}

if ([string]::IsNullOrWhitespace($TokenResponse.AccessToken))   {
    Write-Output "Initiating a Device Code Flow."
    # Acquire an AccessToken for the User.Read scope
    $TokenResponse = $publicClient.AcquireTokenWithDeviceCode([string[]]::"User.Read",[DeviceCodeHelper]::GetDeviceCodeResultCallback()).ExecuteAsync().Result
}

# Configure $GraphRequestParams with the AccessToken received from MSAL as the Bearer token for Graph
$GraphRequestParams = @{
    Method         = "GET"
    Uri            = "https://graph.microsoft.com/v1.0/me"
    Authentication = "Bearer"
    Token          = (ConvertTo-SecureString -String $TokenResponse.AccessToken -AsPlainText)
}

# Send a request to the Graph API with the token to retrieve the values from /me
$GraphResponse = Invoke-RestMethod @GraphRequestParams | ConvertTo-Json

# Display the response to the console
Write-Output $GraphResponse
