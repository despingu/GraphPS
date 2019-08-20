# GraphPS
A PowerShell Module to interact with Microsoft Office 365 Graph API

# Quickstart
To use this module, you'll first need to register a new application in your Azure AD and get an app secret for it.
You can [read this article](https://docs.microsoft.com/en-us/azure/active-directory/develop/quickstart-register-app) for instruction on how to do that.
Keep in mind, that you also need to give your registred app permissions to access Graph.

Now download the repository and use this code snippet to test it:
```
$appId = "your-app-id"
$appSecret = "yourappsecret"
$TenantName = "yourtenant.onmicrosoft.com"

Import-Module "path-to-module\GraphPS.psd1"
Connect-GraphPS -TenantName $TenantName -AppID $appId -AppSecret $appSecret -Version beta
Get-GraphPSConnectionInfo
Get-GraphPSUsers -filterExpression "userType eq 'Guest'"
```