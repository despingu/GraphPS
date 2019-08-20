$appId = "yourappid"
$appSecret = "yoursecret"
$TenantName = "yourtenant.onmicrosoft.com"

Import-Module "PathToModule\GraphPS.psd1"
Get-Module GraphPS
Connect-GraphPS -TenantName $TenantName -AppID $appId -AppSecret $appSecret -Version beta
Get-GraphPSConnectionInfo
Get-GraphPSUsers -filterExpression "userType eq 'Guest'"