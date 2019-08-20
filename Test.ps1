$appId = "922ce248-2659-4958-8632-f6b65d9522a6"
$appSecret = "NnYCkx5g2TFLOdknKaIrpubXj8bK0jKfqzYlbOyQKFA="
$TenantName = "canda.onmicrosoft.com"

Import-Module "C:\Users\u067859\OneDrive - C&A\Projects\GraphPS-Module\GraphPS.psd1"
Get-Module GraphPS
Connect-GraphPS -TenantName $TenantName -AppID $appId -AppSecret $appSecret -Version beta
Get-GraphPSConnectionInfo
Get-GraphPSUsers -filterExpression "userType eq 'Guest'"