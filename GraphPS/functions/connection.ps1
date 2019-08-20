function Connect-GraphPS {
    param(
        [Parameter(Mandatory=$true)]
        [string]$TenantName,
        [Parameter(Mandatory=$true)]
        [string]$AppID,
        [Parameter(Mandatory=$true)]
        [string]$AppSecret,
        [Parameter(Mandatory=$false)]
        [ValidateSet('beta','v1.0')]
        [string]$Version = 'v1.0',
        [Parameter(Mandatory=$false)]
        [string]$ResourceAppIdURI = "https://graph.microsoft.com/.default"
    )
    
    $Script:tenantName = $TenantName
    $Script:clientID = $AppID
    $Script:clientSecret = $AppSecret
    $Script:resourceAppIdURI = $ResourceAppIdURI
    $Script:graphVersion = $Version
    Set-AccessToken -TenantName $TenantName -ClientID $AppID -ClientSecret $AppSecret -ResourceAppIdURI $ResourceAppIdURI
    $Script:connected = $true
}

function Disconnect-GraphPS {
    $Script:tenantName = [string]::Empty
    $Script:clientID = [string]::Empty
    $Script:clientSecret = [string]::Empty
    $Script:resourceAppIdURI = [string]::Empty
    $Script:token = [string]::Empty
    $Script:tokenRenewTime = [datetime]::MinValue
    $Script:connected = $false
    $Script:graphUri = "https://graph.microsoft.com"
    $Script:graphVersion = [string]::Empty
}

function Get-GraphPSConnectionInfo {
    Write-Host "Tenant name: $Script:tenantName"
    Write-Host "App ID: $Script:clientID"
    Write-Host "Resource AppID URI: $Script:resourceAppIdURI"
    Write-Host "Token renew time: $Script:tokenRenewTime"
    Write-Host "Connected: $Script:connected"
    Write-Host "Graph base uri: $Script:graphUri"
    Write-Host "Graph version: $Script:graphVersion"
}