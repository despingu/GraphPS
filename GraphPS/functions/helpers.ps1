function Set-AccessToken {
    $logonUrl = "https://login.microsoftonline.com/$Script:tenantName/oauth2/v2.0/token"
    $body = @{
        "client_id" = $Script:clientID;
        "client_secret" = $Script:clientSecret;
        "scope" = $Script:resourceAppIdURI;
        "grant_type" = "client_credentials"
    }
    $tokenInfo = Invoke-RestMethod -Method Post -Uri $logonUrl -Body $body
    $Script:tokenRenewTime = (Get-Date).AddSeconds($tokenInfo.expires_in)
    $Script:token = "$($tokenInfo.token_type) $($tokenInfo.access_token)"
}

function Get-GraphUri ($Endpoint, $FilterExpression, $SelectExpression, $FormatExpression) {
    $uri = "$($Script:graphUri)/$($Script:graphVersion)/$Endpoint"
    $expressions = @()
    if(-not [string]::IsNullOrEmpty($FilterExpression)) {
        $expressions += "`$filter=$FilterExpression"
    }
    if(-not [string]::IsNullOrEmpty($SelectExpression)) {
        $expressions += "`$select=$SelectExpression"
    }
    if(-not [string]::IsNullOrEmpty($FormatExpression)) {
        $expressions += "`$format=$FormatExpression"
    }
    if($expressions.count -gt 0) {
        $expressionStr = $expressions -join "&"
        $uri += "?$expressionStr" 
    }

    return $uri
}

function Invoke-MSGraphQuery {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Endpoint,
        [Parameter(Mandatory=$true)]
        [ValidateSet('GET','POST','PATCH','PUT','DELETE')]
        [string]$Method, 
        [Parameter(Mandatory=$false)]
        [string]$FilterExpression,
        [Parameter(Mandatory=$false)]
        [string] $SelectExpression, 
        [Parameter(Mandatory=$false)]
        [string] $FormatExpression, 
        [Parameter(Mandatory=$false)]
        [string]$Body, 
        [Parameter(Mandatory=$false)]
        [hashtable]$customHeader
    )
    $Uri = Get-GraphUri -Endpoint $Endpoint -FilterExpression $FilterExpression -SelectExpression $SelectExpression -FormatExpression $FormatExpression

    if ((Get-Date) -ge $Script:tokenRenewTime) {
        $Script:token = Set-AccessToken
    }
    Write-Progress -Id 1 -Activity "Executing query: $Uri" -CurrentOperation "Invoking MS Graph API"
    $defaultHeader = @{
        'Content-Type'  = 'application\json'
        'Authorization' = $Script:token
    }

    $Header = @{} 
    if($null -ne $customHeader) {
        $Header = $defaultHeader, $customHeader | Merge-Hashtables {$_[-1]}
    }
    else {
        $Header = $defaultHeader
    }

    $QueryResults = @()
    if ($Method -eq "Get") {
        $i = 0
        do {
            $i++
            Write-Progress -Id 1 -Activity "Executing query: $Uri" -CurrentOperation "Fetching page $i"
            $Results = Invoke-RestMethod -Headers $Header -Uri $Uri -Method $Method -ContentType "application/json"
            if ($null -ne $Results.value) {
                $QueryResults += $Results.value
            }
            else {
                $QueryResults += $Results
            }
            $uri = $Results.'@odata.nextlink'
        } until ($null -eq $uri)
    }
    else  {
        $QueryResults = Invoke-RestMethod -Headers $Header -Uri $Uri -Method $Method -ContentType "application/json" -Body $Body
    }
    Write-Progress -Id 1 -Activity "Executing query: $Uri" -Completed
    Return $QueryResults
}

Function Merge-Hashtables([ScriptBlock]$Operator) {
    <#
        .SYNOPSIS
            A function to merge two or more hash tables.
            By default, all values from duplicated hash table entries will added to an array
            SYNTAX:
            HashTable[] <Hashtables> | Merge-Hashtables [-Operator <ScriptBlock>]
            From: https://stackoverflow.com/questions/8800375/merging-hashtables-in-powershell-how

        .EXAMPLE
            # Collecting all values in array
            $h1, $h2, $h3 | Merge-Hashtables

        .EXAMPLE
            # Overwrite using last value:
            $h1, $h2, $h3 | Merge-Hashtables {$_[-1]}
            
        .EXAMPLE
            # Overwrite using first value:
            $h1, $h2, $h3 | Merge-Hashtables {$_[0]}

        .EXAMPLE
            # Overwrite using max value:
            $h1, $h2, $h3 | Merge-Hashtables {($_ | Measure-Object -Maximum).Maximum}

        .EXAMPLE
            # Take the average values
            $h1, $h2, $h3 | Merge-Hashtables {($_ | Measure-Object -Average).Average}

        .EXAMPLE
            # Join the values together
            $h1, $h2, $h3 | Merge-Hashtables {$_ -Join ""}

        .EXAMPLE
            # Sort the values list
            $h1, $h2, $h3 | Merge-Hashtables {$_ | Sort-Object}
    #>
    $Output = @{}
    ForEach ($Hashtable in $Input) {
        If ($Hashtable -is [Hashtable]) {
            ForEach ($Key in $Hashtable.Keys) {$Output.$Key = If ($Output.ContainsKey($Key)) {@($Output.$Key) + $Hashtable.$Key} Else  {$Hashtable.$Key}}
        }
    }
    If ($Operator) {ForEach ($Key in @($Output.Keys)) {$_ = @($Output.$Key); $Output.$Key = Invoke-Command $Operator}}
    $Output
}
