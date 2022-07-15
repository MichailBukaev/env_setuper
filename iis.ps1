Import-Module -Name IISAdministration

class RuleCondition {
    [string]$input
    [ValidateSet("Pattern", "IsFile", "IsDirectory")]
    [string]$matchType
    [string]$pattern = ""
    [bool]$ignoreCase = $True
    [bool]$negate = $False
}

function Initialize-IISSite {
    param (
        [PSCredential]$credential,
        [string]$siteName,
        [string]$iisSiteFolder,
        [string]$hostName = "",
        [int]$port = 80,
        [string]$appPoolName = $siteName
    )

    Write-Host Initialization of iis site $siteName started

    New-HostIfNotExist -hostName $hostName
    New-AppPoolIfNotExist -credential $credential -appPoolName $appPoolName
    New-SiteIfNotExist -siteName $siteName -hostName $hostName -port $port -iisSiteFolder $iisSiteFolder -appPoolName $appPoolName

    Write-Host Initialization of iis site $siteName comleted
}

function New-HostIfNotExist {
    param (
        [string]$hostName
    )
    if($hostName -eq ""){
        Return
    }

    $hostFile = "C:\Windows\System32\drivers\etc\hosts"
    
    if((Get-Content $hostFile) -contains "127.0.0.1 `t $hostName"){
        Write-Host Host name for $hostName is already exists. -ForegroundColor Yellow
    }
    else {
        Add-content -path $hostFile -value "127.0.0.1 `t $hostName"
        Write-host Host name "'$($hostName)'" is added to hosts file -ForegroundColor Green
    }  
}

function New-AppPoolIfNotExist {
    param (
        [PSCredential]$credential,
        [string]$appPoolName
    )
    $serverManager = Get-IISServerManager

    if($null -ne $serverManager.ApplicationPools[$appPoolName]){
        Write-Host Application pool $appPoolName alredy exists -ForegroundColor Yellow
    }
    else{
        try {
            Start-IISCommitDelay

            $appPool = $serverManager.ApplicationPools.Add($appPoolName)
            $appPool.ProcessModel.IdentityType = 'SpecificUser'
            $appPool.ProcessModel.UserName = $credential.UserName
            $appPool.ProcessModel.Password = ConvertFrom-SecureString $credential.Password -AsPlainText

            Write-Host Application pool $appPoolName is added -ForegroundColor Green

            Stop-IISCommitDelay -Commit $True
        }
        catch {
            Write-Host $_ -ForegroundColor Red
            Reset-IISServerManager
            Throw
        }
    }
}

function New-SiteIfNotExist {
    param (
        [string]$siteName,
        [string]$iisSiteFolder,
        [string]$hostName,
        [int]$port,
        [string]$appPoolName
    )

    if($null -ne (Get-IISSite -Name $siteName)){
        Write-Host  Site $siteName alredy exists -ForegroundColor Yellow
    }
    else{
        try {
            Start-IISCommitDelay

            $bindingInformation = "*:$($port):$hostName"
            $physicalPath = "D:\XCritical\web\$iisSiteFolder"

            Write-Host Site name is $siteName
            Write-Host BindingInformation is $bindingInformation
            Write-Host PhysicalPath is $physicalPath
    
            $site = New-IISSite `
            -Name $siteName `
            -BindingInformation  $bindingInformation `
            -PhysicalPath $physicalPath `
            -Passthru
            $site.Applications["/"].ApplicationPoolName = $appPoolName
    
            Write-Host IISSite $siteName is created -ForegroundColor Green
    
            Stop-IISCommitDelay -Commit $True
        }
        catch {
            Write-Host $_ -ForegroundColor Red
            Reset-IISServerManager
            Throw
        }
    }
}

function New-Rule {
    param (
        [string]$iisSiteName,
        [string]$ruleName,
        [string]$stopProcessing = $False,
        [string]$isEnabled = $True,
        [ValidateSet("ECMAScript", "Wildcard", "ExactMatch")]
        [string]$patternSyntax = "ECMAScript",
        [ValidateSet("Auto", "Always", "Never", "NotIfRuleMatched")]
        [string]$responseCacheDirective = "Auto"
    )
    Write-Host "Adding rule $ruleName for $iisSiteName iis site"

    if($null -ne (Get-Rule -iisSiteName $iisSiteName -ruleName $ruleName)){
        Write-Warning "Rule $ruleName alredy exists for $iisSiteName iis site"
        Return
    }

    try{
        Start-IISCommitDelay

        $rulesSections = Get-RulesSection -iisSiteName $iisSiteName
        $rulesCollection = $rulesSections.GetCollection()
        $rule = $rulesCollection.CreateElement("rule")
        $rule["name"] = $ruleName
        $rule["stopProcessing"] = $stopProcessing
        $rule["enabled"] = $isEnabled
        $rule["patternSyntax"] = $patternSyntax
        $rule["responseCacheDirective"] = $responseCacheDirective
        $rulesCollection.Add($rule)

        Stop-IISCommitDelay -Commit $True
        Write-Host "Rule $ruleName is added for $iisSiteName iis site" -ForegroundColor Green
    }    
    catch{
        Write-Error $_
        Reset-IISServerManager
        Throw   
    }
    
}

function Set-Match {
    param (
        [string]$iisSiteName,
        [string]$ruleName,
        [string]$mathPattern,
        [bool]$ignoreCase = $True,
        [bool]$negate = $False
    )
    Write-Host "Seting match for rule $ruleName of $iisSiteName iis site"
    $rule = Get-Rule -iisSiteName $iisSiteName -ruleName $ruleName
    if($null -ne $rule){
        try{
            Start-IISCommitDelay
            $matchElement = $rule.GetChildElement("match")
            $matchElement["url"] = $mathPattern
            $matchElement["ignoreCase"] = $ignoreCase
            $matchElement["negate"] = $negate
            Stop-IISCommitDelay -Commit $True
            Write-Host "Match for rule $ruleName of $iisSiteName iis site is seted"
        }
        catch{
            Write-Error $_
            Reset-IISServerManager
            Throw
        }
    }
    else {
        Write-Warning "Rule $ruleName is not exists for $iisSiteName iis site"
    }
}

function Add-Conditions {
    param (
        [string]$ruleName,
        [string]$iisSiteName,
        [ValidateSet("MatchAll", "MatchAny")]
        [string]$logicalGrouping = "MatchAll",
        [bool]$trackAllCaptures = $False,
        [RuleCondition[]]$conditions = @()
    )
    Write-Host "Add conditions for rule $ruleName of $iisSiteName iis site"
    $rule = Get-Rule -iisSiteName $iisSiteName -ruleName $ruleName
    if($null -ne $rule){
        try{
            Start-IISCommitDelay
            $conditionsElement = $rule.GetChildElement("conditions")
            $conditionsElement["logicalGrouping"] = $logicalGrouping
            $conditionsElement["trackAllCaptures"] = $trackAllCaptures
            $conditionsCollection = $conditionsElement.GetCollection()

            foreach ($condition in $conditions){
                Add-Condition -conditionsCollection $conditionsCollection -ruleCondition $condition
            }
    
            Stop-IISCommitDelay -Commit $True
            Write-Host "Conditions for rule $ruleName of $iisSiteName iis site is seted"
        }
        catch{
            Write-Error $_
            Reset-IISServerManager
            Throw
        }
    }
    else {
        Write-Warning "Rule $ruleName is not exists for $iisSiteName iis site"
    }
}

function Add-Condition {
    param (
        [Microsoft.Web.Administration.ConfigurationElementCollection]$conditionsCollection,
        [RuleCondition]$ruleCondition 
    )
    $condition = $conditionsCollection.CreateElement("add");
    $condition["input"] = $ruleCondition.input
    $condition["matchType"] = $ruleCondition.matchType
    $condition["pattern"] = $ruleCondition.matchType -eq "Pattern" ? $ruleCondition.pattern : ""
    $condition["ignoreCase"] = $ruleCondition.ignoreCase
    $condition["negate"] = $ruleCondition.negate

    $conditionsCollection.Add($condition)
}

function Add-ServerVariable {
    param (
        [string]$ruleName,
        [string]$iisSiteName,
        [string]$variableName,
        [string]$variableValue,
        [bool]$replace
    )
    Write-Host "Add server variables for rule $ruleName of $iisSiteName iis site"
    $rule = Get-Rule -iisSiteName $iisSiteName -ruleName $ruleName
    if($null -ne $rule){
        try{
            Start-IISCommitDelay
            $variablesElement = $rule.GetChildElement("serverVariables")
            $variablesCollection = $variablesElement.GetCollection()

            $variable = $variablesCollection.CreateElement("set");
            $variable["name"] = $variableName
            $variable["value"] = $variableValue
            $variable["replace"] = $replace

            $variablesCollection.Add($variable)
    
            Stop-IISCommitDelay -Commit $True
            Write-Host "Server variables for rule $ruleName of $iisSiteName iis site is seted"
        }
        catch{
            Write-Error $_
            Reset-IISServerManager
            Throw
        }
    }
    else {
        Write-Warning "Rule $ruleName is not exists for $iisSiteName iis site"
    }
}

function Set-Action {
    param (
        [string]$ruleName,
        [string]$iisSiteName,
        [string]$url,
        [ValidateSet("None", "Rewrite", "Redirect", "CustomResponse", "AbortRequest")]
        [string]$type = "Rewrite",
        [bool]$appendQueryString = $True,
        [bool]$logRewrittenUrl = $False,
        [int]$redirectType = 301,
        [int]$statusCode = 0,
        [int]$subStatusCode = 0,
        [string]$statusReason = "",
        [string]$statusDescription = ""
    )
    Write-Host "Add action for rule $ruleName of $iisSiteName iis site"
    $rule = Get-Rule -iisSiteName $iisSiteName -ruleName $ruleName
    if($null -ne $rule){
        try{
            Start-IISCommitDelay
            $action = $rule.GetChildElement("action")

            $action["type"] = $type
            $action["url"] = $url
            $action["appendQueryString"] = $appendQueryString
            $action["logRewrittenUrl"] = $logRewrittenUrl
            $action["redirectType"] = $redirectType
            if($statusCode -ne 0){
                 $action["statusCode"] = $statusCode
            }
            if($subStatusCode -ne 0){
                $action["subStatusCode"] = $subStatusCode
            }
            if($statusReason -ne ""){
                $action["statusReason"] = $statusReason
            }
            if($statusDescription -ne ""){
                $action["statusDescription"] = $statusDescription
            }
    
            Stop-IISCommitDelay -Commit $True
            Write-Host "Action for rule $ruleName of $iisSiteName iis site is seted"
        }
        catch{
            Write-Error $_
            Reset-IISServerManager
            Throw
        }
    }
    else {
        Write-Warning "Rule $ruleName is not exists for $iisSiteName iis site"
    }
    
}

function Get-Rule {
    param (
        $iisSiteName,
        $ruleName
    )
    $rulesSections = Get-RulesSection -iisSiteName $iisSiteName
    $rulesCollection = $rulesSections.GetCollection()
    foreach($rule in ($rulesCollection)){
        if($rule.Attributes['Name'].Value -eq $ruleName){
            Return $rule
        }
    }
    Return $null
}

function Get-RulesSection{
    param (
        [string]$iisSiteName
    )
    if($null -eq (Get-IISSite -Name $iisSiteName)){
        throw IIS Site $iisSiteName is not exsits
    }
    $serverManager = Get-IISServerManager
    $config = $serverManager.GetWebConfiguration($iisSiteName)
    $rulesSection = $config.GetSection("system.webServer/rewrite/rules")
    return $rulesSection
}