Import-Module -Name IISAdministration

function Publish-Service () {
    param (
        [string]$package, 
        [string]$serviceName, 
        [string]$serviceFolder
    )

    Write-Host "Publish $serviceName"

    Stop-ServiseIfExists -serviceName $serviceName

    Move-Files -destination D:\XCritical\services\$serviceFolder -source D:\XCritical\packages\$package

    Start-ServiseIfExists -serviceName $serviceName
}

function Publish-IISSite () {
    param (
        [string]$package, 
        [string]$iisSiteName, 
        [string]$iisSiteFolder,
        [string]$exclude = "",
        [string]$appPoolName = ""
    )
    
    Write-Host "Publish $iisSiteName"

    if($appPoolName -eq ""){
        $appPoolName = $iisSiteName;
    }

    Stop-IISSiteIfExists -iisSiteName $iisSiteName
    Stop-IISAppPollIfExists -appPoolName $appPoolName

    Move-Files -destination D:\XCritical\web\$iisSiteFolder -source D:\XCritical\packages\$package -exclude $exclude
    
    Start-IISSiteIfExists -iisSiteName $iisSiteName
    Start-IISAppPollIfExists -appPoolName $appPoolName
}

function Publish-Db () {
    param (
        [string]$package,
        [string]$dbName,
        [string]$sourceFolder,
        [string]$dacpacFileName,
        [string]$build,
        [string]$optionalSqlcmdVaeriables = ""
    )
    Write-Host "Update $dbName"

    Move-Files -destination D:\XCritical\db\$sourceFolder -source D:\XCritical\packages\$package

    If(!(test-path D:\XCritical\deployment_reports\$dbName)){
        New-Item -ItemType Directory -Force -Path D:\XCritical\deployment_reports\$dbName
    }
    
    $cmd = "SqlPackage.exe "+ 
    "/a:Publish "+
    "/drp:D:\XCritical\deployment_reports\$dbName\$($dbName)_$($build)_drp.xml "+
    "/dsp:D:\XCritical\deployment_reports\$dbName\$($dbName)_$($build)_dsp.sql "+
    "/of:True "+
    "/sf:D:\XCritical\db\$sourceFolder\$dacpacFileName.dacpac "+
    "/tsn:. "+
    "/tdn:$dbName "+
    "/p:ScriptNewConstraintValidation=False "+
    "/p:GenerateSmartDefaults=True "+
    "/p:BlockOnPossibleDataLoss=False "+
    "/p:IgnoreColumnOrder=False "+
    "/v:DbType=mock  "+
    "/v:DbVer=$build"

    if($optionalSqlcmdVaeriables -ne "") {
        $cmd += " $optionalSqlcmdVaeriables"
    }

    Invoke-Expression $cmd
}

function Get-Package () {
    param (
        $build,
        $packageName,
        $outputFolder
    )
    Write-Host "Download package $packageName"

    $path = "D:\XCritical\packages\$outputFolder"
    Write-Host Download package to $path
    If(!(test-path $path)){
        New-Item -ItemType Directory -Force -Path $path
    }

    Remove-Item $path\* -Force -Recurse -Confirm:$false
    nuget install $packageName -Version $build -Source https://proget.unitup.space/nuget/Release/ -OutputDirectory $path
}

function Confirm-Arg () {
    param (
        [string]$name,
        [string]$arg
    )

    Write-Host "$name is $arg"
    
    if($null -eq $arg){
        Throw "$($name) value is empty"
    }

    while ($true) {
        $rerun = Read-Host "Continue (y/n)?"
    
       if($rerun -eq "n") {
           Exit 0
       }
       if($rerun -eq "y"){
           Break
       }
    }
}

function Initialize-Service () {
    param (
        [string]$serviceName,
        [PSCredential]$credential,
        [string]$serviceFolder,
        [string]$programFileName,
        $serviceParameters = $null
    )
    $service = Get-Service -Name $serviceName -ErrorAction SilentlyContinue

    if($null -eq $service){
        Write-Host Setup $serviceName

        $binaryPathName = "D:\XCritical\services\$serviceFolder\$programFileName.exe"

        if($serviceParameters -is [hashtable] -and  $null -ne $serviceParameters){
            foreach ($param in $serviceParameters.GetEnumerator()){
                $binaryPathName += " $($param.Name) $($param.Value)"
            }
        }

        if($serviceParameters -is [array] -and $null -ne $serviceParameters){
            foreach ($param in $serviceParameters){
                $binaryPathName += " $param"
            }
        }
        
        Write-Host BinaryPathName is $binaryPathName

        New-Service -Name $serviceName -BinaryPathName $binaryPathName -Credential $credential -StartupType Manual
        Start-Service -Name $serviceName
        Write-Host $serviceName is now running -ForegroundColor Green
    }
    else {
        Write-Host Service $serviceName alredy exist
    } 
    
}

function Uninstall-Service {
    param (
        [string]$serviceName
    )
    $service = Get-Service -Name $serviceName -ErrorAction SilentlyContinue

    if($null -ne $service){
        Write-Host Uninstall $serviceName
        Stop-Service -Name $serviceName
        Remove-Service -Name $serviceName
    }
    else {
        Write-Host Service $serviceName dosn''t exist
    }  
}

function Confirm-PowerShellMajorVersion () {
    param (
        $version
    )
    
    if($version -lt $host.Version.Major){
        throw "Your power shell version don't support this script. Current version is $host.Version Version must be $version.0 or higher. "+ `
        "Visit link https://docs.microsoft.com/ru-ru/powershell/scripting/install/installing-powershell-on-windows?view=powershell-7#msi"
    }
}

function Stop-ServiseIfExists() {
    param (
        $serviceName
    )
    $service = Get-Service -Name $serviceName -ErrorAction SilentlyContinue

    if($null -ne $service) {
        if($service.Status -ne "Stopped"){
            Stop-Service -Name $serviceName
            while ($true) { 
                $service = Get-Service -Name $serviceName
                if($service.Status -ne "Stopped"){
                    Start-Sleep -Seconds 1
                }
                else {
                    Break
                }
            }
        }
        Write-Host Service $serviceName is stoped -ForegroundColor Green
    }
    else {
        Write-Host Service $serviceName dosn''t exist -ForegroundColor Yellow
    }
    
}

function Start-ServiseIfExists(){
    param (
        $serviceName
    )
    $service = Get-Service -Name $serviceName -ErrorAction SilentlyContinue

    if($null -ne $service) {
        Start-Service -Name $serviceName
        Write-Host Service $serviceName is now running -ForegroundColor Green
    }
    else {
        Write-Host Service $serviceName dosn''t exist -ForegroundColor Yellow
    }
}

function Move-Files {
    param (
        $destination,
        $source,
        $exclude = ""
    )

    If(!(test-path $destination)){
      New-Item -ItemType Directory -Force -Path $destination
      $exclude = ""
    }

    if($exclude -eq ""){
        Remove-Item $destination\* -Force -Recurse -Confirm:$false
        Copy-Item -Path $source\* -Recurse -Destination $destination -Force
    }
    else {
        Remove-Item $destination\* -Force -Recurse -Confirm:$false -Exclude $exclude
        Copy-Item -Path $source\* -Recurse -Destination $destination -Force -Exclude $exclude  
    }
}

function Stop-IISSiteIfExists() {
    param (
        $iisSiteName
    )

    $iisSite = Get-IISSite -Name $iisSiteName
    if($null -ne $iisSite) {
        if($iisSite.State -ne "Stopped"){
            Stop-IISSite -Name $iisSiteName -Confirm:$False
            while ($iisSite.State -ne "Stopped") { 
                Start-Sleep -Seconds 1
            }
        }
        Write-Host IIS Site $iisSiteName is stoped -ForegroundColor Green
    }
    else {
        Write-Host IIS Site $iisSiteName dosn''t exist -ForegroundColor Yellow
    }
}

function Stop-IISAppPollIfExists(){
    param (
        $appPoolName
    )

    $appPoll = Get-IISAppPool $appPoolName
    if($null -ne $appPoll) {
        if($appPoll.State -ne "Stopped"){
            $appPoll.Stop()
            while($appPoll.State -ne "Stopped"){
                Start-Sleep -Seconds 1
            }
        }
        Write-Host App pool $appPoolName is stoped -ForegroundColor Green
    }
    else {
        Write-Host App pool $appPoolName dosn''t exist -ForegroundColor Yellow
    }
}

function Start-IISSiteIfExists(){
    param (
        $iisSiteName
    )
    $iisSite = Get-IISSite -Name $iisSiteName
    if($null -ne $iisSite) {
        try {
            Start-IISSite -Name $iisSiteName
            Write-Host IIS Site $iisSiteName is now started -ForegroundColor Green
        }
        catch {
            Write-Host $_
        }        
    }
    else {
        Write-Host IIS Site $iisSiteName dosn''t exist -ForegroundColor Yellow
    }
}

function Start-IISAppPollIfExists(){
    param (
        $appPoolName
    )
    $appPoll = Get-IISAppPool $appPoolName
    
    if($null -ne $appPoll) {
        try {
            $appPoll.Start()
            Write-Host App pool $appPoolName is now started -ForegroundColor Green
        }
        catch {
            Write-Host $_
        }
        
    }
    else {
        Write-Host App pool $appPoolName dosn''t exist -ForegroundColor Yellow
    }
}