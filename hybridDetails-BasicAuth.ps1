<#
.Synopsis
    --This script ensures you are running powershell as admin, then prompts for several variables.
    --Next asks if you wish to collect on-premises data first; if not you can skip to O365 data collection, which you can also choose to skip.
    --For on-premises collection it will prompt for a server fqdn and for the admin's credentials for Basic authentication.
    --For O365 data the admin will be prompted for credentials as well.
    --Once complete, it will close any powershell sessions that were created and instruct the admin to submit data files to Microsoft support.
    

.Description
    --This script collects data related to Exchange/O365 Hybrid Configurations and outputs it to two separate text files.
    --For admins connecting from a non-domain joined machine (external to LAN).
    :::Requirements:::
    --Powershell should be 'Run As Administrator'.
    --MFA not supported in this version.

.Notes
    Script name: hybridDetails-BA.ps1
    Created by: Felix E. Garcia
    ::Version 1.2.1::
    Last Updated: 9/30/2019
#>
<#
    :::Script Body:::
#>
#Check for 'run as admin':
if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator"))
    {Write-host -ForegroundColor Yellow "Please close window and re-run powershell 'as administrator'"
    exit
    }else {Write-Host -ForegroundColor Cyan "Checking execution policy..."}

#Execution Policy Check:
$execPol = Get-ExecutionPolicy
if ($execPol -ne 'Unrestricted'){
    Write-Host -ForegroundColor Cyan "Execution policy is" $execPol
    Write-Host -ForegroundColor Cyan "Changing policy to 'Unrestricted'..."
    Set-ExecutionPolicy Unrestricted -Force
}else {Write-Host -ForegroundColor Green "Execution policy is already '$execPol', continuing..."}

#Initial Variables:
$opConfig = 'OP-Config.txt'
$cloudConfig = 'Cloud-Config.txt'
$ewsConfig = 'WebSrvc-VDir.txt' 

#Prompt for domain name:
Write-Host -ForegroundColor Yellow "Enter your vanity domain name:"
$domain = Read-Host

#Output Directory test:
Write-Host -ForegroundColor Yellow "Enter the folder path or share location for data output:"
$dirPath = Read-Host 
$testPathD = Test-Path $dirPath
if ($testPathD -eq $false){
New-Item -itemtype Directory -Path $dirpath
}

#Build on-prem remote session:
Write-Host -ForegroundColor Yellow "Do you wish to collect on-premises data? Y/N:"
$opQ = Read-Host

if ((!$opQ) -or ($opQ -eq 'y') -or ($opQ -eq 'yes')){
    $opQ = 'yes'
}else {$opQ = 'no'}

if ($opQ -eq 'yes'){
Write-Host -ForegroundColor Yellow "Enter the FQDN of your on-premises Exchange server:"
$fqdn = Read-Host 
Write-Host -ForegroundColor Yellow "Enter your Exchange admin credentials in 'domain\username' format:"
$opCreds = Get-Credential -Message "Enter your Exchange admin credentials:"
$Session1 = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$fqdn/powershell/ -Credential $opCreds -Authentication Basic -AllowRedirection
    if ($Session1 -eq $null){
        Write-Host -ForegroundColor Red "Failed to create remote session - exiting..."
        Write-Host -ForegroundColor Red "Ensure that you have 'basic authentication' enabled for your Powershell virtual directory..."
        exit
    }else {
        Write-Host -ForegroundColor Cyan "Connecting to server $fqdn..."
        Import-PSSession $Session1 -DisableNameChecking
        }

#Output file creation:
Write-Host -ForegroundColor Cyan "Creating output files in $dirPath..."
Start-Sleep -Seconds 2
#New-Item -Type File -Name $opConfig -Path $dirPath -ErrorAction SilentlyContinue
#New-Item -Type File -Name $ewsConfig -Path $dirPath -ErrorAction SilentlyContinue
#On Prem Data collection:
Write-Host -ForegroundColor Cyan "Collecting Hybrid configuration details - please wait..."
#Expand PS output:
$fenlimit = $FormatEnumerationLimit
if ($fenlimit -ne '-1'){
    $FormatEnumerationLimit=-1
}
Set-Location $dirPath
#Add-Content $opConfig -Value "Sharing Policy Details:"
$sharePol = Get-SharingPolicy 
$sharePol | FL | Out-File $opConfig
#Add-Content $opConfig -Value "Sharing Policy"
Add-Content $opConfig -Value "=========="
Start-Sleep -Seconds 2
Add-Content $opConfig -Value "Org Relationship Details:"
$orgRel = Get-OrganizationRelationship 
$orgRel |FL | Out-File -Append $opConfig
Add-Content $opConfig -Value "=========="
Start-Sleep -Seconds 2
Add-Content $opConfig -Value "Hybrid Config Details:"
$hybConf = Get-HybridConfiguration 
$hybConf |FL | Out-File -Append $opConfig
Add-Content $opConfig -Value "=========="
Start-Sleep -Seconds 2
Add-Content $opConfig -Value "Federation Information:"
$fedInfo = Get-FederationInformation -DomainName $domain
$fedInfo |FL | Out-File -Append $opConfig
Add-Content $opConfig -Value "=========="
Start-Sleep -Seconds 2
Add-Content $opConfig -Value "Organization Config Details:"
$orgConfig = Get-OrganizationConfig 
$orgConfig |FL | Out-File -Append $opConfig
Add-Content $opConfig -Value "=========="
Start-Sleep -Seconds 2
Add-Content $opConfig -Value "Send Connector Details:"
$sendConn = Get-SendConnector |?{$_.AddressSpaces -like '*onmicrosoft.com*'}
$sendConn |FL | Out-File -Append $opConfig
Add-Content $opConfig -Value "=========="
Start-Sleep -Seconds 2
Add-Content $opConfig -Value "Receive Connector Details:"
$recvConn = Get-ReceiveConnector |?{$_.TlsDomainCapabilities -like '*outlook*'}
$recvConn |FL | Out-File -Append $opConfig
Add-Content $opConfig -Value "End of File"
Start-Sleep -Seconds 2

$ewsVdir = Get-WebServicesVirtualDirectory -ADPropertiesOnly 
$ewsVdir | FL Server,*AuthenticationMethods,*Url,MRSproxyEnabled| Out-File $ewsConfig
Add-Content $opConfig -Value "End of File"

Write-Host -ForegroundColor Cyan "Collection complete -- closing connection to Exchange..."
Remove-PSSession $Session1
Start-Sleep -Seconds 2
}else {
    Write-Host -ForegroundColor Cyan "Skipping on-premises data collection..."
    }

#O365 Data Prompt:
Write-Host -ForegroundColor Yellow "Do you wish to collect O365 data? Y/N:"
$cloudQ = Read-Host

if ((!$cloudQ) -or ($cloudQ -eq 'y') -or ($cloudQ -eq 'yes')){
    $cloudQ = 'yes'
}else {$cloudQ = 'no'}

#Build remote EXO session:
if ($cloudQ -eq 'yes') {
$exoCred = Get-Credential -Message "Enter Office 365 admin credentials:"
$Session2 = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $exoCred -Authentication Basic -AllowRedirection
if (!$session2){
    Write-Host -ForegroundColor Red "Failed to create remote session - exiting..."
    exit
}else {
    Write-Host -ForegroundColor Cyan "Connecting to Exchange online service..."
    Import-PSSession $Session2
    }

#Cloud data collection:
Write-Host -ForegroundColor Cyan "Collecting data from Exchange online..."
New-Item -Type File -Name $cloudConfig -Path $dirPath -ErrorAction SilentlyContinue
Start-Sleep -Seconds 1
Set-Location $dirPath
Add-Content $cloudConfig -Value "Sharing Policy Details:"
$sharePol = Get-SharingPolicy 
$sharePol |FL | Out-File $cloudConfig
Add-Content $cloudConfig -Value "=========="
Start-Sleep -Seconds 2
Add-Content $cloudConfig -Value "Org Relationship Details:"
$orgRel = Get-OrganizationRelationship 
$orgRel |FL | Out-File -Append $cloudConfig
Add-Content $cloudConfig -Value "=========="
Start-Sleep -Seconds 2
Add-Content $cloudConfig -Value "Organization Config Details:"
$orgConfig = Get-OrganizationConfig 
$orgConfig |FL | Out-File -Append $cloudConfig
Add-Content $cloudConfig -Value "=========="
Add-Content $cloudConfig -Value "O365 Outbound Connector Details:"
$o365OutConn = Get-OutboundConnector |? {$_.enabled -eq 'true'}
$o365OutConn |FL | Out-File -Append $cloudConfig
Add-Content $cloudConfig -Value "=========="
Add-Content $cloudConfig -Value "O365 Inbound Connector Details:"
$o365InConn = Get-InboundConnector |? {$_.enabled -eq 'true'}
$o365InConn |FL | Out-File -Append $cloudConfig
Add-Content $cloudConfig -Value "End of File"

Write-Host -ForegroundColor Cyan "Collection complete -- closing connection to Exchange online..."
Start-Sleep -Seconds 2
Remove-PSSession $Session2
} else {
    Write-Host -ForegroundColor Cyan "Skipping O365 data collection..."
    Start-Sleep -Seconds 2
    }

#Goodbye:
if (($opQ -eq 'yes') -or ($cloudQ -eq 'yes')) {
    Write-Host -ForegroundColor Yellow "Submit the following files to Microsoft:"
    Get-ChildItem $dirPath |FT LastWriteTime,Name
}else {
        Write-Host -ForegroundColor Cyan "No files to submit."
    }
#Revert execution policy if needed:
if ($execPol -ne 'Unrestricted'){
    Write-Host -ForegroundColor Cyan "Changing execution policy back to '$execPol'..."
    Set-ExecutionPolicy $execPol -Force
}