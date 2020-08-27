<# Hybrid Configuration collection script
.Description
Gathers data related to Exchange Hybrid mailflow, free/busy sharing & OAuth.
Some data is output to text files, but most goes to XML files in order to avoid truncation issues.

.Notes
    Script name: hybridDetails.ps1
    Created by: Felix E. Garcia
    Requirements: 
    -Powershell should to be 'Run As Administrator'.
    -Script assumes Kerberos Auth is enabled on-prem.
    -Supports EXO V2 Powershell module

    Last Update: 8/27/20
    -Modified output file creation
#>

#Check for 'run as admin':
if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator"))
    {Write-host -ForegroundColor Yellow "Please close window and re-run powershell 'as administrator'."
    exit
    #Write-Host -ForegroundColor Cyan "& "C:\Users\$env:username\Desktop\Microsoft Exchange Online Powershell Module.appref-ms""
        }else {Write-Host -ForegroundColor Cyan "Checking execution policy..."
    }

#Execution Policy Check:
$execPol = Get-ExecutionPolicy
if ($execPol -ne 'Unrestricted'){
    Write-Host -ForegroundColor Cyan "Execution policy is" $execPol
    Write-Host -ForegroundColor Cyan "Changing policy to 'Unrestricted'..."
    Set-ExecutionPolicy Unrestricted -Force
}else {Write-Host -ForegroundColor Cyan "Execution policy is already '$execPol', continuing..."}

#On-Prem Filename variables:
$hybConfxml = '\Get-HybridConfig.xml'
$sendConnxml = '\Get-SendConnector.xml'
$sendConntxt = '\Get-Sendconnector.txt'
$recConnxml = '\Get-ReceiveConnector.xml'
$recConntxt = '\Get-ReceiveConnector.txt'
$opConfig = '\OnPrem-SharingConfig.txt'
$ewsOutxml = '\EWS-XMLOutput.xml'
$ewsOuttxt = '\EWS-TxtOutput.txt' 
$oPoauthtxt = '\OnPrem-OAuthConfig.txt'
$authsvrxml = '\Get-Authsvr.xml'
$OPintraOrgxml = '\OnPrem-IntraOrgConnector.xml'
$OPintraOrgtxt = '\OnPrem-IntraOrgConnector.txt'

#Cloud filename variables:
$CLIntraOrgxml = '\Cloud-IntraOrgConnector.xml'
$cLoauthtxt = '\Cloud-OauthConfigs.txt'
$cloudConfig = '\Cloud-SharingConfig.txt'
$cloudOrgReltxt = '\Cloud-OrgRelationship.txt'
$cloudOrgRelxml = '\Cloud-OrgRelationship.xml'

#Collection Folder variables:
$outputDir = "c:\temp\HybridConfigs"
$onPremDir = "$outputDir\OnPremCollection"
$cloudDir = "$outputDir\CloudCollection"

#OnPrem output paths:
$hybPath = $onPremDir + $hybConfxml
$sendConnXmlPath = $onPremDir + $sendConnxml
$sendConnTxtPath = $onPremDir + $sendConntxt
$recConnXmlPath = $onPremDir + $recConnxml
$recConnTxtPath = $onPremDir + $recConntxt
$opPath = $onPremDir + $opConfig
$ewsxmlPath = $onPremDir + $ewsOutxml
$ewstxtPath = $onPremDir + $ewsOuttxt 
$OPoauthPath = $onPremDir + $oPoauthtxt
$authsvrPath = $onPremDir + $authsvrxml
$OPintraOrgpath = $onPremDir + $OPintraOrgxml

#Cloud output paths:
$CLIntraOrgpath = $cloudDir + $CLIntraOrgxml
$clOauthpath = $cloudDir + $cLoauthtxt
$cloudPath = $cloudDir + $cloudConfig

#Hybrid folder creation/validation:
    $testPathD = Test-Path $outputDir
    if ($testPathD -eq $false){
    New-Item -itemtype Directory -Path $outputDir
    }

#OnPrem folder creation/validation:
function OnPremDir-Create {
    Write-Host -ForegroundColor Cyan "Creating collection folder..."
    $testPathOP = Test-Path $onPremDir
     if ($testPathOP -eq $false) {
     New-Item -itemtype Directory -Path $onPremDir 
    } 
}

#Cloud folder creation:
function CloudDir-Create {
    Write-Host -ForegroundColor White "Creating collection folder..."
    $testPathCL = Test-Path $cloudDir
    if ($testPathCL -eq $false) {
    New-Item -itemtype Directory -Path $cloudDir
    }
}

#Remote On-Prem Yes-No Function:
function OnPrem-RemoteQ {
Write-Host -ForegroundColor Yellow "Do you wish to collect on-premises data? Y/N:"
  $ans = Read-Host
    if ((!$ans) -or ($ans -eq 'y') -or ($ans -eq 'yes')){
        $ans = 'yes'
    #Prompt for domain name:
    Write-Host -ForegroundColor Yellow "Enter your vanity domain name:"
    $domain = Read-Host
    #Check/Create output folder:
    OnPremDir-Create
    #Connect to Exch on-prem:
    Write-Host -ForegroundColor White "Connecting to Exchange server..."
    Remote-ExchOnPrem
    #Collect data:
    OnPrem-Collection
    } else {
        $ans = 'no'
        Write-Host -ForegroundColor Cyan "Skipping on-premises data collection..."
    }
}
#Remote On-Prem Exchange Function:
function Remote-ExchOnPrem {
    Write-Host -ForegroundColor Yellow "Enter your On-Premises Exchange server FQDN:"
    $fqdn = Read-Host 
    $opCreds = Get-Credential -Message "Enter your Exchange admin credentials:" -UserName $env:USERDOMAIN\$env:USERNAME
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$fqdn/powershell/ -Credential $opCreds -Authentication Kerberos
       try {
        Import-PSSession $Session -DisableNameChecking -AllowClobber
       }
       catch {
           Write-Host -ForegroundColor Red "Failed to create remote session, please try again..."
           exit
       }
}
#On-Prem Data collection:
function OnPrem-Collection {
    Write-Host -ForegroundColor Cyan "Collecting Hybrid configuration details, please wait..."
    #Expand PS output:
    $fenlimit = $FormatEnumerationLimit
    if ($fenlimit -ne '-1'){
        $FormatEnumerationLimit=-1
    }
    #On-Prem Configs Collection:
    $hybConf = Get-HybridConfiguration 
    $hybConf | Export-Clixml $hybPath
    Start-Sleep -Seconds 2
    
    Add-Content $opPath -Value "Sharing Policy Details:"
    $sharePol = Get-SharingPolicy 
    $sharePol | FL | Out-File $opPath
    Add-Content $opPath -Value "=========="
    Start-Sleep -Seconds 2
    Add-Content $opPath -Value "Org Relationship Details:"
    $orgRel = Get-OrganizationRelationship 
    $orgRel |FL | Out-File -Append $opPath
    Add-Content $opPath -Value "=========="
    Start-Sleep -Seconds 2
    Add-Content $opPath -Value "Federation Information:"
    $fedInfo = Get-FederationInformation -DomainName $domain
    $fedInfo |FL | Out-File -Append $opPath
    Add-Content $opPath -Value "=========="
    Start-Sleep -Seconds 2
    Add-Content $opPath -Value "Organization Config Details:"
    $orgConfig = Get-OrganizationConfig
    $orgConfig |FL | Out-File -Append $opPath
    Add-Content $opPath -Value "=========="
    Start-Sleep -Seconds 2
    Add-Content $opPath -Value "Send Connector Details:"
    $sendConn = Get-SendConnector |?{$_.AddressSpaces -like '*onmicrosoft.com*'}
    $sendConn | Export-Clixml $sendConnxmlPath
    $sendConn |FL | Out-File $sendConnTxtPath
    Add-Content $opPath -Value "=========="
    Start-Sleep -Seconds 2
    Add-Content $opPath -Value "Receive Connector Details:"
    $recvConn = Get-ReceiveConnector |?{$_.TlsDomainCapabilities -like '*outlook*'}
    $recvConn | Export-Clixml $recConnxmlPath
    $recvConn |FL | Out-File $recConnTxtPath
    Start-Sleep -Seconds 2
    
    $ewsVdir = Get-WebServicesVirtualDirectory -ADPropertiesOnly 
    $ewsVdir | Export-Clixml $ewsxmlPath
    $ewsVdir | FL | Out-File $ewstxtPath
    
    #OAuth Config Details:
    $iOrgConn = Get-IntraOrganizationConnector
    if (!$iOrgConn){
        Write-Host -ForegroundColor Cyan "No IntraOrg Connector detected, OAUth may not be configured..."
        } else
            { 
    Add-Content $OPoauthPath -Value "IntraOrg Connector:"
    $iOrgConn |fl | Out-File $OPoauthPath
    Add-Content $OPoauthPath -Value "=========="
    Add-Content $OPoauthPath -Value "IntraOrganization Configs:"
    $iOrgConf = Get-IntraOrganizationConfiguration -WarningAction:SilentlyContinue
    $iOrgConf |FL | Out-File -Append $OPoauthPath
    Add-Content $OPoauthPath -Value "=========="
    Add-Content $OPoauthPath -Value "Partner Application Details:"
    $ptnrapp = Get-PartnerApplication 
    $ptnrapp |FL | Out-File -Append $OPoauthPath
    
    $authsvr = Get-AuthServer
    $authsvr | Export-Clixml $authsvrPath
    }
    
    #Close remote connection:
    Write-Host -ForegroundColor White "Collection complete -- closing connection to Exchange..."
    Get-PSSession | Remove-PSSession
    Start-Sleep -Seconds 2
    }
#Remote EXO Yes-No Function:
function EXO-RemoteQ {
    Write-Host -ForegroundColor Yellow "Do you wish to collect M365 data? Y/N:"
    $ans = Read-Host
      if ((!$ans) -or ($ans -eq 'y') -or ($ans -eq 'yes')){
          $ans = 'yes'
          #Create collection folder:
          CloudDir-Create
          #Connect to EXO:
          Remote-EXOPS
          #Collect data:
          EXO-Collection
      } else {
          $ans = 'no'
          Write-Host -ForegroundColor Cyan "Skipping M365 data collection..."
      }
  }
#Remote EXO PS Function:
function Remote-EXOPS {
    Write-Host -ForegroundColor Cyan "Connecting to Exchange online..."
    Write-Host -ForegroundColor Yellow "Enter your M365 admin credentials (Ex: admin@yourdomain.onmicrosoft.com):"
    $exoUPN = Read-Host
    try { 
        Import-Module ExchangeOnlineManagement
        Connect-ExchangeOnline -UserPrincipalName $exoUPN -ShowProgress $true
    } catch {
        $exoV2Fail = "EXO V2 Connection Failed"
        Write-Host -ForegroundColor Cyan "EXO V2 reomte connection failed..."
        }
    if ($null -ne $exoV2Fail) {
        Write-Host -ForegroundColor White "Trying basic auth connection..."
        $exoCred = Get-Credential -UserName $exoUPN -Message "Re-Enter M365 Admin Creds:"
        $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $exoCred -Authentication Basic -AllowRedirection -ErrorAction SilentlyContinue
    try {
        Import-PSSession $Session
    }catch {
       Write-Host -ForegroundColor Red "Remote Powershell to Exchange online failed, please try again..."
       exit
        }
}}
function EXO-Collection {
   #Prompt for domain if empty:
    if ($null -eq $domain) {
        Write-Host -ForegroundColor White "Enter your on-prem domain name:"
        $domain = Read-Host
    }
Write-Host -ForegroundColor Cyan "Collecting data from Exchange online..."

Add-Content $cloudPath -Value "Sharing Policy Details:"
$sharePol = Get-SharingPolicy 
$sharePol |FL | Out-File $cloudPath
Add-Content $cloudPath -Value "=========="
Start-Sleep -Seconds 2
Add-Content $cloudPath -Value "Org Relationship Details:"
$orgRel = Get-OrganizationRelationship 
$orgRel |FL | Out-File -Append $cloudPath
Add-Content $cloudPath -Value "=========="
Start-Sleep -Seconds 2
Add-Content $cloudPath -Value "Organization Config Details:"
$orgConfig = Get-OrganizationConfig 
$orgConfig |FL | Out-File -Append $cloudPath
Add-Content $cloudPath -Value "=========="
Add-Content $cloudPath -Value "IntraOrg Connector:"
$CliOrgConn = Get-IntraOrganizationConnector
$CliOrgConn |Export-Clixml $CLIntraOrgpath
Add-Content $cloudPath -Value "=========="
Start-Sleep -Seconds 2
Add-Content $cloudPath -Value "O365 Outbound Connector Details:"
$o365OutConn = Get-OutboundConnector |? {$_.enabled -eq 'true'}
$o365OutConn |FL | Out-File -Append $cloudPath
Add-Content $cloudPath -Value "=========="
Add-Content $cloudPath -Value "O365 Inbound Connector Details:"
$o365InConn = Get-InboundConnector |? {$_.enabled -eq 'true'}
$o365InConn |FL | Out-File -Append $cloudPath
Add-Content $cloudPath -Value "===End of File==="

Write-Host -ForegroundColor Cyan "Collection complete -- closing connection to Exchange online..."
Start-Sleep -Seconds 2
Get-PSSession | Remove-PSSession
} 
#On-prem Collection Prompt:
$ans = OnPrem-RemoteQ

#EXO Collection Prompt:
$ans = EXO-RemoteQ

#Goodbye:
 Write-Host -ForegroundColor White "Review or submit any files located in '$outputDir' to Microsoft support."

#Revert execution policy if needed:
if ($execPol -ne 'Unrestricted'){
    Write-Host -ForegroundColor Cyan "Changing execution policy back to '$execPol'..."
    Set-ExecutionPolicy $execPol -Force
}