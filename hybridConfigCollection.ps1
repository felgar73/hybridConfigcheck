<# Hybrid Configuration collection script
.Description
    Gathers data related to Exchange Hybrid configurations including: mailflow connectors, free/busy sharing & OAuth.
    
.Notes
    Script name: hybridConfigCollection.ps1
    Created by: Felix E. Garcia - felgar@microsoft.com
    Requirements: 
    -Powershell should to be 'Run As Administrator'.
    -Script assumes Kerberos Auth is enabled on-prem.

    -Supports EXO V2 Powershell module

    ::Updates::
    **Aug 2020**
    --Modified output file creation
    **Sept 2020**
    --Added HCW log parsing for 'Set-' cmdlets
    --Added EWS Collection function to search for specific servers

.Synopsis
    The script will first prompt you on whether you wish to collect on-premises Exchange data. Once complete (or if you answer 'no'), it will ask whether you wish to collect Exchange Online data. 
    Once complete, it will display the location of collection data. Some files will be in '.xml' format and some in text files. In some cases, data is sent to both formats for flexibility.

    When connecting to Exchange Online, the script will attempt to import the EXO V2 module; if this fails it will fallback to the legacy, Basic Auth remote method.

    If you wish to save previous output, make sure to rename output folders before re-running the script or files will be overwritten.
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
$hcwSetCmdOPtxt = '\HCW-SetCmdlets-OnPrem.txt'

#Cloud filename variables:
$CLIntraOrgxml = '\Cloud-IntraOrgConnector.xml'
$cLoauthtxt = '\Cloud-OauthConfigs.txt'
$cloudConfig = '\Cloud-SharingConfig.txt'
$cloudOrgReltxt = '\Cloud-OrgRelationship.txt'
$cloudOrgRelxml = '\Cloud-OrgRelationship.xml'
$hcwSetCmdCloudtxt = '\HCW-SetCmdlets-Cloud.txt'

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
$hcwSetOPPath = $onPremDir + $hcwSetCmdOPtxt

#Cloud output paths:
$CLIntraOrgpath = $cloudDir + $CLIntraOrgxml
$clOauthpath = $cloudDir + $cLoauthtxt
$cloudPath = $cloudDir + $cloudConfig
$hcwSetCloudpath = $cloudDir + $hcwSetCmdCloudtxt

#Hybrid Folder creation/validation:
$testPathD = Test-Path $outputDir
    if ($testPathD -eq $false){
    New-Item -itemtype Directory -Path $outputDir
    }

#Collecting HCW "XHCW" file info:
$hcwPath = Get-ChildItem "$env:APPDATA\Microsoft\Exchange Hybrid Configuration\*.xhcw" | ForEach-Object {Get-Content $_.fullname} -ErrorAction SilentlyContinue
[XML]$hcwLog = "<root>$($hcwPath)</root>"

#Find all 'Set-' cmdlets executed by HCW against OnPrem:
function HCWSetcmds-OnPrem {
if (!(!$hcwPath)){
$title1 = "====HCW 'Set-' Commands Executed On-Premises===="
$title1 | Out-File $hcwSetOPPath
$hcwLog.SelectNodes('//invoke') | Where-Object {$_.cmdlet -like "*Set*" -and $_.type -like "*OnPremises*"} | ForEach-Object {
    New-Object -Type PSObject -Property @{
        'Date'=$_.start;
        'Duration'=$_.duration;
        'Session'=$_.type;
        'Cmdlet'=$_.cmdlet;
        'Comment'=$_.'#comment'
        }
    } | Out-File -Append $hcwSetOPPath
        } else {
            Write-Host -ForegroundColor Cyan "No HCW logs found..."
        }
    }
#Find all 'Set-' cmdlets executed by HCW against EXO:
function HCWSetcmds-Cloud {
if (!(!$hcwPath)) {
$title2 = "====HCW 'Set-' Commands Executed in M365===="
$title2 | Out-File $hcwSetCloudpath
$hcwLog.SelectNodes('//invoke') | Where-Object {$_.cmdlet -like "*Set*" -and $_.type -like "*Tenant*"} | ForEach-Object {
    New-Object -Type PSObject -Property @{
        'Date'=$_.start;
        'Duration'=$_.duration;
        'Session'=$_.type;
        'Cmdlet'=$_.cmdlet;
        'Comment'=$_.'#comment'
        }
    } | Out-File -Append $hcwSetCloudpath
    } else {
        Write-Host -ForegroundColor Cyan "No HCW logs found..."
    }
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
    Write-Host -ForegroundColor Cyan "Checking for HCW logs..."
    HCWSetcmds-OnPrem
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

#EWS VDir Collection:
function EWS-VdirCollect {
Write-Host -ForegroundColor Yellow "Enter your MRS/Hybrid server names separated by commas (Ex: server1,server2):"
$hybsvrs = (Read-Host).split(",") | foreach {$_.trim()}
$ewsVdir = $hybsvrs | foreach {Get-WebServicesVirtualDirectory -Server $_ -ADPropertiesOnly}
$ewsVdir | Export-Clixml $ewsxmlPath
$ewsVdir | FL | Out-File $ewstxtPath
}

#On-Prem Data collection:
function OnPrem-Collection {
    Write-Host -ForegroundColor Cyan "Parsing HCW log..."
    HCWSetcmds-OnPrem
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
    
    $shpol =  "Sharing Policy Details:" 
    $shpol | Out-File $opPath
    $sharePol = Get-SharingPolicy 
    $sharePol | FL | Out-File -Append $opPath
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
    #EWS VDir collect function:
    EWS-VdirCollect
    
    #OAuth Config Details:
    $iOrgConn = Get-IntraOrganizationConnector
    if (!$iOrgConn){
        Write-Host -ForegroundColor Cyan "No IntraOrg Connector detected, OAUth may not be configured..."
        } else
            {
    $iorgtext = "IntraOrg Connector:"
    $iorgtext | Out-File $OPoauthPath
    $iOrgConn | FL | Out-File -Append $OPoauthPath
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
        Write-Host -ForegroundColor Cyan "Parsing HCW execution logs..."
        HCWSetcmds-Cloud
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
        }
    if ($null -ne $exoV2Fail) {
        Write-Host -ForegroundColor Cyan "EXO V2 connection failed. Trying basic auth connection..."
        $exoCred = Get-Credential -UserName $exoUPN -Message "Re-Enter M365 Admin Creds:"
        $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $exoCred -Authentication Basic -AllowRedirection -ErrorAction SilentlyContinue
    try {
        Import-PSSession $Session
    }catch {
       Write-Host -ForegroundColor Cyan "Remote Powershell to Exchange online failed, please try again..."
       exit
        }
}}
function EXO-Collection {
   #Prompt for domain if empty:
    if ($null -eq $domain) {
        Write-Host -ForegroundColor White "Enter your vanity domain name:"
        $domain = Read-Host
    }
Write-Host -ForegroundColor Cyan "Parsing HCW log..."
HCWSetcmds-Cloud

Write-Host -ForegroundColor Cyan "Collecting data from Exchange online..."
$shpol = "Sharing Policy Details:"
$shpol | Out-File $cloudPath
$sharePol = Get-SharingPolicy 
$sharePol |FL | Out-File -Append $cloudPath
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