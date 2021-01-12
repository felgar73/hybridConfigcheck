<# Hybrid Configuration collection script
**Disclaimer**
This script is NOT an official Microsoft tool. Therefore use of the tool is covered exclusively by the license associated with this github repository.
THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. 
IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

.Description
    Gathers data related to Exchange Hybrid configurations including: mailflow connectors, free/busy sharing & OAuth.

.Notes
    Script name: hybridConfigCollection.ps1
    Created by: Felix E. Garcia - felgar@microsoft.com

    Requirements:
    -Powershell should to be 'Run As Administrator'.
    -Script assumes Kerberos Auth is enabled on-prem.

    General Notes:
    -Supports EXO V2 Powershell module
    -Details collected for Exchange certificates will be limited due to remote powershell limitations. If more Exchange cert data is needed go ahead and run the 'Get-ExchangeCertificate' cmdlet locally on the desired server.

    You can run the script on any domain-joined machine via a regular Powershell window (as admin) (not Exchange Mgmt Shell).    
    The script will first prompt you on whether you wish to collect on-premises Exchange data. Once complete (or if you answer 'no'), it will ask whether you wish to collect Exchange Online data. 
    Once complete, it will display the location of collection data. Some files will be in '.xml' format and some in text files. In some cases, data is sent to both formats for flexibility.

    When connecting to Exchange Online, the script will attempt to import the EXO V2 powershell module; if this fails it will fallback to the legacy, Basic Auth remote method.
    When connecting to Azure AD in order to collect OAuth some settings, the script will attempt to import the AzureADPreview powershell module; if this fails it will continue with other operations and this will require manual collection.

    ::Updates::
    **Aug 2020**
    --Modified output file creation
    **Sept 2020**
    --Added HCW log parsing for 'Set-' cmdlets
    --Added EWS Collection function to search for specific servers
    **Oct 2020**
    --Added Exch certificate collection function
    --Added 'New-' cmdlet search within HCW logs
    **Nov 2020**
    --Modified folder creation 'Test-Path' functions
    **Dec 2020**
    --Added MSOL/AAD data collection for OAuth configs
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
$exchCerttxt = '\ExchangeCertificate.txt'
$opsharConfig = '\OnPrem-SharingConfig.txt'
$ewsOutxml = '\EWS-XMLOutput.xml'
$ewsOuttxt = '\EWS-TxtOutput.txt' 
$oPoauthtxt = '\OnPrem-OAuthConfigs.txt'
$authsvrxml = '\Get-AuthServer.xml'
$OPintraOrgxml = '\OnPrem-IntraOrgConnector.xml'
$OPintraOrgtxt = '\OnPrem-IntraOrgConnector.txt'
$hcwSetCmdOPtxt = '\HCW-Cmds_OnPrem.txt'

#Cloud filename variables:
$clIntraOrgtxt = '\Cloud-IntraOrgConnector.txt'
$CLIntraOrgxml = '\Cloud-IntraOrgConnector.xml'
$cloauthtxt = '\Cloud-OauthConfigs.txt'
$cloudsharetxt = '\Cloud-SharingConfig.txt'
$cloudOrgtxt = '\Cloud-OrgConfig.txt'
$cloudOrgReltxt = '\Cloud-OrgRelationship.txt'
$cloudOrgRelxml = '\Cloud-OrgRelationship.xml'
$hcwSetCmdCloudtxt = '\HCW-Cmds_Cloud.txt'
$inbConntxt = '\O365-Inbound-Connector.txt'
$outbConntxt = '\O365-Outbound-Connector.txt'
$migEndptxt = '\O365-Migration-Endpoints.txt'
$accptDomtxt = '\O365-Accepted-Domain.txt'

#Collection Folder variables:
$outputDir = 'c:\temp\HybridConfigs'
$onPremDir = $outputDir + '\OnPremCollection'
$cloudDir = $outputDir + '\CloudCollection'

#OnPrem output paths:
$hybxmlPath = $onPremDir + '\Hybrid-Config.xml'
$hybtxtPath = $onPremDir + '\Hybrid-Config.txt'
$sendConnXmlPath = $onPremDir + '\Get-SendConnector.xml'
$sendConnTxtPath = $onPremDir + '\Get-Sendconnector.txt'
$recConnXmlPath = $onPremDir + '\Get-ReceiveConnector.xml'
$recConnTxtPath = $onPremDir + '\Get-ReceiveConnector.txt'
$exchCertpath = $onPremDir + '\ExchangeCertificate.txt'
$opsharPath = $onPremDir + '\OnPrem-SharingConfig.txt'
$ewsxmlPath = $onPremDir + '\EWS-XMLOutput.xml'
$ewstxtPath = $onPremDir + '\EWS-TxtOutput.txt' 
$opOauthPath = $onPremDir + '\OnPrem-OAuthConfigs.txt'
$authsvrxmlPath = $onPremDir + '\Get-AuthServer.xml'
$hcwSetOPPath = $onPremDir + '\HCW-Cmds_OnPrem.txt'
#$OPintraOrgxmlpath = $onPremDir + '\OnPrem-IntraOrgConnector.xml'
#$OPintraOrgtxtpath = $onPremDir + '\OnPrem-IntraOrgConnector.txt'

#Cloud output paths:
$clIntraOrgtxtpath = $cloudDir + $clIntraOrgtxt
$clIntraOrgxmlpath = $cloudDir + $CLIntraOrgxml
$clOauthpath = $cloudDir + $cloauthtxt
$cloudPath = $cloudDir + $cloudsharetxt
$cloudOrgpath = $cloudDir + $cloudOrgtxt
$hcwSetCloudpath = $cloudDir + $hcwSetCmdCloudtxt
$inbConnpath = $cloudDir + $inbConntxt
$outbConnpath = $cloudDir + $outbConntxt
$migEndpath = $cloudDir + $migEndptxt
$accDompath = $cloudDir + $accptDomtxt

#Hybrid Folder creation/validation:
    if (!(Test-Path $outputDir)){
    New-Item -itemtype Directory -Path $outputDir
    }

#Collecting HCW "XHCW" file info:
$hcwPath = Get-ChildItem "$env:APPDATA\Microsoft\Exchange Hybrid Configuration\*.xhcw" | ForEach-Object {Get-Content $_.fullname} -ErrorAction SilentlyContinue
[XML]$hcwLog = "<root>$($hcwPath)</root>"

#Find all 'Set-' cmdlets executed by HCW against OnPrem:
function HCWSetcmds-OnPrem {
if (!(!$hcwPath)){
$title1 = "===='Set-' Commands Executed On-Premises===="
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

$title2 = "===='New-' Commands Executed On-Premises===="
$title2 | Out-File -Append $hcwSetOPPath
$hcwLog.SelectNodes('//invoke') | Where-Object {$_.cmdlet -like "*New*" -and $_.type -like "*OnPremises*"} | ForEach-Object {
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
$title3 = "===='Set-' Commands Executed in M365===="
$title3 | Out-File $hcwSetCloudpath
$hcwLog.SelectNodes('//invoke') | Where-Object {$_.cmdlet -like "*Set*" -and $_.type -like "*Tenant*"} | ForEach-Object {
    New-Object -Type PSObject -Property @{
        'Date'=$_.start;
        'Duration'=$_.duration;
        'Session'=$_.type;
        'Cmdlet'=$_.cmdlet;
        'Comment'=$_.'#comment'
        }
    } | Out-File -Append $hcwSetCloudpath

$title4 = "===='New-' Commands Executed in M365===="
$title4 | Out-File -Append $hcwSetCloudpath
$hcwLog.SelectNodes('//invoke') | Where-Object {$_.cmdlet -like "*New*" -and $_.type -like "*Tenant*"} | ForEach-Object {
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
    if (!(Test-Path $onPremDir)) {
     New-Item -itemtype Directory -Path $onPremDir 
    } 
}

#Cloud folder creation:
function CloudDir-Create {
    Write-Host -ForegroundColor Cyan "Creating collection folder..."
    if (!(Test-Path $cloudDir)) {
    New-Item -itemtype Directory -Path $cloudDir
    }
}

#Collect On-Prem Yes-No Function:
function OnPrem-RemoteQ {
Write-Host -ForegroundColor Yellow "Do you wish to collect on-premises data? Y/N:"
  $ans = Read-Host
    if ((!$ans) -or ($ans -eq 'y') -or ($ans -eq 'yes')){
        $ans = 'yes'
    #Enter Hybrid server names:
    Write-Host -ForegroundColor Yellow "Enter your MRS/Hybrid server names separated by commas (Ex: server1,server2):"
    $hybsvrs = (Read-Host).split(",") | foreach {$_.trim()}
    #Check/Create output folder:
    OnPremDir-Create
    Write-Host -ForegroundColor Cyan "Checking for HCW logs..."
    HCWSetcmds-OnPrem
    #Connect to Exch on-prem:
    Write-Host -ForegroundColor Cyan "Creating remote Exchange session..."
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
#Exch Certs Function:
function ExchCert-Collection {
        $exchCerts = $hybsvrs | foreach {Get-ExchangeCertificate -Server $_ }
        $exchCerts | FL | Out-File $onPremDir\Exch-Certs.txt
}
#EWS VDir Collection:
function EWS-VdirCollect {
    $ewsVdir = $hybsvrs | foreach {Get-WebServicesVirtualDirectory -Server $_ -ADPropertiesOnly}
    $ewsVdir | Export-Clixml $ewsxmlPath
    $ewsVdir | FL | Out-File $ewstxtPath
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
    $hybConf | FL | Out-File $hybtxtPath
    $hybConf | Export-Clixml $hybxmlPath
    Start-Sleep -Seconds 2
    
    $shpol =  "===Sharing Policy Details===:" 
    $shpol | Out-File $opsharPath
    $sharePol = Get-SharingPolicy 
    $sharePol | FL | Out-File -Append $opsharPath
    Start-Sleep -Seconds 2
    Add-Content $opsharPath -Value "===Org Relationship Details===:"
    $orgRel = Get-OrganizationRelationship 
    $orgRel |FL | Out-File -Append $opsharPath
    Start-Sleep -Seconds 2
    Add-Content $opsharPath -Value "===Federation Information===:"
    $fedInfo = Get-FederationInformation -DomainName $domain
    $fedInfo |FL | Out-File -Append $opsharPath
    Start-Sleep -Seconds 2
    Add-Content $opsharPath -Value "===Organization Config Details===:"
    $orgConfig = Get-OrganizationConfig
    $orgConfig |FL | Out-File -Append $opsharPath
    Start-Sleep -Seconds 2
    #Exch Certs:
    ExchCert-Collection
    #Mail Connectors:
    $sendTitle = "===Send Connector Details===:"
    $sendtitle | Out-File $sendConnTxtPath
    $sendConn = Get-SendConnector |? {$_.AddressSpaces -like '*onmicrosoft.com*'}
    $sendConn | Export-Clixml $sendConnxmlPath
    $sendConn | FL | Out-File -Append $sendConnTxtPath
    Start-Sleep -Seconds 2
    $recTitle = "===Receive Connector Details===:"
    $recTitle  | Out-File $recConnTxtPath
    $recvConn = Get-ReceiveConnector |?{$_.TlsDomainCapabilities -like '*outlook*'}
    $recvConn | Export-Clixml $recConnxmlPath
    $recvConn |FL | Out-File -Append $recConnTxtPath

    Start-Sleep -Seconds 2
    #EWS VDir collect function:
    EWS-VdirCollect
    
    #OAuth Config Details:
    $iOrgConn = Get-IntraOrganizationConnector
    if (!$iOrgConn){
        Write-Host -ForegroundColor Cyan "No IntraOrg Connector detected, OAUth may not be configured..."
        } else
            {
    $iorgtext = "===IntraOrg Connector===:"
    $iorgtext | Out-File $opOauthPath
    $iOrgConn | FL | Out-File -Append $OPoauthPath
    Add-Content $OPoauthPath -Value "===IntraOrganization Configs===:"
    $iOrgConf = Get-IntraOrganizationConfiguration -WarningAction:SilentlyContinue
    $iOrgConf |FL | Out-File -Append $OPoauthPath
    Add-Content $OPoauthPath -Value "===Partner Application Details===:"
    $ptnrapp = Get-PartnerApplication 
    $ptnrapp |FL | Out-File -Append $OPoauthPath
    Add-Content $opOauthPath -Value "===Auth Server Settings===:"
    $authsvr = Get-AuthServer
    $authsvr | FL Name,type,realm,TokenIssuingEndpoint,AuthorizationEndpoint | Out-File -Append $opOauthPath
    Add-Content $OPoauthPath -Value "**Additional Auth Server details found in XML file."
    $authsvr | Export-Clixml $authsvrxmlPath
    }

    #Close remote connection:
    Write-Host -ForegroundColor Green "Collection complete. Closing connection to Exchange..."
    Get-PSSession | Remove-PSSession
    Start-Sleep -Seconds 2
    }
#Collect EXO Yes-No Function:
function EXO-RemoteQ {
    Write-Host -ForegroundColor Yellow "Do you wish to collect M365 data? Y/N:"
    $ans = Read-Host
    if ((!$ans) -or ($ans -eq 'y') -or ($ans -eq 'yes')){
        $ans = 'yes'
        #Create collection folder:
        CloudDir-Create
        if ($null -eq $domain) {
            Write-Host -ForegroundColor Yellow "Enter your vanity domain name:"
            $domain = Read-Host
        }
        Write-Host -ForegroundColor Cyan "Checking HCW logs..."
        HCWSetcmds-Cloud
        #Connect to EXO:
        Remote-EXOPS
        #Collect data:
        EXO-Collection
        AAD-Collection
    } else {
            $ans = 'no'
            Write-Host -ForegroundColor Cyan "Skipping M365 data collection..."
        }
}
#Remote EXO PS Function:
function Remote-EXOPS {
    Write-Host -ForegroundColor Cyan "Connecting to Exchange Online..."
    Write-Host -ForegroundColor Yellow "Enter your M365 admin credentials (Ex: admin@yourdomain.onmicrosoft.com):"
    $exoUPN = Read-Host
    try { 
        Import-Module ExchangeOnlineManagement
        Connect-ExchangeOnline -UserPrincipalName $exoUPN -ShowBanner:$false
    } catch {
        $exoV2Fail = "EXO V2 Connection Failed"
        }
    if ($null -ne $exoV2Fail) {
        Write-Host -ForegroundColor Cyan "EXO V2 connection failed. Consider installing the EXO V2 module as basic auth is being deprecated (http://aka.ms/exopspreview)." 
        Write-Host -ForegroundColor Cyan "Trying basic auth connection..."
        $exoCred = Get-Credential -UserName $exoUPN -Message "Re-Enter M365 Admin Creds:"
        $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $exoCred -Authentication Basic -AllowRedirection -ErrorAction SilentlyContinue
    try {
        Import-PSSession $Session
    }catch {
       Write-Host -ForegroundColor Cyan "Remote Powershell to Exchange Online failed, please try again..."
       exit
        }
    }
}
function AAD-Collection {
Write-Host -ForegroundColor Cyan "Connecting to Azure AD..."
    try {
        Import-Module AzureADPreview
        Connect-AzureAD
    }
    catch {
        $msolFail = "Connection to Azure AD Failed. Ensure AAD Powershell Module (https://aka.ms/aadposh) is installed and try again..." 
    }
    if ($null -ne $msolFail) {
        Write-Host -ForegroundColor Cyan $msolFail
        Write-Host -ForegroundColor Cyan "Refer to: https://docs.microsoft.com/en-us/powershell/azure/active-directory/install-msonlinev1?view=azureadps-1.0 for additional details."
    } else {
        $svcPrinText = "=== EXO Service Principal for OAuth ==="
        $svcPrinText | Out-File -Append $clOauthpath
        $exoSvcid = '00000002-0000-0ff1-ce00-000000000000'
        $svcPrinc = Get-AzureADServicePrincipal -Filter "AppId eq '$exoSvcid'"
        $svcPrinc | FL | Out-File -Append $clOauthpath
        Add-Content $clOauthpath -Value "*** Expanded ServicePrincipalNames Output***"
        $svcPrinc | select -ExpandProperty ServicePrincipalNames | Out-File -Append $clOauthpath
    }
}
function EXO-Collection {
Write-Host -ForegroundColor Cyan "Collecting data from Exchange Online..."
$shpol = "===Sharing Policy Details===:"
$shpol | Out-File $cloudPath
$sharePol = Get-SharingPolicy 
$sharePol |FL | Out-File -Append $cloudPath
Add-Content $cloudPath -Value "===Org Relationship Details===:"
$orgRel = Get-OrganizationRelationship 
$orgRel |FL | Out-File -Append $cloudPath
Start-Sleep -Seconds 2

$cloudOrgtext = "===Organization Config Details===:"
$cloudOrgtext | Out-File $cloudOrgpath
$orgConfig = Get-OrganizationConfig 
$orgConfig |FL | Out-File -Append $cloudOrgpath

$migEndtext = "===Migration Endpoints===:"
$migEndtext | Out-File $migEndpath
$migEnd = Get-MigrationEndpoint
$migEnd | FL | Out-File -Append $migEndpath

#OAuth Configs:
$iOrgText = "===IntraOrg Connector===:"
$iOrgText | Out-File $clOauthpath
$CliOrgConn = Get-IntraOrganizationConnector
if ($null -ne $CliOrgConn) {
    $CliOrgConn | FL | Out-File -Append $clOauthpath
    #$CliOrgConn |Export-Clixml $CLIntraOrgpath
} else {
    $iOrgText2 = "***No OAuth Configs found***"
    $iOrgText2 | Out-File -Append $clOauthpath
}
Start-Sleep -Seconds 2

#MailFlow Configs:
$outbCtext = "===O365 Outbound Connector Details===:"
$outbCtext | Out-File $outbConnpath
$o365OutConn = Get-OutboundConnector |? {$_.enabled -eq 'true'}
$o365OutConn |FL | Out-File -Append $outbConnpath
$inbCtext = "===O365 Inbound Connector Details===:"
$inbCtext | Out-File $inbConnpath
$o365InConn = Get-InboundConnector |? {$_.enabled -eq 'true'}
$o365InConn |FL | Out-File -Append $inbConnpath

$accDtext = "===Accepted Domain===:"
$accDtext | Out-File $accDompath
$accDomain = Get-AcceptedDomain $domain
$accDomain | FL | Out-File -Append $accDompath

Write-Host -ForegroundColor White "Collection complete. Closing connection to Exchange Online..."
Get-PSSession | Remove-PSSession
}

#Prompt for domain name:
Write-Host -ForegroundColor Yellow "Enter your vanity domain name:"
$domain = Read-Host

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