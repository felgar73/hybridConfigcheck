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

    General Notes:
    -Script assumes Kerberos Auth is enabled on-prem for remote Exchange session.
    -Requires EXO V3 Powershell module for EXO collection.
    -Script can be run from any domain-joined machine, but it's recommended to run from Exchange Mgmt Shell directly and bypass remote powershell connection.
    -If using remote powershell, details collected for Exchange certificates will be limited due to a powershell limitation. 

    The script will first prompt you on whether you wish to collect on-premises Exchange data and whether or not a remote powershell session is needed for this (if running from local Exchange server, simply answer to this question). 
    Once On-prem collection is complete (or if you answer 'no'), it will ask whether you wish to collect Exchange Online data. 
    Once complete, it will display the location of collection data. Some files will be in '.xml' format and some in text files. In some cases, data is sent to both formats for flexibility.

    The script will attempt to import the MgGraph powershell module in order to collect OAuth configs from Entra ID; if this fails it will continue with other operations but this data will require manual collection.

    ###Updates:
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
    **Jan 2021**
    --Added HCW log parsing for 'Remove-' cmdlets
    --Added additional 'Get-' commands to be pulled
    **Feb 2021**
    --Added 'Get-FederatedOrganizationIdentifier' collection in EXO
    **April 2021**
    --Added option to bypass remote session to on-prem Exchange in the case you are already logged into an Exchange server.
    --Added 'Get-EmailAddressPolicy' to data collection functions.
    **May 2022**
    --Modified output for Partner Applicaiton details:
        --trucated output within OAuth Configs text file
        --added XML file for full output
    --Adding collection of Skype Integration configs
    **July 2022**
    --Added Silent Error action for Skype config & Federation checks
    **Sept 2022**
    --Removing Basic Auth as an optional login
    **April 2023**
    --Renamed variable for EXO v3 PS failure
    **Nov 2023**
    --Now collecting ALL EXO mailflow connectors including "test mode" connectors
    **Jan 2024**
    --Added Json output commands to replace xml output
    --Replaced MSO commands with MS Graph for Entra ID/OAuth data collection
    **Dec 2024**
    --Moved EXO-Collection function to avoid creating empty output files if PS connection fails.
    **Jan 2025**
    --Added additional Json output files.
    --Added Test-MigrationServerAvailability cmd
    **Aug 2025**
    --Modified MS Graph import to avoid application module error.
    --Added XML output for EXO first-party app details.
    --Added output for Exchange Hybrid Application details.
    **Oct 2025**
    --Removed MS Graph import command due to it hanging.
#>

#Check for 'run as admin':
if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator"))
    {Write-host -ForegroundColor Yellow "Please close window and re-run powershell 'as administrator'."
    exit
    #Write-Host -ForegroundColor Cyan "& "C:\Users\$env:username\Desktop\Microsoft Exchange Online Powershell Module.appref-ms""
        } else {Write-Host -ForegroundColor Cyan "Checking execution policy..."
    }

#Execution Policy Check:
$execPol = Get-ExecutionPolicy
if ($execPol -ne 'Unrestricted'){
    Write-Host -ForegroundColor Cyan "Execution policy is" $execPol
    Write-Host -ForegroundColor Cyan "Changing policy to 'Unrestricted'..."
    Set-ExecutionPolicy Unrestricted -Force
}else {Write-Host -ForegroundColor Cyan "Execution policy is already '$execPol', continuing..."}

#Collection Folder variables:
$date = Get-Date -UFormat %b-%d-%Y
$outputDir = 'c:\temp\HybridConfigs' + '_' + $date
$onPremDir = $outputDir + '\OnPremises'
$cloudDir = $outputDir + '\Office365'
$errorLog = $outputDir + '\ErrorLog.log'

#OnPrem output paths:
$hybConfigjson = $onPremDir + '\HybridConfig.json'
#$hybxmlPath = $onPremDir + '\Hybrid-Config.xml'
$hybtxtPath = $onPremDir + '\Hybrid-Config.txt'
$sendConnJson = $onPremDir + '\SendConnectors.json'
#$sendConnXmlPath = $onPremDir + '\SendConnector.xml'
$sendConnTxtPath = $onPremDir + '\Sendconnector.txt'
$recConnjson = $onPremDir + '\ReceiveConnectors.json'
#$recConnXmlPath = $onPremDir + '\ReceiveConnectors.xml'
$recConnTxtPath = $onPremDir + '\ReceiveConnectors.txt'
$exchCertjson = $onPremDir + '\Get-ExchangeCertificate.json'
$exchCertpath = $onPremDir + '\Exchange-Certificates.txt'
$opsharePolicyjson = $onPremDir + '\Get-SharingPolicy.json'
$opsharPath = $onPremDir + '\Sharing-Configs_OnPrem.txt'
$opfedorgIdjson = $onPremDir + '\Get-FederatedOrganizationIdentifier.json'
$opfederatConfpath = $onPremDir + '\Federation-Configs_OnPrem.txt'
$opfedTrustjson = $onPremDir + '\Get-FederationTrust.json'
$ewsVdirjson = $onPremDir + '\Get-WebServicesVirtualDirectory.json'
$ewstxtPath = $onPremDir + '\EWS-TxtOutput.txt'
$opOauthConfigjson = $onPremDir + '\OAuthConfig.json'
$opOauthPath = $onPremDir + '\OAuth-Configs_OnPrem.txt'
#$authsvrxmlPath = $onPremDir + '\Get-AuthServer.xml'
$authSvrjson = $onPremDir + '\Get-AuthServer.json'
$hcwLogsOPPath = $onPremDir + '\HCW-LogCmds_OnPrem.txt'
$opRemDomtxt = $onPremDir + '\RemoteDomains_OnPrem.txt'
$opRemdomjson = $onPremDir + '\Get-RemoteDomain.json'
$authConfigpath = $onPremDir + '\Get-AuthConfig_OnPrem.txt'
$OPaddpoltxt = $onPremDir + '\EmailAddressPolicy_OnPrem.txt'
$opAddpoljson = $onPremDir + '\Get-EmailAddressPolicy.json'
$opOrgConfigtxt = $onPremDir + '\OrganizationConfig-OnPrem.txt'
#$partnerAppxml = $onPremDir + '\PartnerApplication-OnPrem.xml'
$partnerAppjson = $onPremDir + '\Get-PartnerApplication.json'
$skypeIntTxt = $onPremDir + '\SkypeIntegration-Configs.txt'
$opIOConnjson = $onPremDir + '\Get-IntraOrganizationConnector.json'
$opOrgReljson =  $onPremDir + '\Get-OrganizationRelationship.json'
$OpacceptedDomJson = $onPremDir + '\Get-AcceptedDomain.json'
$allserversJson = $onPremDir + '\Get-AllServers.json'

#Cloud output paths:
$clIntraOrgtxtpath = $cloudDir + '\Get-IntraOrgConnector_EXO.txt'
#$clIntraOrgxmlpath = $cloudDir + '\Get-IntraOrgConnector_EXO.xml'
$clIOConnjson = $cloudDir + '\Get-IntraOrganizationConnector.json'
$clOauthpath = $cloudDir + '\Oauth-Configs_EXO.txt'
$clSharingPath = $cloudDir + '\Sharing-Configs_EXO.txt'
$exosharingpoljson = $cloudDir + '\Get-SharingPolicy.json'
$cloudOrgpath = $cloudDir + '\Get-OrganizationConfig_EXO.txt'
$hcwLogsCloudpath = $cloudDir + '\HCW-LogCmds_M365.txt'
$inbConnpath = $cloudDir + '\Inbound-Connector_EXO.txt'
$inbConnjson = $cloudDir + '\Get-InboundConnector.json'
$outbConnpath = $cloudDir + '\Outbound-Connector_EXO.txt'
$outbConnjson = $cloudDir + '\Get-OutboundConnector.json'
$migEndpath = $cloudDir + '\Migration-Endpoints_EXO.txt'
$accDompath = $cloudDir + '\Accepted-Domain_EXO.txt'
$exoacceptedDomjson = $cloudDir + '\Get-AcceptedDomain.json'
$remDomEXOPath = $cloudDir + '\RemoteDomains_EXO.txt'
$exoRemdomjson = $cloudDir + '\Get-RemoteDomain.json'
$onPremOrgpath = $cloudDir + '\Get-OnPremisesOrganization.txt'
$clfederatConfpath = $cloudDir + '\Federation-Configs_EXO.txt'
#$exofederatConfxml = $cloudDir + '\Federation-Trust_EXO.xml'
$exoFedTrustjson = $cloudDir + '\Get-FederationTrust.json'
$clfedorgIdjson = $cloudDir + '\Get-FederatedOrganizationIdentifier.json'
$EntraSpnpath = $cloudDir + '\Entra-SvcPrincipals-OAuth.txt'
$exoSvcPrincjson = $cloudDir + '\Get-MgServicePrincipal-EXO.json'
$exoSvcPrincxml = $cloudDir + '\Get-MgServicePrincipal-EXO.xml'
$spoSvcPrincjson = $cloudDir + '\Get-MgServicePrincipal-SPO.json'
$hybridAppjson = $cloudDir + '\Get-MgServicePrincipal-HybridApp.json'
$hybridAppxml = $cloudDir + '\Get-MgServicePrincipal-HybridApp.xml'
$exoAddpoltxt = $cloudDir + '\EmailAddressPolicies_EXO.txt'
$exoAddpoljson = $cloudDir + '\Get-EmailAddressPolicy.json'
$MigServerTestJson = $cloudDir + '\Test-MigrationServerAvailability_AutoD.json'
$exoOrgReljson =  $cloudDir + '\Get-OrganizationRelationship.json'

#Hybrid Folder creation/validation:
    if (!(Test-Path $outputDir)){
    New-Item -itemtype Directory -Path $outputDir
    }

#Collecting HCW "XHCW" file info:
$hcwPath = Get-ChildItem "$env:APPDATA\Microsoft\Exchange Hybrid Configuration\*.xhcw" -ErrorAction SilentlyContinue | ForEach-Object {Get-Content $_.fullname} -ErrorAction SilentlyContinue
[XML]$hcwLog = "<root>$($hcwPath)</root>"

#Find all 'Set, New, Remove', cmdlets executed by HCW against On-Prem:
function HCWLogs-OnPrem {
if (!(!$hcwPath)){
$title1 = "===='Set-' Commands Executed On-Premises===="
$title1 | Out-File $hcwLogsOPPath
$hcwLog.SelectNodes('//invoke') | Where-Object {$_.cmdlet -like "*Set*" -and $_.type -like "*OnPremises*"} | ForEach-Object {
    New-Object -Type PSObject -Property @{
        'Date'=$_.start;
        'Duration'=$_.duration;
        'Session'=$_.type;
        'Cmdlet'=$_.cmdlet;
        'Comment'=$_.'#comment'
        }
    } | Out-File -Append $hcwLogsOPPath

$title2 = "===='New-' Commands Executed On-Premises===="
$title2 | Out-File -Append $hcwLogsOPPath
$hcwLog.SelectNodes('//invoke') | Where-Object {$_.cmdlet -like "*New*" -and $_.type -like "*OnPremises*"} | ForEach-Object {
    New-Object -Type PSObject -Property @{
        'Date'=$_.start;
        'Duration'=$_.duration;
        'Session'=$_.type;
        'Cmdlet'=$_.cmdlet;
        'Comment'=$_.'#comment'
        }
    } | Out-File -Append $hcwLogsOPPath

$title3 = "===='Remove-' Commands Executed On-Premises===="
$title3 | Out-File -Append $hcwLogsOPPath
$hcwLog.SelectNodes('//invoke') | Where-Object {$_.cmdlet -like "*Remove*" -and $_.type -like "*OnPremises*"} | ForEach-Object {
    New-Object -Type PSObject -Property @{
        'Date'=$_.start;
        'Duration'=$_.duration;
        'Session'=$_.type;
        'Cmdlet'=$_.cmdlet;
        'Comment'=$_.'#comment'
        }
    } | Out-File -Append $hcwLogsOPPath
} else {
    Write-Host -ForegroundColor White "No HCW logs found on this machine..."
}
}
#Find all 'Set, New, Remove', cmdlets executed by HCW against EXO:
function HCWLogs-Cloud {
if (!(!$hcwPath)) {
$title = "===='Set-' Commands Executed in M365===="
$title | Out-File $hcwLogsCloudpath
$hcwLog.SelectNodes('//invoke') | Where-Object {$_.cmdlet -like "*Set*" -and $_.type -like "*Tenant*"} | ForEach-Object {
    New-Object -Type PSObject -Property @{
        'Date'=$_.start;
        'Duration'=$_.duration;
        'Session'=$_.type;
        'Cmdlet'=$_.cmdlet;
        'Comment'=$_.'#comment'
        }
    } | Out-File -Append $hcwLogsCloudpath

$title = "===='New-' Commands Executed in M365===="
$title | Out-File -Append $hcwLogsCloudpath
$hcwLog.SelectNodes('//invoke') | Where-Object {$_.cmdlet -like "*New*" -and $_.type -like "*Tenant*"} | ForEach-Object {
    New-Object -Type PSObject -Property @{
        'Date'=$_.start;
        'Duration'=$_.duration;
        'Session'=$_.type;
        'Cmdlet'=$_.cmdlet;
        'Comment'=$_.'#comment'
        }
    } | Out-File -Append $hcwLogsCloudpath

$title = "===='Remove-' Commands Executed in M365===="
$title | Out-File -Append $hcwLogsCloudpath
$hcwLog.SelectNodes('//invoke') | Where-Object {$_.cmdlet -like "*Remove*" -and $_.type -like "*Tenant*"} | ForEach-Object {
    New-Object -Type PSObject -Property @{
        'Date'=$_.start;
        'Duration'=$_.duration;
        'Session'=$_.type;
        'Cmdlet'=$_.cmdlet;
        'Comment'=$_.'#comment'
        }
    } | Out-File -Append $hcwLogsCloudpath
    }else {
        Write-Host -ForegroundColor Magenta "No HCW logs found on this machine..."
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
function OnPrem-CollectQ{
Write-Host -ForegroundColor Yellow "Do you wish to collect on-premises data? Y/N:"
  $ans = Read-Host
    if ((!$ans) -or ($ans -eq 'y') -or ($ans -eq 'yes')){
        $ans = 'yes'
    #Remote PS Connection Prompt:
    OnPrem-RemoteQ
    #Enter Hybrid server names:
    Write-Host -ForegroundColor Yellow "Enter your Hybrid server names separated by commas (Ex: server1,server2):"
    $script:hybsvrs = (Read-Host).split(",") | foreach {$_.trim()}
    #$hybsvrs | ConvertTo-Json | Out-File $allserversJson

    #Check/Create output folder:
    OnPremDir-Create
    Write-Host -ForegroundColor Cyan "Checking for HCW logs..."
    HCWLogs-OnPrem
    #Collect data:
    OnPrem-Collection
    } else {
        $ans = 'no'
        Write-Host -ForegroundColor Cyan "Skipping on-premises data collection..."
    }
}
#Remote On-Prem Exchange Functions:
function OnPrem-RemoteQ {
    Write-Host -ForegroundColor Yellow "Do you need to create a remote connection to Exchange On-Premises? Y/N:"
  $ans = Read-Host
    if ((!$ans) -or ($ans -eq 'y') -or ($ans -eq 'yes')){
        $ans = 'yes'
        Remote-ExchOnPrem
    } else {Write-Host -ForegroundColor Cyan "Skipping remote powershell connection to on-premises..."}
}
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
    $script:exchCerts = $hybsvrs | foreach {Get-ExchangeCertificate -Server $_}
    $exchCerts | FL | Out-File $exchCertpath
    $exchCerts | ConvertTo-Json | Out-File $exchCertjson
}
#EWS VDir Collection:
function EWS-VdirCollect {
    $script:ewsVdir = $hybsvrs | foreach {Get-WebServicesVirtualDirectory -Server $_ -ADPropertiesOnly}
    $ewsVdir | FL | Out-File $ewstxtPath
    $ewsVdir | ConvertTo-Json | Out-File $ewsVdirjson
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
    $hybtext = "===Hybrid Servers Entered===:"
    $hybtext | Out-File $hybtxtPath
    $hybsvrs | Out-File -Append $hybtxtPath
    $hybsvrs | ConvertTo-Json | Out-File $allserversJson
    Add-Content $hybtxtPath -Value "===Hybrid Configuration===:"
    $hybConf = Get-HybridConfiguration
    $hybConf | FL | Out-File -Append $hybtxtPath
    $hybConf | ConvertTo-Json | Out-File $hybConfigjson
    #$hybConf | Export-Clixml $hybxmlPath
    Start-Sleep -Seconds 2
    
    $OrgConfigtitle = "===On-Premises Organization Config Details===:"
    $OrgConfigtitle | Out-File $opOrgConfigtxt
    $orgConfig = Get-OrganizationConfig
    $orgConfig |FL | Out-File -Append $opOrgConfigtxt
    Start-Sleep -Seconds 2

    $shpol =  "===Sharing Policy Details===:" 
    $shpol | Out-File $opsharPath
    $sharePol = Get-SharingPolicy 
    $sharePol | FL | Out-File -Append $opsharPath
    $sharePol | ConvertTo-Json | Out-File $opsharePolicyjson
    Start-Sleep -Seconds 2
    Add-Content $opsharPath -Value "===Org Relationship Details===:"
    $orgRel = Get-OrganizationRelationship 
    $orgRel |FL | Out-File -Append $opsharPath
    $orgRel | ConvertTo-Json | Out-File $opOrgReljson
    Start-Sleep -Seconds 2

    #Federation Config Info:
    $fedinfotext = "===Federated Organization Identifier===:"
    $fedinfotext | Out-File $opfederatConfpath
    $fedIdent = Get-FederatedOrganizationIdentifier -IncludeExtendedDomainInfo: $false
    $fedIdent | FL | Out-File -Append $opfederatConfpath
    $fedIdent | ConvertTo-Json | Out-File $opfedorgIdjson
    Add-Content $opfederatConfpath -Value "===Federation Information===:"
    $fedInfo = Get-FederationInformation -DomainName $domain -ErrorAction SilentlyContinue
    $fedInfo |FL | Out-File -Append $opfederatConfpath
    Add-Content $opfederatConfpath -Value "===Federation Trust Info===:"
    $fedtrust = Get-FederationTrust
    $fedtrust | ConvertTo-Json | Out-File $opfedTrustjson
    #$fedtrust | Export-Clixml $opfederatConfxml
    $fedtrust | FL Name,Org*certificate,TokenIssuerUri,TokenIssuerEpr,WebRequestorRedirectEpr | Out-File -Append $opfederatConfpath
    $fedtrusttxt = "For additional Federation Trust details, see 'Get-FederationTrust.json' file."
    Add-Content $opfederatConfpath -Value $fedtrusttxt
    Start-Sleep -Seconds 2
    
    #Exch Certs:
    ExchCert-Collection

    #Mail Flow:
    $sendTitle = "===Send Connector Details===:"
    $sendtitle | Out-File $sendConnTxtPath
    $sendConn = Get-SendConnector |? {$_.AddressSpaces -like '*onmicrosoft.com*'}
    $sendConn | ConvertTo-Json | Out-File $sendConnjson
    #$sendConn | Export-Clixml $sendConnxmlPath
    $sendConn | FL | Out-File -Append $sendConnTxtPath
    Start-Sleep -Seconds 2
    $recTitle = "===Receive Connector Details===:"
    $recTitle  | Out-File $recConnTxtPath
    $recvConn = Get-ReceiveConnector |?{$_.TlsDomainCapabilities -like '*outlook*'}
    #$recvConn | Export-Clixml $recConnxmlPath
    $recvConn | ConvertTo-Json | Out-File $recConnjson
    $recvConn |FL | Out-File -Append $recConnTxtPath
    $accDomain = Get-AcceptedDomain
    $accDomain | ConvertTo-Json | Out-File $OpacceptedDomJson
    #Remote Domains:
    $remText = "===Remote Domains===:"
    $remText | Out-File $opRemDomtxt
    $remDom = Get-RemoteDomain
    $remDom | FL | Out-File -Append $opRemDomtxt
    $remDom | ConvertTo-Json | Out-File $opremDomjson
    #Email Address Policies:
    $addpoltext = "===On-Premises Email Address Policies===:"
    $addpoltext | Out-File $OPaddpoltxt
    $addpolOP = Get-EmailAddressPolicy
    $addpolOP | FL | Out-File -Append $OPaddpoltxt
    $addpolOP | convertto-json | Out-File $opAddpoljson
    Start-Sleep -Seconds 2
    
    #EWS VDir Collect function:
    EWS-VdirCollect
    
    #OAuth Config Details:
    $iOrgConn = Get-IntraOrganizationConnector
    if (!$iOrgConn){
        $iocFailtext = "No IntraOrg Connector detected, OAuth may not be configured..."
        Write-Host -ForegroundColor Cyan $iocFailtext
        }
    $iorgtext = "===IntraOrg Connector===:"
    $iorgtext | Out-File $opOauthPath
    if ($iocFailtext -ne $null) {
        $iocFailtext | Out-File -Append $opOauthPath
    }
    $iOrgConn | FL | Out-File -Append $opOauthPath
    Add-Content $OPoauthPath -Value "===IntraOrganization Configs===:"
    $iOrgConf = Get-IntraOrganizationConfiguration -WarningAction:SilentlyContinue
    $iOrgConf | FL | Out-File -Append $OPoauthPath
    $iOrgConf | ConvertTo-Json | Out-File $opIOConnjson
    Add-Content $OPoauthPath -Value "===Partner Application Details===:"
    $ptnrapp = Get-PartnerApplication
    $ptnrapp | ConvertTo-Json | Out-File $partnerAppjson 
    $ptnrapp |FL Name,Enabled,Applicationidentifier,UseAuthServer,LinkedAccount| Out-File -Append $OPoauthPath
    #$ptnrapp | Export-Clixml $partnerappxml
    Add-Content $opOauthPath -Value "===Auth Server Settings===:"
    $authsvr = Get-AuthServer
    $authsvr | ConvertTo-Json | Out-File $authSvrjson
    $authsvr | FL Name,type,realm,enabled,TokenIssuingEndpoint,AuthorizationEndpoint,IsDefaultAuthorizationEndpoint | Out-File -Append $opOauthPath
    Add-Content $OPoauthPath -Value "**Additional Auth Server details found in Json file."
    #$authsvr | Export-Clixml $authsvrxmlPath
    $authConftext = "===On-Premises Auth Config===:"
    $authConftext | Out-File $authConfigpath
    $authConf = Get-AuthConfig
    $authConf | Out-File -Append $authConfigpath

    #Test OAuth Config: 
    function OAuth-Test-OP {
    Write-Host -ForegroundColor Yellow "Would you like to test OAuth on-prem? Y/N:"
        $ans = Read-Host
        if ((!$ans) -or ($ans -eq 'y') -or ($ans -eq 'yes')){
            $ans = 'yes'
            Test-OAuthConnectivity -Service EWS -TargetUri https://outlook.office365.com/ews/exchange.asmx -Mailbox $testMbx -Verbose 
        } else {$ans = 'no'
        Write-Host -ForegroundColor White "Skipping OAuth Test..."
            }
    }    
       
    #Skype Integration Details:
    $skypText = "===Skype On-prem Integration Details===:"
    $skypText | Out-File $skypeIntTxt
    $sfbUser = Get-MailUser Sfb* -ErrorAction SilentlyContinue
    $sfbUser | FL | Out-File -Append $skypeIntTxt
    $userAppRole = Get-ManagementRoleAssignment -Role UserApplication -GetEffectiveUsers |? {$_.EffectiveUserName -like 'Exchange*'}
    $archiveAppRole = Get-ManagementRoleAssignment -Role ArchiveApplication -GetEffectiveUsers |? {$_.EffectiveUserName -like 'Exchange*'}
    $userAppRole |FL Role, *User*, WhenCreated | Out-File -Append $skypeIntTxt
    $archiveAppRole |FL Role, *User*, WhenCreated | Out-File -Append $skypeIntTxt

    #Close remote connection:
    Write-Host -ForegroundColor White "Closing connection to Exchange..."
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
        HCWLogs-Cloud
        #Connect to EXO & Collect Data:
        Remote-EXOPS
        #Collect Entra Data:
        AAD-Collection
    } else {
            $ans = 'no'
            Write-Host -ForegroundColor Cyan "Skipping M365 data collection..."
        }
}
#Remote EXO PS Function:
function Remote-EXOPS {
    Write-Host -ForegroundColor Cyan "Connecting to Exchange Online..."
    Write-Host -ForegroundColor Yellow "Enter your M365 admin username (Ex: admin@yourdomain.onmicrosoft.com):"
    $exoUPN = Read-Host
    try {
        Import-Module ExchangeOnlineManagement
        Connect-ExchangeOnline -UserPrincipalName $exoUPN -ShowBanner:$false
    } catch {
        $exoV3Fail = "EXO Remote PS Connection Failed"
        }
    if ($null -ne $exoV3Fail) {
        Write-Host -ForegroundColor Cyan "Remote EXO connection failed." 
        Write-Host -ForegroundColor Cyan "Ensure that the EXO V3 module is installed (https://aka.ms/exops-docs) and that basic auth is disabled in your tenant." 
    }
    else {
        EXO-Collection
    }
}
function Test-MigrationEndPtQ {
    Write-Host -ForegroundColor Yellow "Do you wish to test connectivity to your on-prem the migration endpoint? Y/N:"
    $ans = Read-Host
      if ((!$ans) -or ($ans -eq 'y') -or ($ans -eq 'yes')){
          $ans = 'yes'
          MigEndpoint-Test
      } else {
          $ans = 'no'
          Write-Host -ForegroundColor Cyan "Skipping migration endpoint test..."
      }
}
#Migration Server Test:
function MigEndpoint-Test {
    if ($null -eq $opCreds){
        Write-Host -ForegroundColor Yellow "Enter your on-premises Exchange admin credentials:"
        $opCreds = Get-Credential
    }    
    Write-Host -ForegroundColor Yellow "Enter the email address of an on-premises mailbox to test migration server availability:"
    $opMbx = Read-Host
    try {
       $migsvctest = Test-MigrationServerAvailability -ExchangeRemoteMove -Autodiscover -EmailAddress $opMbx -Credentials $opCreds
       $migsvctest | ConvertTo-Json | Out-File $MigServerTestJson
    }
    catch {
        $errorMsg = "Test command failed, please run 'Test-MigrationServerAvailability' manually if needed..."
        Write-Error -Message $errorMsg
        $errorMsg | Out-File $errorLog
    } 
}

#Service Principal Collection for OAuth:
function AAD-Collection {
Write-Host -ForegroundColor Cyan "Connecting to Entra ID..."
    try {
        #Install-Module Microsoft.Graph -Scope CurrentUser (if not installed, this will install the MS Graph module)
        Connect-MgGraph -Scopes "Application.Read.All" -NoWelcome
    }
    catch {
        $GraphFail = "Connection to Entra ID Failed. Ensure MgGraph Powershell Module is installed and run the cmdlet below manually to collect OAuth service principal info:"
        $AADSvcPrinCmdlet = "Get-MgServicePrincipal -Filter 'AppId eq '00000002-0000-0ff1-ce00-000000000000''"
    }
    if ($null -ne $GraphFail) {
        Write-Host -ForegroundColor Yellow $GraphFail
        Write-Host -ForegroundColor White $AADSvcPrinCmdlet
        Write-Host -ForegroundColor Cyan "Refer to: https://learn.microsoft.com/en-us/powershell/microsoftgraph/installation?view=graph-powershell-1.0 for additional details on MS Graph module."
    } else {
        $svcPrinText = "=== Entra ID Service Principals for OAuth ==="
        $svcPrinText | Out-File -Append $EntraSpnpath
        $exoSvcId = '00000002-0000-0ff1-ce00-000000000000'
        $skypeSvcId = '00000004-0000-0ff1-ce00-000000000000'
        $exosvcPrinc = Get-MgServicePrincipal -Filter "AppId eq '$exoSvcid'"
        $skypsvcPrinc = Get-MgServicePrincipal -Filter "AppId eq '$skypeSvcid'"
        $hybridapp = Get-MgServicePrincipal -Filter "startsWith(DisplayName, 'ExchangeServerApp')"
        $exosvcPrinc | ConvertTo-Json | Out-File $exoSvcPrincjson
        $exosvcPrinc | Export-Clixml $exoSvcPrincxml
        $skypsvcPrinc | ConvertTo-Json | Out-File $spoSvcPrincjson
        $exosvcPrinc | FL AppDisplayName,ObjectType,AccountEnabled,AppId | Out-File -Append $EntraSpnpath

        Add-Content $EntraSpnpath -Value "Registered 'ServicePrincipalNames':"
        Add-Content $EntraSpnpath -Value ""
        $exosvcPrinc | Select -ExpandProperty ServicePrincipalNames | Out-File -Append $EntraSpnpath
        Add-Content $EntraSpnpath -Value ""

        Add-Content $EntraSpnpath -Value "===Hybrid Application Info===:"
        $hybridapp | FL DisplayName,Appid,Description,Notes | Out-File -Append $EntraSpnpath
    }
}

#EXO Data Collection:
function EXO-Collection {
Write-Host -ForegroundColor Cyan "Collecting data from Exchange Online..."
$shpol = "===Sharing Policy Details===:"
$shpol | Out-File $clSharingPath
$sharePol = Get-SharingPolicy 
$sharePol |FL | Out-File -Append $clSharingPath
$sharePol | ConvertTo-Json | Out-File $exosharingpoljson
Add-Content $clSharingPath -Value "===Org Relationship Details===:"
$orgRel = Get-OrganizationRelationship 
$orgRel |FL | Out-File -Append $clSharingPath
$orgRel | ConvertTo-Json | Out-File $exoOrgReljson

#Fed Org Info:
$fedoiText = "===Federated Organization Information==="
$fedoiText | Out-File $clfederatConfpath
$fedOI = Get-FederatedOrganizationIdentifier
$fedOI | FL | Out-File -Append $clfederatConfpath
$fedOI | ConvertTo-Json | Out-File $clfedorgIdjson
Start-Sleep -Seconds 2
$fedtrusttext = "===Federation Trust Info===:"
$fedtrusttext | Out-File -Append $clfederatConfpath
$fedtrustexo = Get-FederationTrust
#$fedtrustexo | Export-Clixml $exofederatConfxml
$fedtrustexo | ConvertTo-Json | Out-File $exoFedTrustjson
$fedtrustexo |FL Name,TokenIssuer*,WebRequestorRedirectEpr | Out-File -Append $clfederatConfpath
$fedtrusttext = "For full Federation Trust details, see 'Get-FederationTrust.json' file."
$fedtrusttext | Out-File -Append $clfederatConfpath

$cloudOrgtext = "===Organization Config Details===:"
$cloudOrgtext | Out-File $cloudOrgpath
$orgConfig = Get-OrganizationConfig 
$orgConfig |FL | Out-File -Append $cloudOrgpath

$opOrgtext = "===On-Premises Organization==="
$opOrgtext | Out-File $onPremOrgpath
$opOrg = Get-OnPremisesOrganization
$opOrg | FL | Out-File -Append $onPremOrgpath

$migEndtext = "===Migration Endpoints===:"
$migEndtext | Out-File $migEndpath
$migEndp = Get-MigrationEndpoint
$migEndp | FL | Out-File -Append $migEndpath

#OAuth Configs:
$iOrgText = "===IntraOrg Connector===:"
$iOrgText | Out-File $clOauthpath
$CliOrgConn = Get-IntraOrganizationConnector
if ($null -ne $CliOrgConn) {
    $CliOrgConn | FL | Out-File -Append $clOauthpath
    $CliOrgConn | ConvertTo-Json | Out-File $clIOConnjson
    #$CliOrgConn |Export-Clixml $clIntraOrgxmlpath

} else {
    $iOrgText2 = "***No OAuth Configs found***"
    $iOrgText2 | Out-File -Append $clOauthpath
}
Start-Sleep -Seconds 2

#Mail Flow Configs:
$outbCtext = "===O365 Outbound Connector Details===:"
$outbCtext | Out-File $outbConnpath
$o365OutConn = Get-OutboundConnector -IncludeTestModeConnectors $true
$o365OutConn | FL | Out-File -Append $outbConnpath
$o365OutConn | ConvertTo-Json | Out-File $outbConnjson 
$inbCtext = "===O365 Inbound Connector Details===:"
$inbCtext | Out-File $inbConnpath
$o365InConn = Get-InboundConnector 
$o365InConn |FL | Out-File -Append $inbConnpath
$o365InConn | ConvertTo-Json | Out-File $inbConnjson

$accDtext = "===Accepted Domain===:"
$accDtext | Out-File $accDompath
$accDomain = Get-AcceptedDomain $domain
$accDomain | FL | Out-File -Append $accDompath
$accDomain | ConvertTo-Json | Out-File $exoacceptedDomjson

$remDexotext = '===EXO Remote Domains==='
$remDexotext | Out-File $remDomEXOPath
$remDomEXO = Get-RemoteDomain
$remDomEXO | FL | Out-File -Append $remDomEXOPath
$remDomEXO | ConvertTo-Json | Out-File $exoRemdomjson

$addpoltext = "===EXO Email Address Policies===:"
$addpoltext | Out-File $ExOaddpoltxt
$addpolexo = Get-EmailAddressPolicy
$addpolexo | FL | Out-File -Append $ExOaddpoltxt

#Test Migration Endpoint:
Test-MigrationEndPtQ

Write-Host -ForegroundColor White "Closing connection to Exchange Online..."
Get-PSSession | Remove-PSSession
}

#Prompt for domain name:
Write-Host -ForegroundColor Yellow "Enter your primary domain name:"
$domain = Read-Host

#On-prem Collection Prompt:
$ans = OnPrem-CollectQ

#EXO Collection Prompt:
$ans = EXO-RemoteQ

#Goodbye:
Write-Host -ForegroundColor White "Collection complete. Review or submit any files located in '$outputDir' to Microsoft support."

#Revert execution policy if needed:
if ($execPol -ne 'Unrestricted'){
    Write-Host -ForegroundColor Cyan "Changing execution policy back to '$execPol'..."
    Set-ExecutionPolicy $execPol -Force
}
