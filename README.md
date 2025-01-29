# Hybrid Configuration Collection Script
- Script assumes Kerberos Auth is enabled on-prem for remote Exchange session.  
- Supports EXO V3 Powershell module.  
- Script can be run from any domain-joined machine, but it's recommended to run from Exchange Mgmt Shell directly and bypass remote powershell connection.  
- Details collected for Exchange certificates will be limited due to remote powershell limitations.  

The script will first prompt you on whether you wish to collect on-premises Exchange data and whether or not a remote powershell session is needed for this (if running from local Exchange server, simply answer to this question). 
Once On-prem collection is complete (or if you answer 'no'), it will ask whether you wish to collect Exchange Online data. 
Once complete, it will display the location of collection data. Some files will be in '.xml' format and some in text files. In some cases, data is sent to both formats for flexibility.

When connecting to Exchange Online, the script will attempt to import the EXO V3 powershell module; if this fails it will fallback to the legacy, Basic Auth remote method.
When connecting to Azure AD in order to collect OAuth some settings, the script will attempt to import the AzureADPreview powershell module; if this fails it will continue with other operations and this will require manual collection.

Example: `.\hybridconfigcollection.ps1`

Once all collection is complete, the folder containing collected data will be displayed. Some files will be in '.xml' format and some in text files. In some cases, data is sent to both formats for flexibility.

If you wish to save previous output, make sure to rename output folders before re-running the script or files will be overwritten.

### Updates
**Aug 2020**  
- Modified output file creation  

**Sept 2020**  
- Added HCW log parsing for 'Set-' cmdlets  
- Added EWS Collection function to search for specific servers  

**Oct 2020**  
- Added Exch certificate collection function  
- Added 'New-' cmdlet search within HCW logs  

**Nov 2020**  
- Modified folder creation 'Test-Path' functions  

**Dec 2020**  
- Added MSOL/AAD data collection for OAuth configs  

**Jan 2021**
- Added HCW log parsing for 'Remove-' cmdlets
- Added additional 'Get-' commands to be pulled

**Feb 2021**
- Added 'Get-FederatedOrganizationIdentifier' collection in EXO

**April 2021**
- Added option to bypass remote session to on-prem Exchange in the case you are already logged into an Exchange server.
- Added 'Get-EmailAddressPolicy' to data collection functions.

**May 2022**
- Modified output for Partner Applicaiton details:
- truncated output within OAuth Configs text file
- added XML file for full output
- Adding collection of Skype Integration configs

**July 2022**
- Added Silent Error action for Skype config & Federation checks

**April 2023**
- Removed Basic Auth as an optional login method for EXO powershell
- Renamed variable used for EXO v3 PS failure
- Removed mentions of v2 module and and replaced with v3

**Nov 2023**
- Now collecting ALL EXO mailflow connectors including "test mode" connectors

**Jan 2024**
- Added Json output commands to replace xml output
- Replaced MSO commands with MS Graph for Entra ID/OAuth data collection

 **Dec 2024**
 - Moved EXO-Collection function to avoid creating empty output files if PS connection fails.
   
**Jan 2025**
- Added additional Json output files.
- Added Test-MigrationServerAvailability function for on-prem migration endpoint testing.
