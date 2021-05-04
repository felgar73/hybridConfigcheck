# Hybrid Configuration Collection Script
    -Supports EXO V2 Powershell module.
    -Script can be run from any domain-joined machine, but I recommended running from Exchange Mgmt Shell directly and bypass remote powershell connection.
    -If run from remote machine, Exchange certificate details will be limited due to remote powershell limitations. 
 
The script will first prompt you on whether you wish to collect on-premises Exchange data and whether or not a remote powershell session is needed for this (if running from local Exchange server, simply answer to this question). Once On-prem collection is complete (or if you answer 'no'), it will ask whether you wish to collect Exchange Online data. 

    -When connecting to Exchange Online, the script will attempt to import the EXO V2 powershell module; if this fails it will fallback to the legacy, Basic Auth remote method.
    -When connecting to Azure AD in order to collect OAuth some settings, the script will attempt to import the AzureADPreview powershell module; if this fails it will continue with other operations and this will require manual collection.

Example: ./hybridconfigcollection.ps1

When connecting to Exchange Online, the script will attempt to import the EXO V2 module; if this fails it will fallback to the legacy, Basic Auth remote method.

Once all collection is complete, the folder containing collected data will be displayed. Some files will be in '.xml' format and some in text files. In some cases, data is sent to both formats for flexibility.

If you wish to save previous output, make sure to rename output folders before re-running the script or files will be overwritten.
**************************************************
Requirements: 
-Powershell should to be 'Run As Administrator'.
-Script assumes Kerberos Auth is enabled on-prem.
-Supports EXO V2 Powershell module
-Details collected for Exchange certificates will be limited due to remote powershell limitations.
-If more Exch cert data is needed go ahead and run the 'Get-ExchangeCertificate' cmdlet locally on the server.
