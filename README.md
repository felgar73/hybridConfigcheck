# Hybrid Configuration Collection Script
 -Script assumes Kerberos Auth is enabled on-prem for remote Exchange session.
    -Supports EXO V2 Powershell module.
    -Script can be run from any domain-joined machine, but it's recommended to run from Exchange Mgmt Shell directly and bypass remote powershell connection.
    -Details collected for Exchange certificates will be limited due to remote powershell limitations. 

    The script will first prompt you on whether you wish to collect on-premises Exchange data and whether or not a remote powershell session is needed for this (if running from local Exchange server, simply answer to this question). 
    Once On-prem collection is complete (or if you answer 'no'), it will ask whether you wish to collect Exchange Online data. 
    Once complete, it will display the location of collection data. Some files will be in '.xml' format and some in text files. In some cases, data is sent to both formats for flexibility.

    When connecting to Exchange Online, the script will attempt to import the EXO V2 powershell module; if this fails it will fallback to the legacy, Basic Auth remote method.
    When connecting to Azure AD in order to collect OAuth some settings, the script will attempt to import the AzureADPreview powershell module; if this fails it will continue with other operations and this will require manual collection.

Example: ./hybridconfigcollection.ps1

Once all collection is complete, the folder containing collected data will be displayed. Some files will be in '.xml' format and some in text files. In some cases, data is sent to both formats for flexibility.

If you wish to save previous output, make sure to rename output folders before re-running the script or files will be overwritten.
