# Hybrid Configuration Collection Script
Gathers data related to Exchange Hybrid mailflow, free/busy sharing & OAuth.

 The script will first prompt you on whether you wish to collect on-premises Exchange data. Once complete (or if you answer 'no'), it will ask whether you wish to collect Exchange Online data. Once complete, it will display the location of collection data. Some files will be in '.xml' format and some in text files. In some cases, data is sent to both formats for flexibility. 
 
 When connecting to Exchange Online, the script will attempt to import the EXO V2 module; if this fails it will fallback to the legacy, Basic Auth remote method.
