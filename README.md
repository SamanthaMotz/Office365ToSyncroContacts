# Office 365 to Snycro Contact Importer

The primary purpose of this script is to take all licensed used in an Office 365 tenant and import then into Syncro under a specific customer account. It does not include any code to delete or modify data already in Snycro, only add to it. 

Sytax Example:
Office365ToSnycro.ps1 -subdomain Companyname.shield -API 82dgb45kjgnfrq98asdf-9adfaa9045309g09ad094 -CustID 587314

# Main Features:

**Data Collection**: Automatic collection and sorting of all licensed Office 365 users on a tenant.

**Customer Confirmation**:  Ensures you are importing into the correct customer. This can be skipped if needed. 

**Duplication Checking**: Checks for and skips over duplicate contacts. This may be triggered by names that are very similar. I.E. Bob Smith and Bob Smithson

**Robust Error Checking**: While it won’t catch everything, it will catch the most common ones and make suggestions on how to fix them. 

**Data Segregation**: Separates all run files and logs into separate sub folders, to prevent cross customer contamination

**Customer Indexing**: to easy find customer IDs in the future. 

**Verbose Logging**: This tracks and timestamps start time, end time, duplicates found, the number of API calls made, and errors reported. Logs are stored in C:\Temp\O365ToSyncro\CustomerID# by default. This can be adjusted. An index file is created in the main folder to help find the correct subfolders. 

**API Rate Limiting**: Prevents the script from hitting the 180 per minute API limit in the case of large uploads. In this case the script may appear to hang for 60 seconds. Do not cancel the script mid run or it may not clean up files properly and could compromise future runs


# Requirements to run properly. 
This script requires Curl to run properly. Curl is installed by default on Windows 10 1803 and above.

Microsoft Azure Active Directory Module for Windows must be installed for this to work properly. If it is not, see the following site for information on how to install it. This script will NOT install anything on your system

https://docs.microsoft.com/en-us/microsoft-365/enterprise/connect-to-microsoft-365-powershell?view=o365-worldwide#connect-with-the-microsoft-azure-active-directory-module-for-windows-powershell

This script has 3 main parameters, they are all required.

**-Subdomain** is your Syncro Subdomain. Something like Companyname or companyname.shield. Do not include your full Syncro URL. Just the subdomain without the trailing period.

**-CustID** is your customers ID number in Syncro. When viewing their profile, it's at the end of the URL. It's all numbers.

**-API** is your API Key. It must have at least the following permissions. Additional permissions are not recommended.
Customers - View Detail
Customers – Edit 

There is a 4th optional switch you can use, but it is not recommended. 

**-noconfirm** will skip client confirmation. Be VERY careful or you could import contacts into the wrong client!

# Known issues:

Error checking on the 2 critical Office 365 commands is not yet implemented. If these fail the script will continue but won’t upload any data. It’s best to let it finish if this happens so it can cleanup properly. 

There may be API errors that are not recognized. If you run into any API errors that are reported as unknown, you may post the error line in the Issues section on [GitHub](https://github.com/SamanthaMotz/Office365ToSnycroContacts). If possible the script will be updated to handle these new errors.

The duplicate checking is limited to 100 contacts. This may be increased if users request it. 




