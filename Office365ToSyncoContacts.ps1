param ($subdomain,
	$API,
	$CustID,
	[switch]$noconfirm)

#User editable vars
$TempRoot = 'C:\Temp\O365ToSyncro\'

#Editing these vars may break stuff
$Temp = "$TempRoot$CustID\"
$MainLog = ("$subdomain" + "-CustomerID" + "$CustID" + "-Office365ToSyncro.log")
$CurlPresent = Curl.exe -V
$Office365ToSnycroLog = "$Temp$MainLog"
$GetContactLog = '$TempGetContact-Combined.log'
$UserCSV = '$TempLicensedUsers.csv'
$UserCSVTrim = '$TempLicensedUsersTrim.csv'
$CustIndex = "$TempRoot\CustomerIndex.txt"
$APICounter = [int]0
$UploadCounter = [int]0
$DupeCounter = [int]0
$TimeStamp = (Get-Date -Format HH:mm:ss.fff)
$DateStamp = (Get-Date -Format yyyy-MM-dd)
$APILimit = [int]170

function APILimiter ()
{
	if ($APICounter -gt $APILimit)
	{
		Start-Sleep -Seconds 60
		$APILimit = ($APICounter + $APILimit)
	}
	
}

function BuildBat ()
{
	param ($Type)
	
	$BatchLog = ("$Temp" + $Type + ".log")
	write "$DateStamp T$TimeStamp" | Out-File $BatchLog -Append -Encoding Ascii -NoNewline
	$BatchLog = ">> $BatchLog"
	$CurlGetArgs = ("curl.exe" + " " + $Curlarguments + [char]34 + " " + $BatchLog)
	$BatName = ($Type + ".bat")
	write $CurlGetArgs | Out-file "$Temp$Type.bat" -Encoding Ascii
	Start-Process "$Temp$Type.bat" -Wait -WindowStyle Hidden
	$APICounter++
	APILimiter
	$Log = ("$Temp" + $Type + ".log")
	Remove-Item -LiteralPath "$Temp$Type.bat" -Force -ErrorAction SilentlyContinue
	return $Log
}

function CleanData ()
{
	param ($SplitMarker,
		$Type)
	
	
	$WorkingFile = ("$Temp" + "$Type" + "-Working.log")
	$Type = ("$Temp" + "$Type" + ".log")
	Get-Content $Type -Tail 1 | Out-File $WorkingFile -Force
	$TypeLog = ("$Type" + "Parse")
	$TypeLog = (Get-Content $WorkingFile)
	$TypeLog = ([string]$TypeLog)
	$TypeLog = ($TypeLog -split "$SplitMarker")
	CleanUp -LogName $WorkingFile
	return $TypeLog
}

function CleanUp ()
{
	param ($LogName)
	
	$LogExtension = (write $LogName | select-string -pattern ".log")
	if (!$LogExtension)
	{
		$LogName = ("$Logname" + ".log")
	}
	
	$LogPath = (write $LogName | select-string -pattern 'Temp')
	
	if (!$LogPath)
	{
		$LogName = ("$Temp" + "$Logname")
	}
	Remove-Item -LiteralPath $Logname -Force -ErrorAction SilentlyContinue
}

function ErrorCheck()
{
	param ($LogName)
	
	$LogName = ("$Temp" + "$LogName" + ".Log")
	
	$AuthError = (Get-Content $LogName | select-string -pattern "Failed to authenticate account" | select -last 1)
	$PermError = (Get-Content $LogName | select-string -pattern "Not authorized" | select -last 1)
	$CustMissing = (Get-Content $LogName | select-string -pattern "Not found" | select -last 1)
	$SecurityError = (Get-Content $LogName | select-string -pattern "0x80090322" | select -last 1)
	
	if ($AuthError)
	{
		Get-Content $LogName | Out-File $Office365ToSnycroLog -Append -Encoding Ascii
		write "Account authentication error reported. Probably a bad API key. Exact Error codes follows. $AuthError"
		
	}
	
	if ($PermError)
	{
		Get-Content $LogName | Out-File $Office365ToSnycroLog -Append -Encoding Ascii
		write "Permission error reported. API Good but likly wrong permissions. Exact Error codes follows. $PermError"
		
	}
	If ($CustMissing)
	{
		Get-Content $LogName | Out-File $Office365ToSnycroLog -Append -Encoding Ascii
		write "Customer not found, check ID Number. Exact Error codes follows. $CustMissing"
	}
	
	If ($SecurityError)
	{
		Get-Content $LogName | Out-File $Office365ToSnycroLog -Append -Encoding Ascii
		write "Security error. Probably a mistyped subdomain. Exact Error codes follows. $SecurityError"
	}
	
	
	if ((!$AuthError) -and (!$PermError) -and (!$CustMissing) -and (!$SecurityError))
	{
		Get-Content $LogName | Out-File $Office365ToSnycroLog -Append -Encoding Ascii
		write "Unknown error. Check $Office365ToSnycroLog"
		
	}
	write "$DateStamp T$TimeStamp WARNING: Ending Run due to errors during Curl process" | Out-File $Office365ToSnycroLog -Append -Encoding Ascii
	FinalCleanUp
	exit
}

function FinalCleanUp()
{
	CleanUp -LogName $PostContact
	CleanUp -LogName $GetLog
	CleanUp -LogName $GetLogWork
	CleanUp -LogName $GetContactLog
	CleanUp -LogName $GetContact
	CleanUp -LogName $UserCSV
	CleanUp -LogName $GetCust
}

function IndexCust ()
{
	if ($CustBName)
	{
		
		write "$CustBName --> $CustID" | Out-File ($CustIndex + "Working") -Append -Encoding Ascii
	}
	Else
	{
		write "$CustName--> $CustID" | Out-File ($CustIndex + "Working") -Append -Encoding Ascii
	}
	
	gc ($CustIndex + "Working") | sort | Get-Unique > $CustIndex
	Remove-Item ($CustIndex + "Working") -Force
	
	
}

mkdir  $Temp -ErrorAction SilentlyContinue > $null

write "$DateStamp T$TimeStamp Starting Run" | Out-File $Office365ToSnycroLog -Append -Encoding Ascii

#Checks to make sure all required params are entered and responds with help info if not. 
if ((!$subdomain) -or (!$CustID) -or (!$API))
{
	write "Parameters are missing"
	write "Subdomain is $subdomain"
	write "CustID is $CustID"
	write "API is $API"
	write ""
	Write-Host "This program imports all licensed users in an Office 365 enviroment and imports them into a Syncro Customer"
	write ""
	Write-Host "There are three parameters needed to run, they are all required"
	Write-Host "-Subdomain is your custom subdomain in Syncro. (I.E. YourName or YourName.Shield) do not include the trailing period"
	Write-Host "-API is your API Key. The following permissions are required are Customers - View Detail and Customers - Edit"
	Write-Host "-CustID is your customers ID number in Syncro. When viewing their profile, it's at the end of the URL. It's all numbers"
	Write-Host "-noconfirm is a switch that skips client confirmation. WARNING: Be VERY careful or you could import contacts into the wrong client!"
	write ""
	Write-Host "EXAMPLE SYTAX Office365ToSnycro.ps1 -subdomain CompanyName.shield -API 49dgb45kjgnfrqweg98asdf-9adfaa9045309g09ad093434 -CustID 587314"
	write ""
	Write-Host "This is a BETA version. Error checking is present but may not catch all errors. Use at your own risk."
	write "$DateStamp T$TimeStamp WARNING: Ending Run due to missing parameters" | Out-File $Office365ToSnycroLog -Append -Encoding Ascii
	FinalCleanUp
	exit
}

#Checks to make sure Curl is present
If (!$CurlPresent -Match "Protocols")
{
	write 'Curl is not installed and is required for this script. Please install Curl and try again'
	write "$DateStamp T$TimeStamp WARNING: Ending Run due to Curl missing" | Out-File $Office365ToSnycroLog -Append -Encoding Ascii
	FinalCleanUp
	exit
	
}
Write "This is a BETA version. Error checking is present but may not catch all errors. Use at your own risk."
write ""
write "Syncro has an API Limit of 180 items per minute. API counting is included as part of this script."
write "If this script makes more then $APILimit API calls, it will pause for 60 seconds to prevent it from hitting the API limit and being blocked."
write "Do not cancel the script mid run or it may not clean up files properly and could compromise future runs"
write ""


#Build and run the batch file for to collect customer info
$Curlarguments = "-X GET ""https://$subdomain.syncromsp.com/api/v1/customers/$CustID"" -H ""accept: application/json"" -H ""Authorization: $API"

$GetCust = (BuildBat -Type GetCust)
$APICounter++
APILimiter

#Process customer data
$GetCustLog = CleanData -SplitMarker "," -Type GetCust
$CustName = (write $GetCustLog | select-string -pattern "fullname" | select -last 1)
$CustBName = (write $GetCustLog | select-string -pattern "business_name" | select -last 1)

if (($null -eq $CustName) -or ($null -eq $CustBNAme))
{
	ErrorCheck -LogName GetCust
	
}

$CustName = ([string]$CustName)
$CustName = ($CustName.replace('"fullname":"', '').Replace('"', ''))

$CustBName = ([string]$CustBName)
$CustBName = ($CustBName.replace('"business_name":', '').Replace('"', ''))

#Either prompt console to check names or skip if noconfirm switch was used. 
if ($noconfirm.IsPresent)
{
	write 'WARNING: Skipping client confirmation'
	IndexCust
}

Else
{
	write ""
	write "WARNING: Please confirm the below information is the for right customer. Otherwise contacts will be imported into the wrong customer!"
	write "Name: $CustName"
	write "Business Name: $CustBName"
	
	write "Do you wish to continue? Press y and enter for yes or just enter to cancel."
	write ""
	$userInput = Read-Host
	
	
	if (($userInput -eq "y") -or ($userInput -eq "yes"))
	{
		write 'Correct client confirmed, continuing'
		write ''
		IndexCust
	}
	else
	{
		write 'Existing client information was not confimed. No data has been upload to Syncro. Exiting'
		write "Total API calls for this run: $APICounter"
		write "$DateStamp T$TimeStamp WARNING: Ending Run due to client not being confirmed" | Out-File $Office365ToSnycroLog -Append -Encoding Ascii
		IndexCust
		FinalCleanup
		exit
	}
	
}


#connect to Office 365 to collect user list. Commented out for testing
Connect-MsolService

Get-MSOLUser | Where-Object { $_.isLicensed -eq "True" } | Select-Object FirstName, LastName, UserPrincipalName | Export-Csv $UserCSV

Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue

$UserCSVPresent = (Test-Path $UserCSV)
If (!$UserCSVPresent)
{
	write "User list is missing from $UserCSV. Cannot continue, exiting"
	write "$DateStamp T$TimeStamp WARNING: Ending Run due to missing user list" | Out-File $Office365ToSnycroLog -Append -Encoding Ascii
	FinalCleanUp
	exit
}

#Format data
$Table = Import-Csv -Path $UserCSV
$Table = ($Table | Format-Table)


#Collect and collate up to 100 existing Syncro Contacts
$Type = 'GetContact'
$Curlarguments = "-X GET ""https://$subdomain.syncromsp.com/api/v1/contacts?customer_id=$CustID&page=1"" -H ""accept: application/json"" -H ""Authorization: $API"
$Log1 = (BuildBat -Type GetContact)
$APICounter++
APILimiter


$Curlarguments = "-X GET ""https://$subdomain.syncromsp.com/api/v1/contacts?customer_id=$CustID&page=2"" -H ""accept: application/json"" -H ""Authorization: $API"
$Log2 = (BuildBat -Type GetContact2)
$APICounter++
APILimiter


$Curlarguments = "-X GET ""https://$subdomain.syncromsp.com/api/v1/contacts?customer_id=$CustID&page=3"" -H ""accept: application/json"" -H ""Authorization: $API"
$Log3 = (BuildBat -Type GetContact3)
$APICounter++
APILimiter

$Curlarguments = "-X GET ""https://$subdomain.syncromsp.com/api/v1/contacts?customer_id=$CustID&page=4"" -H ""accept: application/json"" -H ""Authorization: $API"
$Log4 = (BuildBat -Type GetContact4)
$APICounter++
APILimiter


$Log1Data = Get-Content $Log1
$Log2Data = Get-Content $Log2
$Log3Data = Get-Content $Log3
$Log4Data = Get-Content $Log4

Write ($Log1Data + $Log2Data + $Log3Data + $Log4Data) | Out-file "$GetContactLog" -Encoding Ascii
Remove-Item -LiteralPath $Log1 -ErrorAction SilentlyContinue
Remove-Item -LiteralPath $Log2 -ErrorAction SilentlyContinue
Remove-Item -LiteralPath $Log3 -ErrorAction SilentlyContinue
Remove-Item -LiteralPath $Log4 -ErrorAction SilentlyContinue


#For each item in CSV create a batch file containing the required Curl command, trim that users line from the list, run the batch file, repeat.  
foreach ($line in Get-Content $UserCSV)
{
	$FirstName = Import-Csv -Path "$UserCSV" | Select-Object -First 1 | Select -Last 1 -ExpandProperty 'FirstName'
	$LastName = Import-Csv -Path "$UserCSV" | Select-Object -First 1 | Select -Last 1 -ExpandProperty 'LastName'
	$UPN = Import-Csv -Path "$UserCSV" | Select-Object -First 1 | Select -Last 1 -ExpandProperty 'UserPrincipalName'
	
	if (($null -eq $Lastname) -and ($null -eq $LastName))
	{
		break
	}
	
	$SyncroName = $FirstName + " " + $LastName
	
	#Checks if contacts are already listed and skips to the next if they are. 
	$Log = (Get-Content $GetContactLog)
	
	
	$Name = (write $Log | select-string -pattern "$SyncroName")
	$Email = (write $Log | select-string -pattern "$UPN")
	
	if ($Name)
	{
		write "$DateStamp T$TimeStamp INFO: $SyncroName found, skipping this contact" | Out-File $Office365ToSnycroLog -Append -Encoding Ascii
		Get-Content $UserCSV | Where-Object ReadCount -ne 3 | Set-Content -Encoding Utf8 $UserCSVTrim
		Remove-Item $UserCSV
		Rename-Item -Path $UserCSVTrim -NewName $UserCSV -Force
		$Duplicates = "yes"
		$DupeCounter++
		continue
	}
	
	if ($Email)
	{
		write "$DateStamp T$TimeStamp INFO: $UPN found, skipping this contact" | Out-File $Office365ToSnycroLog -Append -Encoding Ascii
		Get-Content $UserCSV | Where-Object ReadCount -ne 3 | Set-Content -Encoding Utf8 $UserCSVTrim
		Remove-Item $UserCSV
		Rename-Item -Path $UserCSVTrim -NewName $UserCSV -Force
		$Duplicates = "yes"
		$DupeCounter++
		continue
	}
	
	#write $Curlarguments to upload the remaining contacts into Syncro
	
	
	$Curlarguments = "-X POST ""https://$subdomain.syncromsp.com/api/v1/contacts"" -H ""accept: application/json"" -H ""Authorization: $API"" -H ""Content-Type: application/json"" -d ""{\""customer_id\"":$CustID,\""name\"":\""$SyncroName\"",\""address1\"":\""\"",\""address2\"":\""\"",\""city\"":\""\"",\""state\"":\""\"",\""zip\"":\""\"",\""email\"":\""$UPN\"",\""phone\"":\""\"",\""mobile\"":\""\"",\""notes\"":\""\""}"
	$PostContact = (BuildBat -Type PostContact)
	$APICounter++
	APILimiter
	
	#Process contact data
	$PostContactLog = CleanData -SplitMarker "," -Type PostContact
	$CreatedName = (write $PostContactLog | select-string -pattern "name" | select -last 1)
	$CreatedEmail = (write $PostContactLog | select-string -pattern 'email"' | select -last 1)
	$CreateTime = (write $PostContactLog | select-string -pattern "created_at" | select -last 1)
	
	
	#Check reponse data for missing data and run error check if anything is missing. 
	if (($null -eq $CreatedName) -or ($null -eq $CreatedEmail) -or ($null -eq $CreateTime))
	{
		ErrorCheck -LogName PostContact
	}
	
	#Convert and clean data
	$CreatedName = ([string]$CreatedName)
	$CreatedName = ($CreatedName.replace('"name":"', '').Replace('"', ''))
	
	$CreatedEmail = ([string]$CreatedEmail)
	$CreatedEmail = ($CreatedEmail.replace('"email":"', '').Replace('"', ''))
	
	$CreateTime = ([string]$CreateTime)
	$CreateTime = ($CreateTime.replace('"created_at":"', '').Replace('"', ''))
	
	
	if ($CustBName -eq 'N/A')
	{
		$CustBName = ''
	}
	
	#Report results to console
	write "Contact uploaded to $CustName $CustBName"
	write "Creation Time is $CreateTime"
	write "Name is $CreatedName"
	write "Email is $CreatedEmail"
	write "Logged to $Office365ToSnycroLog"
	write " "
	write "$CreateTime  INFO: $CreatedName  $CreatedEmail Uploaded to $CustName $CustBName" | Out-File $Office365ToSnycroLog -Append -Encoding Ascii
	$UploadCounter++
	
	
	#Trim file for the next run.
	Get-Content $UserCSV | Where-Object ReadCount -ne 3 | Set-Content -Encoding Utf8 $UserCSVTrim
	Remove-Item $UserCSV
	Rename-Item -Path $UserCSVTrim -NewName $UserCSV -Force
	
}


#Write final report to console and log. 
write "Total Contact Uploads for this run: $UploadCounter"
write "Total API calls for this run: $APICounter"
write "$DateStamp T$TimeStamp Total Contact Uploads for this run: $UploadCounter" | Out-File $Office365ToSnycroLog -Append -Encoding Ascii
write "$DateStamp T$TimeStamp Total API calls for this run: $APICounter" | Out-File $Office365ToSnycroLog -Append -Encoding Ascii
write "$DateStamp T$TimeStamp Ending Run - Run Complete" | Out-File $Office365ToSnycroLog -Append -Encoding Ascii

#Write alert to console is duplicate contacts are skipped.
if ($Duplicates)
{
	write "WARNING: $DupeCounter Duplicate entires detected and skipped. Check $Office365ToSnycroLog for details"
}
#Final Cleanup
FinalCleanUp
