#Requires -Version 3.0
#Requires -Module dnsserver
#This File is in Unicode format.  Do not edit in an ASCII editor.

#region help text

<#
.SYNOPSIS
	Creates an inventory of Microsoft DNS using Microsoft Word, PDF, formatted text, or 
	HTML.
.DESCRIPTION
	Creates an inventory of Microsoft DNS using Microsoft Word, PDF, formatted text, or 
	HTML.

	Creates a document named either: 
		DNS Inventory Report for Server DNSServerName for the Domain <domain>.HTML 
		(or .DOCX or .PDF or .TXT).
		DNS Inventory Report for All DNS Servers for the Domain <domain>.HTML 
		(or .DOCX or .PDF or .TXT).
	
	Version 2.0 changes the default output report from Word to HTML.
	
	Word is NOT needed to run the script. This script outputs in Text and HTML.

	You do NOT have to run this script on a DNS server. This script was developed 
	and run from a Windows 10 VM.

	Requires the DNSServer module.
	
	Word and PDF documents include a Cover Page, Table of Contents, and Footer.
	
	Includes support for the following language versions of Microsoft Word:
		Catalan
		Chinese
		Danish
		Dutch
		English
		Finnish
		French
		German
		Norwegian
		Portuguese
		Spanish
		Swedish

	To run the script from a workstation, RSAT is required.
	
	Remote Server Administration Tools for Windows 7 with Service Pack 1 (SP1)
		http://www.microsoft.com/en-us/download/details.aspx?id=7887
		
	Remote Server Administration Tools for Windows 8 
		http://www.microsoft.com/en-us/download/details.aspx?id=28972
		
	Remote Server Administration Tools for Windows 8.1 
		http://www.microsoft.com/en-us/download/details.aspx?id=39296
		
	Remote Server Administration Tools for Windows 10
		http://www.microsoft.com/en-us/download/details.aspx?id=45520
		
.PARAMETER HTML
	Creates an HTML file with an .html extension.
	This parameter is set True if no other output format is selected.
.PARAMETER MSWord
	SaveAs DOCX file
	This parameter is disabled by default.
.PARAMETER PDF
	SaveAs PDF file instead of DOCX file.
	This parameter is disabled by default.
	The PDF file is roughly 5X to 10X larger than the DOCX file.
	This parameter requires Microsoft Word to be installed.
	This parameter uses the Word SaveAs PDF capability.
.PARAMETER Text
	Creates a formatted text file with a .txt extension.
	This parameter is disabled by default.
.PARAMETER AddDateTime
	Adds a date Timestamp to the end of the file name.
	The timestamp is in the format of yyyy-MM-dd_HHmm.
	June 1, 2020 at 6PM is 2020-06-01_1800.
	The output filename will be either:
		DNS Inventory Report for Server <server> for the Domain 
		<domain>_2020-06-01_1800.html (or .txt or .docx or .pdf)
		DNS Inventory for All DNS Servers for the Domain 
		<domain>_2020-06-01_1800.html (or .txt or .docx or .pdf)
	This parameter is disabled by default.
.PARAMETER AllDNSServers
	The script will process all AD DNS servers that are online.
	"DNS Inventory Report for All DNS Servers for the Domain <domain>" is used for the 
	report title.
	This parameter is disabled by default.
	
	If both ComputerName and AllDNSServers are used, AllDNSServers is used.
	This parameter has an alias of ALL.
.PARAMETER CompanyAddress
	Company Address used for the Cover Page, if the Cover Page has the Address field.
	
	The following Cover Pages have an Address field:
		Banded (Word 2013/2016)
		Contrast (Word 2010)
		Exposure (Word 2010)
		Filigree (Word 2013/2016)
		Ion (Dark) (Word 2013/2016)
		Retrospect (Word 2013/2016)
		Semaphore (Word 2013/2016)
		Tiles (Word 2010)
		ViewMaster (Word 2013/2016)
		
	This parameter is only valid with the MSWORD and PDF output parameters.
	This parameter has an alias of CA.
.PARAMETER CompanyEmail
	Company Email used for the Cover Page, if the Cover Page has the Email field.  
	
	The following Cover Pages have an Email field:
		Facet (Word 2013/2016)
	
	This parameter is only valid with the MSWORD and PDF output parameters.
	This parameter has an alias of CE.
.PARAMETER CompanyFax
	Company Fax used for the Cover Page, if the Cover Page has the Fax field.  
	
	The following Cover Pages have a Fax field:
		Contrast (Word 2010)
		Exposure (Word 2010)
	
	This parameter is only valid with the MSWORD and PDF output parameters.
	This parameter has an alias of CF.
.PARAMETER CompanyName
	Company Name used for the Cover Page.  
	The default value is contained in 
	HKCU:\Software\Microsoft\Office\Common\UserInfo\CompanyName or
	HKCU:\Software\Microsoft\Office\Common\UserInfo\Company, whichever is populated 
	on the computer running the script.

	This parameter is only valid with the MSWORD and PDF output parameters.
	This parameter has an alias of CN.
.PARAMETER CompanyPhone
	Company Phone used for the Cover Page, if the Cover Page has the Phone field.  
	
	The following Cover Pages have a Phone field:
		Contrast (Word 2010)
		Exposure (Word 2010)
	
	This parameter is only valid with the MSWORD and PDF output parameters.
	This parameter has an alias of CPh.
.PARAMETER ComputerName
	Specifies a computer used to run the script against.
	ComputerName can be entered as the NetBIOS name, FQDN, localhost or IP Address.
	If entered as localhost, the actual computer name is determined and used.
	If entered as an IP address, an attempt is made to determine and use the actual 
	computer name.
	
	Default is localhost.
	
	If both ComputerName and AllDNSServers are used, AllDNSServers is used.
.PARAMETER CoverPage
	What Microsoft Word Cover Page to use.
	Only Word 2010, 2013 and 2016 are supported.
	(default cover pages in Word en-US)
	
	Valid input is:
		Alphabet (Word 2010. Works)
		Annual (Word 2010. Doesn't work well for this report)
		Austere (Word 2010. Works)
		Austin (Word 2010/2013/2016. Doesn't work in 2013 or 2016, mostly 
		works in 2010 but Subtitle/Subject & Author fields need to be moved 
		after title box is moved up)
		Banded (Word 2013/2016. Works)
		Conservative (Word 2010. Works)
		Contrast (Word 2010. Works)
		Cubicles (Word 2010. Works)
		Exposure (Word 2010. Works if you like looking sideways)
		Facet (Word 2013/2016. Works)
		Filigree (Word 2013/2016. Works)
		Grid (Word 2010/2013/2016. Works in 2010)
		Integral (Word 2013/2016. Works)
		Ion (Dark) (Word 2013/2016. Top date doesn't fit; box needs to be 
		manually resized or font changed to 8 point)
		Ion (Light) (Word 2013/2016. Top date doesn't fit; box needs to be 
		manually resized or font changed to 8 point)
		Mod (Word 2010. Works)
		Motion (Word 2010/2013/2016. Works if top date is manually changed to 
		36 point)
		Newsprint (Word 2010. Works but date is not populated)
		Perspective (Word 2010. Works)
		Pinstripes (Word 2010. Works)
		Puzzle (Word 2010. Top date doesn't fit; box needs to be manually 
		resized or font changed to 14 point)
		Retrospect (Word 2013/2016. Works)
		Semaphore (Word 2013/2016. Works)
		Sideline (Word 2010/2013/2016. Doesn't work in 2013 or 2016, works in 
		2010)
		Slice (Dark) (Word 2013/2016. Doesn't work)
		Slice (Light) (Word 2013/2016. Doesn't work)
		Stacks (Word 2010. Works)
		Tiles (Word 2010. Date doesn't fit unless changed to 26 point)
		Transcend (Word 2010. Works)
		ViewMaster (Word 2013/2016. Works)
		Whisp (Word 2013/2016. Works)
		
	Default value is Sideline.
	This parameter has an alias of CP.
	This parameter is only valid with the MSWORD and PDF output parameters.
.PARAMETER Details
	Include Resource Record data for both Forward and Reverse lookup zones.
	
	Using this parameter can create an extremely large report.
	
	Default is to not include Resource Record information.
.PARAMETER Dev
	Clears errors at the beginning of the script.
	Outputs all errors to a text file at the end of the script.
	
	This is used when the script developer requests more troubleshooting data.
	The text file is placed in the same folder from where the script runs.
	
	This parameter is disabled by default.
.PARAMETER Folder
	Specifies the optional output folder to save the output report. 
.PARAMETER Log
	Generates a log file for troubleshooting.
.PARAMETER ScriptInfo
	Outputs information about the script to a text file.
	The text file is placed in the same folder from where the script runs.
	
	This parameter is disabled by default.
	This parameter has an alias of SI.
.PARAMETER Username
	Username used for the Cover Page and Footer.
	Default value is contained in $env:username
	This parameter has an alias of UN.
	This parameter is only valid with the MSWORD and PDF output parameters.
.PARAMETER SmtpServer
	Specifies the optional email server to send the output report. 
.PARAMETER SmtpPort
	Specifies the SMTP port. 
	Default is 25.
.PARAMETER UseSSL
	Specifies whether to use SSL for the SmtpServer.
	Default is False.
.PARAMETER From
	Specifies the username for the From email address.
	If SmtpServer is used, this is a required parameter.
.PARAMETER To
	Specifies the username for the To email address.
	If SmtpServer is used, this is a required parameter.
.EXAMPLE
	PS C:\PSScript > .\DNS_Inventory_V2.ps1 -MSWord 
	
	Will use all default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl 
	Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the Username.
	
	Tests to see if the computer, localhost, is a DNS server. 
	If it is, the script runs. If not, the script aborts.

	Creates a Microsoft Word document.
.EXAMPLE
	PS C:\PSScript > .\DNS_Inventory_V2.ps1 -ComputerName DNS01
	
	Will use all default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl 
	Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the Username.
	
	Runs the script against the DNS server named DNS01.

	Creates a Microsoft Word document.
.EXAMPLE
	PS C:\PSScript > .\DNS_Inventory_V2.ps1 -PDF
	
	Will use all default values and save the document as a PDF file.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl 
	Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the Username.
.EXAMPLE
	PS C:\PSScript > .\DNS_Inventory_V2.ps1 -TEXT

	Tests to see if the computer, localhost, is a DNS server. 
	If it is, the script runs. If not, the script aborts.

	Creates a formatted text file.
.EXAMPLE
	PS C:\PSScript > .\DNS_Inventory_V2.ps1

	Tests to see if the computer, localhost, is a DNS server. 
	If it is, the script runs. If not, the script aborts.

	Creates an HTML document.
.EXAMPLE
	PS C:\PSScript > .\DNS_Inventory_V2.ps1 -ComputerName localhost
	
	The script resolves localhost to $env:computername, for example, DNSServer01.
	The script runs remotely against the DNS server DNSServer01 and not localhost.
	The output filename uses the server name DNSServer01 and not localhost.
	
	Creates an HTML file.
.EXAMPLE
	PS C:\PSScript > .\DNS_Inventory_V2.ps1 -ComputerName 192.168.1.222
	
	The script attempts to resolve 192.168.1.222 to the DNS hostname, for example, 
	DNSServer01.
	The script runs remotely against the DNS server DNSServer01 and not 192.18.1.222.
	The output filename uses the server name DNSServer01 and not 192.168.1.222.

	Creates an HTML file.
.EXAMPLE
	PS C:\PSScript .\DNS_Inventory_V2.ps1 -CompanyName "Carl Webster Consulting" 
	-CoverPage "Mod" -UserName "Carl Webster" -ComputerName DNSServer01 -MSWord

	Will use:
		Carl Webster Consulting for the Company Name.
		Mod for the Cover Page format.
		Carl Webster for the Username.
	
	The script runs remotely against the DNS server DNSServer01.
.EXAMPLE
	PS C:\PSScript .\DNS_Inventory_V2.ps1 -CN "Carl Webster Consulting" -CP "Mod"
	-UN "Carl Webster" -ComputerName DNSServer02 -Details -MSWord

	Will use:
		Carl Webster Consulting for the Company Name (alias CN).
		Mod for the Cover Page format (alias CP).
		Carl Webster for the Username (alias UN).

	The script runs remotely against the DNS server DNSServer02.
	The output contains DNS Resource Record information.
.EXAMPLE
	PS C:\PSScript > .\DNS_Inventory_V2.ps1 -AddDateTime
	
	Adds a date Timestamp to the end of the file name.
	The timestamp is in the format of yyyy-MM-dd_HHmm.
	July 25, 2020 at 6PM is 2020-07-25_1800.
	The output filename is DNS Inventory Report for Server <server> for the Domain 
	<domain>_2020-07-25_1800.html.
.EXAMPLE
	PS C:\PSScript > .\DNS_Inventory_V2.ps1 -PDF -AddDateTime
	
	Will use all default values and save the document as a PDF file.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl 
	Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the Username.

	Adds a date Timestamp to the end of the file name.
	The timestamp is in the format of yyyy-MM-dd_HHmm.
	July 25, 2020 at 6PM is 2020-07-25_1800.
	The output filename is DNS Inventory Report for Server <server> for the Domain 
	<domain>_2020-07-25_1800.pdf.
.EXAMPLE
	PS C:\PSScript > .\DNS_Inventory_V2.ps1 -Folder \\FileServer\ShareName
	
	The output HTML file is saved in the path \\FileServer\ShareName.
.EXAMPLE
	PS C:\PSScript > .\DNS_Inventory_V2.ps1 -Details -MSWord
	
	Will use all default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl 
	Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the Username.
	
	Includes details for all Resource Records for both Forward and Reverse lookup zones.
.EXAMPLE
	PS C:\PSScript > .\DNS_Inventory_V2.ps1 -AllDNSServers
	
	The script finds all AD DNS servers and processes all online servers.
.EXAMPLE
	PS C:\PSScript > .\DNS_Inventory_V2.ps1 -ComputerName DNSServer01 -AllDNSServers
	
	Even though DNSServer01 is specified, the script finds all AD DNS servers 
	and processes all online servers.
.EXAMPLE
	PS C:\PSScript > .\DNS_Inventory_V2.ps1 -HTML -MSWord -PDF -Text -Dev -ScriptInfo -Log 
	-ComputerName DNSServer
	
	Creates four reports: HTML, Microsoft Word, PDF, and plain text.
	
	Creates a text file named DNSInventoryScriptErrors_yyyy-MM-dd_HHmm for the Domain 
	<domain>.txt that contains up to the last 250 errors reported by the script.
	
	Creates a text file named DNSInventoryScriptInfo_yyyy-MM-dd_HHmm for the Domain 
	<domain>.txt that contains all the script parameters and other basic information.
	
	Creates a text file for transcript logging named 
	DNSDocScriptTranscript_yyyy-MM-dd_HHmm for the Domain <domain>.txt.

	For Microsoft Word and PDF, uses all Default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl 
	Webster" or 
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the Username.
	
	The script runs remotely against the DNS server DNSServer.
.EXAMPLE
	PS C:\PSScript > .\DNS_Inventory_V2.ps1 -SmtpServer mail.domain.tld -From 
	XDAdmin@domain.tld -To ITGroup@domain.tld	

	The script will use the email server mail.domain.tld, sending from XDAdmin@domain.tld, 
	sending to ITGroup@domain.tld.

	The script will use the default SMTP port 25 and does not use SSL.

	If the current user's credentials are not valid to send email, 
	the user will be prompted to enter valid credentials.
.EXAMPLE
	PS C:\PSScript > .\DNS_Inventory_V2.ps1 -SmtpServer mailrelay.domain.tld -From 
	Anonymous@domain.tld -To ITGroup@domain.tld	

	***SENDING UNAUTHENTICATED EMAIL***

	The script will use the email server mailrelay.domain.tld, sending from 
	anonymous@domain.tld, sending to ITGroup@domain.tld.

	To send unauthenticated email using an email relay server requires the From email account 
	to use the name Anonymous.

	The script will use the default SMTP port 25 and does not use SSL.
	
	***GMAIL/G SUITE SMTP RELAY***
	https://support.google.com/a/answer/2956491?hl=en
	https://support.google.com/a/answer/176600?hl=en

	To send email using a Gmail or g-suite account, you may have to turn ON
	the "Less secure app access" option on your account.
	***GMAIL/G SUITE SMTP RELAY***

	The script generates an anonymous, secure password for the anonymous@domain.tld 
	account.
.EXAMPLE
	PS C:\PSScript > .\DNS_Inventory_V2.ps1 -SmtpServer 
	labaddomain-com.mail.protection.outlook.com -UseSSL -From 
	SomeEmailAddress@labaddomain.com -To ITGroupDL@labaddomain.com	

	***OFFICE 365 Example***

	https://docs.microsoft.com/en-us/exchange/mail-flow-best-practices/how-to-set-up-a-multifunction-device-or-application-to-send-email-using-office-3
	
	This uses Option 2 from the above link.
	
	***OFFICE 365 Example***

	The script will use the email server labaddomain-com.mail.protection.outlook.com, 
	sending from SomeEmailAddress@labaddomain.com, sending to ITGroupDL@labaddomain.com.

	The script will use the default SMTP port 25 and will use SSL.
.EXAMPLE
	PS C:\PSScript > .\DNS_Inventory_V2.ps1 -SmtpServer smtp.office365.com -SmtpPort 587
	-UseSSL -From Webster@CarlWebster.com -To ITGroup@CarlWebster.com	

	The script will use the email server smtp.office365.com on port 587 using SSL, 
	sending from webster@carlwebster.com, sending to ITGroup@carlwebster.com.

	If the current user's credentials are not valid to send email, 
	the user will be prompted to enter valid credentials.
.EXAMPLE
	PS C:\PSScript > .\DNS_Inventory_V2.ps1 -SmtpServer smtp.gmail.com -SmtpPort 587
	-UseSSL -From Webster@CarlWebster.com -To ITGroup@CarlWebster.com	

	*** NOTE ***
	To send email using a Gmail or g-suite account, you may have to turn ON
	the "Less secure app access" option on your account.
	*** NOTE ***
	
	The script will use the email server smtp.gmail.com on port 587 using SSL, 
	sending from webster@gmail.com, sending to ITGroup@carlwebster.com.

	If the current user's credentials are not valid to send email, 
	the user will be prompted to enter valid credentials.
.INPUTS
	None. You cannot pipe objects to this script.
.OUTPUTS
	No objects are output from this script.  This script creates a Word, PDF, HTML or 
	formatted text document.
.NOTES
	NAME: DNS_Inventory_V2.ps1
	VERSION: 2.00
	AUTHOR: Carl Webster and Michael B. Smith
	LASTEDIT: November 4, 2020
#>

#endregion

#region script parameters
#thanks to @jeffwouters and Michael B. Smith for helping me with these parameters
[CmdletBinding(SupportsShouldProcess = $False, ConfirmImpact = "None", DefaultParameterSetName = "") ]

Param(
	[parameter(Mandatory=$False)] 
	[Switch]$HTML=$False,

	[parameter(ParameterSetName="WordPDF",Mandatory=$False)] 
	[Switch]$MSWord=$False,

	[parameter(ParameterSetName="WordPDF",Mandatory=$False)] 
	[Switch]$PDF=$False,

	[parameter(Mandatory=$False)] 
	[Switch]$Text=$False,

	[parameter(Mandatory=$False)] 
	[string]$ComputerName="LocalHost",

	[parameter(Mandatory=$False)] 
	[Alias("ALL")]
	[Switch]$AllDNSServers=$False,
	
	[parameter(Mandatory=$False)] 
	[Switch]$AddDateTime=$False,
	
	[parameter(ParameterSetName="WordPDF",Mandatory=$False)] 
	[Alias("CA")]
	[ValidateNotNullOrEmpty()]
	[string]$CompanyAddress="",
    
	[parameter(ParameterSetName="WordPDF",Mandatory=$False)] 
	[Alias("CE")]
	[ValidateNotNullOrEmpty()]
	[string]$CompanyEmail="",
    
	[parameter(ParameterSetName="WordPDF",Mandatory=$False)] 
	[Alias("CF")]
	[ValidateNotNullOrEmpty()]
	[string]$CompanyFax="",
    
	[parameter(ParameterSetName="WordPDF",Mandatory=$False)] 
	[Alias("CN")]
	[ValidateNotNullOrEmpty()]
	[string]$CompanyName="",
    
	[parameter(ParameterSetName="WordPDF",Mandatory=$False)] 
	[Alias("CPh")]
	[ValidateNotNullOrEmpty()]
	[string]$CompanyPhone="",
    
	[parameter(ParameterSetName="WordPDF",Mandatory=$False)] 
	[Alias("CP")]
	[ValidateNotNullOrEmpty()]
	[string]$CoverPage="Sideline", 

	[parameter(Mandatory=$False)] 
	[Switch]$Details=$False,
	
	[parameter(Mandatory=$False)] 
	[Switch]$Dev=$False,
	
	[parameter(Mandatory=$False)] 
	[string]$Folder="",
	
	[parameter(Mandatory=$False)] 
	[Switch]$Log=$False,
	
	[parameter(Mandatory=$False)] 
	[Alias("SI")]
	[Switch]$ScriptInfo=$False,
	
	[parameter(ParameterSetName="WordPDF",Mandatory=$False)] 
	[Alias("UN")]
	[ValidateNotNullOrEmpty()]
	[string]$Username=$env:username,

	[parameter(Mandatory=$False)] 
	[string]$SmtpServer="",

	[parameter(Mandatory=$False)] 
	[int]$SmtpPort=25,

	[parameter(Mandatory=$False)] 
	[switch]$UseSSL=$False,

	[parameter(Mandatory=$False)] 
	[string]$From="",

	[parameter(Mandatory=$False)] 
	[string]$To=""

	)
#endregion

#region script change log	
#Created by Carl Webster and Michael B. Smith
#webster@carlwebster.com
#@carlwebster on Twitter
#http://www.CarlWebster.com
#Created on February 10, 2016
#Version 1.00 released to the community on July 25, 2016

#Version 2.00 4-Nov-2020
#	Added an Appendix A to give an overview of Several DNS server and zone configuration Items when using -AllDNSServers:
#		DNS Forwarders
#		Zone Type
#		AD Integration
#		Signed
#		Dynamic Updates
#		Replication Scope
#		Aging Enabled
#		Refresh Interval
#		NoRefresh Interval
#		Scavenge Servers
#	Added processing Forward Lookup Zones that are Signed
#		Key Master
#		Next Secure (NSEC)
#		Trust Anchor
#		Advanced
#	Changed all Write-Verbose $(Get-Date) to add -Format G to put the dates in the user's locale
#	Changed color variables $wdColorGray15 and $wdColorGray05 from [long] to [int]
#	Cleaned up the formatting of Text output
#	Commented out Function CheckHTMLColor as it is no longer needed
#	For the -Dev, -Log, and -ScriptInfo output files, add the text "for the Domain <domain>"
#	General code cleanup
#	HTML is now the default output format
#	Removed the original function TestComputerName and renamed TestComputerName2 to TestComputerName
#		Added Functions testPort and testPortsOnOneIP
#	Stopped using a Switch statement for HTML colors and use a pre-calculated HTML array (for speed)
#	Updated Function ShowScriptOptions and ProcessScriptEnd for allowing multiple output types
#	Updated the help text
#	Updated the ReadMe file (https://carlwebster.sharefile.com/d-s247b4252c4e4865a)
#	Updated the following Functions to the latest versions
#		AddHTMLTable
#		AddWordTable
#		CheckWordPrereq
#		FormatHTMLTable
#		ProcessDocumentOutput
#		SaveandCloseDocumentandShutdownWord
#		SaveandCloseHTMLDocument
#		SaveandCloseTextDocument
#		SetFilenames
#		SetWordCellFormat
#		SetWordHashTable
#		SetupHTML
#		SetupText
#		SetupWord
#		WriteHTMLLine
#	Updated the report title and output filenames to:
#		For using -Computername:
#			DNS Inventory Report for Server <DNSServerName> for the Domain <domain>
#		For using -AllDNSServers:
#			DNS Inventory Report for All DNS Servers for the Domain <domain>
#	You can now select multiple output formats. This required extensive code changes.
#
#Version 1.22 8-May-2020
#	Add checking for a Word version of 0, which indicates the Office installation needs repairing
#	Change color variables $wdColorGray15 and $wdColorGray05 from [long] to [int]
#	Change location of the -Dev, -Log, and -ScriptInfo output files from the script folder to the -Folder location (Thanks to Guy Leech for the "suggestion")
#	Change Text output to use [System.Text.StringBuilder]
#		Updated Functions Line and SaveAndCloseTextDocument
#	Fixed by MBS: When the root hint IP address is an array, report on all entries of the array, instead of just the first entry
#	For all the uses of the Get-DNSServer cmdlet, to stop the excess "garbage" spewed forth by that cmdlet, use 2>$Null 3>$Null 4>$Null 
#	Reformatted the terminating Write-Error messages to make them more visible and readable in the console
#	Remove the SMTP parameterset and manually verify the parameters
#	Update Function SendEmail to handle anonymous unauthenticated email
#	Update Function SetWordCellFormat to change parameter $BackgroundColor to [int]
#	Update Help Text
#
#Version 1.21 Not released
#	Fixed by MBS: When the root hint IP address is an array, report on all entries of the array, instead of just the first entry
#
#Version 1.20 13-Feb-2020
#	Added -AllDNSServers (ALL) parameter to process all AD DNS servers that are online
#		Added text file (BadDNSServers_yyyy-MM-dd_HHmm.txt) of the AD DNS servers that 
#		are either offline or no longer have DNS installed
#	Added a new Function TestComputerName2 to support the -AllDNSServers parameter
#	Added the DNS server name to the section title for the DNS server
#	Fix Swedish Table of Contents (Thanks to Johan Kallio)
#		From 
#			'sv-'	{ 'Automatisk innehållsförteckning2'; Break }
#		To
#			'sv-'	{ 'Automatisk innehållsförteckn2'; Break }
#	General code cleanup
#	The following functions were updated to support the -AllDNSServers parameter:
#		ProcessDNSServer
#		OutputDNSServer
#		ProcessForwardLookupZones
#		OutputLookupZone
#		ProcessLookupZoneDetails
#		ProcessTrustPoints
#		ShowScriptOptions
#		ProcessScriptStart
#		ProcessScriptEnd
#	Updated Function CheckWordPrereq to match the other scripts
#	Updated the help text
#
#Version 1.12 6-Dec-2019
#	Fixed text string "Use root hint if no forwarders are available" to "Use root hints if no forwarders are available"
#	Fixed spacing error in Text output for "Use root hints if no forwarders are available"
#	For Name Servers, if the IP Address is Null or Empty, use "Unable to retrieve an IP Address"
#		For Word/PDF and HTML output put the invalid Name Server and "Unable to retrieve an IP Address" in Red
#		For Text output use "***Unable to retrieve an IP Address***"
#	Reorder parameters
#	Update help text
#
#Version 1.11 25-Oct-2019
#	Fixed the sorting of Root Hint servers thanks to MBS
#	Fixed the sorting on Name Servers
#
#Version 1.10 6-Apr-2018
#	Code clean up from Visual Studio Code
#
#Version 1.09 2-Mar-2018
#	Added Log switch to create a transcript log
#	I found two "If($Var = something)" which are now "If($Var -eq something)"
#	In the function OutputLookupZoneDetails, with the "=" changed to "-eq" fix, the hostname was now always blank. Fixed.
#	Many Switch bocks I never added "; Break" to. Those are now fixed.
#	Update functions ShowScriptOutput and ProcessScriptEnd for new Log parameter
#	Updated help text
#	Updated the WriteWordLine function 
#
#Version 1.08 8-Dec-2017
#	Updated Function WriteHTMLLine with fixes from the script template
#
#Version 1.07 13-Nov-2017
#	Added Scavenge Server(s) to Zone Properties General section
#	Added the domain name of the computer used for -ComputerName to the output filename
#	Fixed output of Name Server IP address(es) in Zone properties
#	For Word/PDF output added the domain name of the computer used for -ComputerName to the report title
#	General code cleanup
#	In Text output, fixed alignment of "Scavenging period" in DNS Server Properties
#	Removed code that made sure all Parameters were set to default values if for some reason they did exist or values were $Null
#	Reordered the parameters in the help text and parameter list so they match and are grouped better
#	Replaced _SetDocumentProperty function with Jim Moyle's Set-DocumentProperty function
#	Updated Function ProcessScriptEnd for the new Cover Page properties and Parameters
#	Updated Function ShowScriptOptions for the new Cover Page properties and Parameters
#	Updated Function UpdateDocumentProperties for the new Cover Page properties and Parameters
#	Updated help text
#
#Version 1.06 13-Feb-2017
#	Fixed French wording for Table of Contents 2 (Thanks to David Rouquier)
#
#Version 1.05 7-Nov-2016
#	Added Chinese language support
#
#Version 1.04 22-Oct-2016
#	More refinement of HTML output
#
#Version 1.03 19-Oct-2016
#	Fixed formatting issues with HTML headings output
#
#Version 1.02 19-Aug-2016
#	Fixed several misspelled words
#
#Version 1.01 16-Aug-2016
#	Added support for the four Record Types created by implementing DNSSEC
#		NSec
#		NSec3
#		NSec3Param
#		RRSig
#
#HTML functions contributed by Ken Avram October 2014
#HTML Functions FormatHTMLTable and AddHTMLTable modified by Jake Rutski May 2015
#endregion

#region initial variable testing and setup
Set-StrictMode -Version Latest

#force  on
$PSDefaultParameterValues = @{"*:Verbose"=$True}
$SaveEAPreference = $ErrorActionPreference
$ErrorActionPreference = 'SilentlyContinue'
$global:emailCredentials = $Null
$Script:RptDomain = (Get-WmiObject -computername $ComputerName win32_computersystem).Domain

If($ComputerName -eq "localhost")
{
	$ComputerName = $env:ComputerName
	Write-Verbose "$(Get-Date -Format G): Computer name has been changed from localhost to $ComputerName"
}
	
If($MSWord -eq $False -and $PDF -eq $False -and $Text -eq $False -and $HTML -eq $False)
{
	$HTML = $True
}

If($MSWord)
{
	Write-Verbose "$(Get-Date -Format G): MSWord is set"
}
If($PDF)
{
	Write-Verbose "$(Get-Date -Format G): PDF is set"
}
If($Text)
{
	Write-Verbose "$(Get-Date -Format G): Text is set"
}
If($HTML)
{
	Write-Verbose "$(Get-Date -Format G): HTML is set"
}

If($Folder -ne "")
{
	Write-Verbose "$(Get-Date -Format G): Testing folder path"
	#does it exist
	If(Test-Path $Folder -EA 0)
	{
		#it exists, now check to see if it is a folder and not a file
		If(Test-Path $Folder -pathType Container -EA 0)
		{
			#it exists and it is a folder
			Write-Verbose "$(Get-Date -Format G): Folder path $Folder exists and is a folder"
		}
		Else
		{
			#it exists but it is a file not a folder
			Write-Error "
			`n`n
			`t`t
			Folder $Folder is a file, not a folder.
			`n`n
			`t`t
			Script cannot continue.
			`n`n
			"
			Exit
		}
	}
	Else
	{
		#does not exist
		Write-Error "
		`n`n
		`t`t
		Folder $Folder does not exist.
		`n`n
		`t`t
		Script cannot continue.
		`n`n
		"
		Exit
	}
}

If($Folder -eq "")
{
	$Script:pwdpath = $pwd.Path
}
Else
{
	$Script:pwdpath = $Folder
}

If($Script:pwdpath.EndsWith("\"))
{
	#remove the trailing \
	$Script:pwdpath = $Script:pwdpath.SubString(0, ($Script:pwdpath.Length - 1))
}

#V1.09 added
If($Log) 
{
	#start transcript logging
	$Script:LogPath = "$Script:pwdpath\DNSDocScriptTranscript_$(Get-Date -f yyyy-MM-dd_HHmm) for the Domain $Script:RptDomain.txt"
	
	try 
	{
		Start-Transcript -Path $Script:LogPath -Force -Verbose:$false | Out-Null
		Write-Verbose "$(Get-Date -Format G): Transcript/log started at $Script:LogPath"
		$Script:StartLog = $true
	} 
	catch 
	{
		Write-Verbose "$(Get-Date -Format G): Transcript/log failed at $Script:LogPath"
		$Script:StartLog = $false
	}
}

If($Dev)
{
	$Error.Clear()
	$Script:DevErrorFile = "$Script:pwdPath\DNSInventoryScriptErrors_$(Get-Date -f yyyy-MM-dd_HHmm) for the Domain $Script:RptDomain.txt"
}

If(![String]::IsNullOrEmpty($SmtpServer) -and [String]::IsNullOrEmpty($From) -and [String]::IsNullOrEmpty($To))
{
	Write-Error "
	`n`n
	`t`t
	You specified an SmtpServer but did not include a From or To email address.
	`n`n
	`t`t
	Script cannot continue.
	`n`n"
	Exit
}
If(![String]::IsNullOrEmpty($SmtpServer) -and [String]::IsNullOrEmpty($From) -and ![String]::IsNullOrEmpty($To))
{
	Write-Error "
	`n`n
	`t`t
	You specified an SmtpServer and a To email address but did not include a From email address.
	`n`n
	`t`t
	Script cannot continue.
	`n`n"
	Exit
}
If(![String]::IsNullOrEmpty($SmtpServer) -and [String]::IsNullOrEmpty($To) -and ![String]::IsNullOrEmpty($From))
{
	Write-Error "
	`n`n
	`t`t
	You specified an SmtpServer and a From email address but did not include a To email address.
	`n`n
	`t`t
	Script cannot continue.
	`n`n"
	Exit
}
If(![String]::IsNullOrEmpty($From) -and ![String]::IsNullOrEmpty($To) -and [String]::IsNullOrEmpty($SmtpServer))
{
	Write-Error "
	`n`n
	`t`t
	You specified From and To email addresses but did not include the SmtpServer.
	`n`n
	`t`t
	Script cannot continue.
	`n`n"
	Exit
}
If(![String]::IsNullOrEmpty($From) -and [String]::IsNullOrEmpty($SmtpServer))
{
	Write-Error "
	`n`n
	`t`t
	You specified a From email address but did not include the SmtpServer.
	`n`n
	`t`t
	Script cannot continue.
	`n`n"
	Exit
}
If(![String]::IsNullOrEmpty($To) -and [String]::IsNullOrEmpty($SmtpServer))
{
	Write-Error "
	`n`n
	`t`t
	You specified a To email address but did not include the SmtpServer.
	`n`n
	`t`t
	Script cannot continue.
	`n`n"
	Exit
}

#endregion

#region initialize variables for word html and text
[string]$Script:RunningOS = (Get-WmiObject -class Win32_OperatingSystem -EA 0).Caption

If($MSWord -or $PDF)
{
	#try and fix the issue with the $CompanyName variable
	$Script:CoName = $CompanyName
	Write-Verbose "$(Get-Date -Format G): CoName is $($Script:CoName)"
	
	#the following values were attained from 
	#http://groovy.codehaus.org/modules/scriptom/1.6.0/scriptom-office-2K3-tlb/apidocs/
	#http://msdn.microsoft.com/en-us/library/office/aa211923(v=office.11).aspx
	[int]$wdAlignPageNumberRight = 2
	[int]$wdColorGray15 = 14277081
	[int]$wdColorGray05 = 15987699 
	[int]$wdMove = 0
	[int]$wdSeekMainDocument = 0
	[int]$wdSeekPrimaryFooter = 4
	[int]$wdStory = 6
	[int]$wdColorRed = 255
	[int]$wdColorWhite = 16777215
	[int]$wdColorBlack = 0
	[int]$wdWord2007 = 12
	[int]$wdWord2010 = 14
	[int]$wdWord2013 = 15
	[int]$wdWord2016 = 16
	[int]$wdFormatDocumentDefault = 16
	[int]$wdFormatPDF = 17
	#http://blogs.technet.com/b/heyscriptingguy/archive/2006/03/01/how-can-i-right-align-a-single-column-in-a-word-table.aspx
	#http://msdn.microsoft.com/en-us/library/office/ff835817%28v=office.15%29.aspx
	[int]$wdAlignParagraphLeft = 0
	[int]$wdAlignParagraphCenter = 1
	[int]$wdAlignParagraphRight = 2
	#http://msdn.microsoft.com/en-us/library/office/ff193345%28v=office.15%29.aspx
	[int]$wdCellAlignVerticalTop = 0
	[int]$wdCellAlignVerticalCenter = 1
	[int]$wdCellAlignVerticalBottom = 2
	#http://msdn.microsoft.com/en-us/library/office/ff844856%28v=office.15%29.aspx
	[int]$wdAutoFitFixed = 0
	[int]$wdAutoFitContent = 1
	[int]$wdAutoFitWindow = 2
	#http://msdn.microsoft.com/en-us/library/office/ff821928%28v=office.15%29.aspx
	[int]$wdAdjustNone = 0
	[int]$wdAdjustProportional = 1
	[int]$wdAdjustFirstColumn = 2
	[int]$wdAdjustSameWidth = 3

	[int]$PointsPerTabStop = 36
	[int]$Indent0TabStops = 0 * $PointsPerTabStop
	[int]$Indent1TabStops = 1 * $PointsPerTabStop
	[int]$Indent2TabStops = 2 * $PointsPerTabStop
	[int]$Indent3TabStops = 3 * $PointsPerTabStop
	[int]$Indent4TabStops = 4 * $PointsPerTabStop

	# http://www.thedoctools.com/index.php?show=wt_style_names_english_danish_german_french
	[int]$wdStyleHeading1 = -2
	[int]$wdStyleHeading2 = -3
	[int]$wdStyleHeading3 = -4
	[int]$wdStyleHeading4 = -5
	[int]$wdStyleNoSpacing = -158
	[int]$wdTableGrid = -155
	[int]$wdTableLightListAccent3 = -206

	#http://groovy.codehaus.org/modules/scriptom/1.6.0/scriptom-office-2K3-tlb/apidocs/org/codehaus/groovy/scriptom/tlb/office/word/WdLineStyle.html
	[int]$wdLineStyleNone = 0
	[int]$wdLineStyleSingle = 1

	[int]$wdHeadingFormatTrue = -1
	[int]$wdHeadingFormatFalse = 0 
	
	[string]$Script:RunningOS = (Get-WmiObject -class Win32_OperatingSystem -EA 0).Caption
}

If($HTML)
{
	#V2.23 Prior versions used Set-Variable. That hid the variables
	#from @code. So MBS switched to using $global:

    $global:htmlredmask       = "#FF0000" 4>$Null
    $global:htmlcyanmask      = "#00FFFF" 4>$Null
    $global:htmlbluemask      = "#0000FF" 4>$Null
    $global:htmldarkbluemask  = "#0000A0" 4>$Null
    $global:htmllightbluemask = "#ADD8E6" 4>$Null
    $global:htmlpurplemask    = "#800080" 4>$Null
    $global:htmlyellowmask    = "#FFFF00" 4>$Null
    $global:htmllimemask      = "#00FF00" 4>$Null
    $global:htmlmagentamask   = "#FF00FF" 4>$Null
    $global:htmlwhitemask     = "#FFFFFF" 4>$Null
    $global:htmlsilvermask    = "#C0C0C0" 4>$Null
    $global:htmlgraymask      = "#808080" 4>$Null
    $global:htmlblackmask     = "#000000" 4>$Null
    $global:htmlorangemask    = "#FFA500" 4>$Null
    $global:htmlmaroonmask    = "#800000" 4>$Null
    $global:htmlgreenmask     = "#008000" 4>$Null
    $global:htmlolivemask     = "#808000" 4>$Null

    $global:htmlbold        = 1 4>$Null
    $global:htmlitalics     = 2 4>$Null
    $global:htmlred         = 4 4>$Null
    $global:htmlcyan        = 8 4>$Null
    $global:htmlblue        = 16 4>$Null
    $global:htmldarkblue    = 32 4>$Null
    $global:htmllightblue   = 64 4>$Null
    $global:htmlpurple      = 128 4>$Null
    $global:htmlyellow      = 256 4>$Null
    $global:htmllime        = 512 4>$Null
    $global:htmlmagenta     = 1024 4>$Null
    $global:htmlwhite       = 2048 4>$Null
    $global:htmlsilver      = 4096 4>$Null
    $global:htmlgray        = 8192 4>$Null
    $global:htmlolive       = 16384 4>$Null
    $global:htmlorange      = 32768 4>$Null
    $global:htmlmaroon      = 65536 4>$Null
    $global:htmlgreen       = 131072 4>$Null
	$global:htmlblack       = 262144 4>$Null

	$global:htmlsb          = ( $htmlsilver -bor $htmlBold ) ## point optimization

	$global:htmlColor = 
	@{
		$htmlred       = $htmlredmask
		$htmlcyan      = $htmlcyanmask
		$htmlblue      = $htmlbluemask
		$htmldarkblue  = $htmldarkbluemask
		$htmllightblue = $htmllightbluemask
		$htmlpurple    = $htmlpurplemask
		$htmlyellow    = $htmlyellowmask
		$htmllime      = $htmllimemask
		$htmlmagenta   = $htmlmagentamask
		$htmlwhite     = $htmlwhitemask
		$htmlsilver    = $htmlsilvermask
		$htmlgray      = $htmlgraymask
		$htmlolive     = $htmlolivemask
		$htmlorange    = $htmlorangemask
		$htmlmaroon    = $htmlmaroonmask
		$htmlgreen     = $htmlgreenmask
		$htmlblack     = $htmlblackmask
	}
}
#endregion

#region word specific functions
Function SetWordHashTable
{
	Param([string]$CultureCode)

	#optimized by Michael B. Smith
	
	# DE and FR translations for Word 2010 by Vladimir Radojevic
	# Vladimir.Radojevic@Commerzreal.com

	# DA translations for Word 2010 by Thomas Daugaard
	# Citrix Infrastructure Specialist at edgemo A/S

	# CA translations by Javier Sanchez 
	# CEO & Founder 101 Consulting

	#ca - Catalan
	#da - Danish
	#de - German
	#en - English
	#es - Spanish
	#fi - Finnish
	#fr - French
	#nb - Norwegian
	#nl - Dutch
	#pt - Portuguese
	#sv - Swedish
	#zh - Chinese
	
	[string]$toc = $(
		Switch ($CultureCode)
		{
			'ca-'	{ 'Taula automática 2'; Break }
			'da-'	{ 'Automatisk tabel 2'; Break }
			'de-'	{ 'Automatische Tabelle 2'; Break }
			'en-'	{ 'Automatic Table 2'; Break }
			'es-'	{ 'Tabla automática 2'; Break }
			'fi-'	{ 'Automaattinen taulukko 2'; Break }
			'fr-'	{ 'Table automatique 2'; Break } #changed 10-feb-2017 david roquier and samuel legrand
			'nb-'	{ 'Automatisk tabell 2'; Break }
			'nl-'	{ 'Automatische inhoudsopgave 2'; Break }
			'pt-'	{ 'Sumário Automático 2'; Break }
			'sv-'	{ 'Automatisk innehållsförteckn2'; Break }
			'zh-'	{ '自动目录 2'; Break }
		}
	)

	$Script:myHash                      = @{}
	$Script:myHash.Word_TableOfContents = $toc
	$Script:myHash.Word_NoSpacing       = $wdStyleNoSpacing
	$Script:myHash.Word_Heading1        = $wdStyleheading1
	$Script:myHash.Word_Heading2        = $wdStyleheading2
	$Script:myHash.Word_Heading3        = $wdStyleheading3
	$Script:myHash.Word_Heading4        = $wdStyleheading4
	$Script:myHash.Word_TableGrid       = $wdTableGrid
}

Function GetCulture
{
	Param([int]$WordValue)
	
	#codes obtained from http://support.microsoft.com/kb/221435
	#http://msdn.microsoft.com/en-us/library/bb213877(v=office.12).aspx
	$CatalanArray = 1027
	$ChineseArray = 2052,3076,5124,4100
	$DanishArray = 1030
	$DutchArray = 2067, 1043
	$EnglishArray = 3081, 10249, 4105, 9225, 6153, 8201, 5129, 13321, 7177, 11273, 2057, 1033, 12297
	$FinnishArray = 1035
	$FrenchArray = 2060, 1036, 11276, 3084, 12300, 5132, 13324, 6156, 8204, 10252, 7180, 9228, 4108
	$GermanArray = 1031, 3079, 5127, 4103, 2055
	$NorwegianArray = 1044, 2068
	$PortugueseArray = 1046, 2070
	$SpanishArray = 1034, 11274, 16394, 13322, 9226, 5130, 7178, 12298, 17418, 4106, 18442, 19466, 6154, 15370, 10250, 20490, 3082, 14346, 8202
	$SwedishArray = 1053, 2077

	#ca - Catalan
	#da - Danish
	#de - German
	#en - English
	#es - Spanish
	#fi - Finnish
	#fr - French
	#nb - Norwegian
	#nl - Dutch
	#pt - Portuguese
	#sv - Swedish
	#zh - Chinese

	Switch ($WordValue)
	{
		{$CatalanArray -contains $_}	{$CultureCode = "ca-"}
		{$ChineseArray -contains $_}	{$CultureCode = "zh-"}
		{$DanishArray -contains $_}		{$CultureCode = "da-"}
		{$DutchArray -contains $_}		{$CultureCode = "nl-"}
		{$EnglishArray -contains $_}	{$CultureCode = "en-"}
		{$FinnishArray -contains $_}	{$CultureCode = "fi-"}
		{$FrenchArray -contains $_}		{$CultureCode = "fr-"}
		{$GermanArray -contains $_}		{$CultureCode = "de-"}
		{$NorwegianArray -contains $_}	{$CultureCode = "nb-"}
		{$PortugueseArray -contains $_}	{$CultureCode = "pt-"}
		{$SpanishArray -contains $_}	{$CultureCode = "es-"}
		{$SwedishArray -contains $_}	{$CultureCode = "sv-"}
		Default {$CultureCode = "en-"}
	}
	
	Return $CultureCode
}

Function ValidateCoverPage
{
	Param([int]$xWordVersion, [string]$xCP, [string]$CultureCode)
	
	$xArray = ""
	
	Switch ($CultureCode)
	{
		'ca-'	{
				If($xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Austin", "En bandes", "Faceta", "Filigrana",
					"Integral", "Ió (clar)", "Ió (fosc)", "Línia lateral",
					"Moviment", "Quadrícula", "Retrospectiu", "Sector (clar)",
					"Sector (fosc)", "Semàfor", "Visualització principal", "Whisp")
				}
				ElseIf($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("Austin", "En bandes", "Faceta", "Filigrana",
					"Integral", "Ió (clar)", "Ió (fosc)", "Línia lateral",
					"Moviment", "Quadrícula", "Retrospectiu", "Sector (clar)",
					"Sector (fosc)", "Semàfor", "Visualització", "Whisp")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alfabet", "Anual", "Austin", "Conservador",
					"Contrast", "Cubicles", "Diplomàtic", "Exposició",
					"Línia lateral", "Mod", "Mosiac", "Moviment", "Paper de diari",
					"Perspectiva", "Piles", "Quadrícula", "Sobri",
					"Transcendir", "Trencaclosques")
				}
			}

		'da-'	{
				If($xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Austin", "BevægElse", "Brusen", "Facet", "Filigran", 
					"Gitter", "Integral", "Ion (lys)", "Ion (mørk)", 
					"Retro", "Semafor", "Sidelinje", "Stribet", 
					"Udsnit (lys)", "Udsnit (mørk)", "Visningsmaster")
				}
				ElseIf($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("BevægElse", "Brusen", "Ion (lys)", "Filigran",
					"Retro", "Semafor", "Visningsmaster", "Integral",
					"Facet", "Gitter", "Stribet", "Sidelinje", "Udsnit (lys)",
					"Udsnit (mørk)", "Ion (mørk)", "Austin")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("BevægElse", "Moderat", "Perspektiv", "Firkanter",
					"Overskrid", "Alfabet", "Kontrast", "Stakke", "Fliser", "Gåde",
					"Gitter", "Austin", "Eksponering", "Sidelinje", "Enkel",
					"Nålestribet", "Årlig", "Avispapir", "Tradionel")
				}
			}

		'de-'	{
				If($xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Austin", "Bewegung", "Facette", "Filigran", 
					"Gebändert", "Integral", "Ion (dunkel)", "Ion (hell)", 
					"Pfiff", "Randlinie", "Raster", "Rückblick", 
					"Segment (dunkel)", "Segment (hell)", "Semaphor", 
					"ViewMaster")
				}
				ElseIf($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("Semaphor", "Segment (hell)", "Ion (hell)",
					"Raster", "Ion (dunkel)", "Filigran", "Rückblick", "Pfiff",
					"ViewMaster", "Segment (dunkel)", "Verbunden", "Bewegung",
					"Randlinie", "Austin", "Integral", "Facette")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alphabet", "Austin", "Bewegung", "Durchscheinend",
					"Herausgestellt", "Jährlich", "Kacheln", "Kontrast", "Kubistisch",
					"Modern", "Nadelstreifen", "Perspektive", "Puzzle", "Randlinie",
					"Raster", "Schlicht", "Stapel", "Traditionell", "Zeitungspapier")
				}
			}

		'en-'	{
				If($xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Austin", "Banded", "Facet", "Filigree", "Grid",
					"Integral", "Ion (Dark)", "Ion (Light)", "Motion", "Retrospect",
					"Semaphore", "Sideline", "Slice (Dark)", "Slice (Light)", "ViewMaster",
					"Whisp")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alphabet", "Annual", "Austere", "Austin", "Conservative",
					"Contrast", "Cubicles", "Exposure", "Grid", "Mod", "Motion", "Newsprint",
					"Perspective", "Pinstripes", "Puzzle", "Sideline", "Stacks", "Tiles", "Transcend")
				}
			}

		'es-'	{
				If($xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Austin", "Con bandas", "Cortar (oscuro)", "Cuadrícula", 
					"Whisp", "Faceta", "Filigrana", "Integral", "Ion (claro)", 
					"Ion (oscuro)", "Línea lateral", "Movimiento", "Retrospectiva", 
					"Semáforo", "Slice (luz)", "Vista principal", "Whisp")
				}
				ElseIf($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("Whisp", "Vista principal", "Filigrana", "Austin",
					"Slice (luz)", "Faceta", "Semáforo", "Retrospectiva", "Cuadrícula",
					"Movimiento", "Cortar (oscuro)", "Línea lateral", "Ion (oscuro)",
					"Ion (claro)", "Integral", "Con bandas")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alfabeto", "Anual", "Austero", "Austin", "Conservador",
					"Contraste", "Cuadrícula", "Cubículos", "Exposición", "Línea lateral",
					"Moderno", "Mosaicos", "Movimiento", "Papel periódico",
					"Perspectiva", "Pilas", "Puzzle", "Rayas", "Sobrepasar")
				}
			}

		'fi-'	{
				If($xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Filigraani", "Integraali", "Ioni (tumma)",
					"Ioni (vaalea)", "Opastin", "Pinta", "Retro", "Sektori (tumma)",
					"Sektori (vaalea)", "Vaihtuvavärinen", "ViewMaster", "Austin",
					"Kuiskaus", "Liike", "Ruudukko", "Sivussa")
				}
				ElseIf($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("Filigraani", "Integraali", "Ioni (tumma)",
					"Ioni (vaalea)", "Opastin", "Pinta", "Retro", "Sektori (tumma)",
					"Sektori (vaalea)", "Vaihtuvavärinen", "ViewMaster", "Austin",
					"Kiehkura", "Liike", "Ruudukko", "Sivussa")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Aakkoset", "Askeettinen", "Austin", "Kontrasti",
					"Laatikot", "Liike", "Liituraita", "Mod", "Osittain peitossa",
					"Palapeli", "Perinteinen", "Perspektiivi", "Pinot", "Ruudukko",
					"Ruudut", "Sanomalehtipaperi", "Sivussa", "Vuotuinen", "Ylitys")
				}
			}

		'fr-'	{
				If($xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016)
				{
					$xArray = ("À bandes", "Austin", "Facette", "Filigrane", 
					"Guide", "Intégrale", "Ion (clair)", "Ion (foncé)", 
					"Lignes latérales", "Quadrillage", "Rétrospective", "Secteur (clair)", 
					"Secteur (foncé)", "Sémaphore", "ViewMaster", "Whisp")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alphabet", "Annuel", "Austère", "Austin", 
					"Blocs empilés", "Classique", "Contraste", "Emplacements de bureau", 
					"Exposition", "Guide", "Ligne latérale", "Moderne", 
					"Mosaïques", "Mots croisés", "Papier journal", "Perspective",
					"Quadrillage", "Rayures fines", "Transcendant")
				}
			}

		'nb-'	{
				If($xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Austin", "BevegElse", "Dempet", "Fasett", "Filigran",
					"Integral", "Ion (lys)", "Ion (mørk)", "Retrospekt", "Rutenett",
					"Sektor (lys)", "Sektor (mørk)", "Semafor", "Sidelinje", "Stripet",
					"ViewMaster")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alfabet", "Årlig", "Avistrykk", "Austin", "Avlukker",
					"BevegElse", "Engasjement", "Enkel", "Fliser", "Konservativ",
					"Kontrast", "Mod", "Perspektiv", "Puslespill", "Rutenett", "Sidelinje",
					"Smale striper", "Stabler", "Transcenderende")
				}
			}

		'nl-'	{
				If($xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Austin", "Beweging", "Facet", "Filigraan", "Gestreept",
					"Integraal", "Ion (donker)", "Ion (licht)", "Raster",
					"Segment (Light)", "Semafoor", "Slice (donker)", "Spriet",
					"Terugblik", "Terzijde", "ViewMaster")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Aantrekkelijk", "Alfabet", "Austin", "Bescheiden",
					"Beweging", "Blikvanger", "Contrast", "Eenvoudig", "Jaarlijks",
					"Krantenpapier", "Krijtstreep", "Kubussen", "Mod", "Perspectief",
					"Puzzel", "Raster", "Stapels",
					"Tegels", "Terzijde")
				}
			}

		'pt-'	{
				If($xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Animação", "Austin", "Em Tiras", "Exibição Mestra",
					"Faceta", "Fatia (Clara)", "Fatia (Escura)", "Filete", "Filigrana", 
					"Grade", "Integral", "Íon (Claro)", "Íon (Escuro)", "Linha Lateral",
					"Retrospectiva", "Semáforo")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alfabeto", "Animação", "Anual", "Austero", "Austin", "Baias",
					"Conservador", "Contraste", "Exposição", "Grade", "Ladrilhos",
					"Linha Lateral", "Listras", "Mod", "Papel Jornal", "Perspectiva", "Pilhas",
					"Quebra-cabeça", "Transcend")
				}
			}

		'sv-'	{
				If($xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Austin", "Band", "Fasett", "Filigran", "Integrerad", "Jon (ljust)",
					"Jon (mörkt)", "Knippe", "Rutnät", "RörElse", "Sektor (ljus)", "Sektor (mörk)",
					"Semafor", "Sidlinje", "VisaHuvudsida", "Återblick")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alfabetmönster", "Austin", "Enkelt", "Exponering", "Konservativt",
					"Kontrast", "Kritstreck", "Kuber", "Perspektiv", "Plattor", "Pussel", "Rutnät",
					"RörElse", "Sidlinje", "Sobert", "Staplat", "Tidningspapper", "Årligt",
					"Övergående")
				}
			}

		'zh-'	{
				If($xWordVersion -eq $wdWord2010 -or $xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016)
				{
					$xArray = ('奥斯汀', '边线型', '花丝', '怀旧', '积分',
					'离子(浅色)', '离子(深色)', '母版型', '平面', '切片(浅色)',
					'切片(深色)', '丝状', '网格', '镶边', '信号灯',
					'运动型')
				}
			}

		Default	{
					If($xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016)
					{
						$xArray = ("Austin", "Banded", "Facet", "Filigree", "Grid",
						"Integral", "Ion (Dark)", "Ion (Light)", "Motion", "Retrospect",
						"Semaphore", "Sideline", "Slice (Dark)", "Slice (Light)", "ViewMaster",
						"Whisp")
					}
					ElseIf($xWordVersion -eq $wdWord2010)
					{
						$xArray = ("Alphabet", "Annual", "Austere", "Austin", "Conservative",
						"Contrast", "Cubicles", "Exposure", "Grid", "Mod", "Motion", "Newsprint",
						"Perspective", "Pinstripes", "Puzzle", "Sideline", "Stacks", "Tiles", "Transcend")
					}
				}
	}
	
	If($xArray -contains $xCP)
	{
		$xArray = $Null
		Return $True
	}
	Else
	{
		$xArray = $Null
		Return $False
	}
}

Function CheckWordPrereq
{
	If((Test-Path  REGISTRY::HKEY_CLASSES_ROOT\Word.Application) -eq $False)
	{
		$ErrorActionPreference = $SaveEAPreference
		
		If(($MSWord -eq $False) -and ($PDF -eq $True))
		{
			Write-Host "`n`n`t`tThis script uses Microsoft Word's SaveAs PDF function, please install Microsoft Word`n`n"
			Exit
		}
		Else
		{
			Write-Host "`n`n`t`tThis script directly outputs to Microsoft Word, please install Microsoft Word`n`n"
			Exit
		}
	}

	#find out our session (usually "1" except on TS/RDC or Citrix)
	$SessionID = (Get-Process -PID $PID).SessionId
	
	#Find out if winword is running in our session
	[bool]$wordrunning = $null –ne ((Get-Process 'WinWord' -ea 0) | Where-Object {$_.SessionId -eq $SessionID})
	If($wordrunning)
	{
		$ErrorActionPreference = $SaveEAPreference
		Write-Host "`n`n`tPlease close all instances of Microsoft Word before running this report.`n`n"
		Exit
	}
}

Function ValidateCompanyName
{
	[bool]$xResult = Test-RegistryValue "HKCU:\Software\Microsoft\Office\Common\UserInfo" "CompanyName"
	If($xResult)
	{
		Return Get-LocalRegistryValue "HKCU:\Software\Microsoft\Office\Common\UserInfo" "CompanyName"
	}
	Else
	{
		$xResult = Test-RegistryValue "HKCU:\Software\Microsoft\Office\Common\UserInfo" "Company"
		If($xResult)
		{
			Return Get-LocalRegistryValue "HKCU:\Software\Microsoft\Office\Common\UserInfo" "Company"
		}
		Else
		{
			Return ""
		}
	}
}

Function Set-DocumentProperty {
    <#
	.SYNOPSIS
	Function to set the Title Page document properties in MS Word
	.DESCRIPTION
	Long description
	.PARAMETER Document
	Current Document Object
	.PARAMETER DocProperty
	Parameter description
	.PARAMETER Value
	Parameter description
	.EXAMPLE
	Set-DocumentProperty -Document $Script:Doc -DocProperty Title -Value 'MyTitle'
	.EXAMPLE
	Set-DocumentProperty -Document $Script:Doc -DocProperty Company -Value 'MyCompany'
	.EXAMPLE
	Set-DocumentProperty -Document $Script:Doc -DocProperty Author -Value 'Jim Moyle'
	.EXAMPLE
	Set-DocumentProperty -Document $Script:Doc -DocProperty Subject -Value 'MySubjectTitle'
	.NOTES
	Function Created by Jim Moyle June 2017
	Twitter : @JimMoyle
	#>
    param (
        [object]$Document,
        [String]$DocProperty,
        [string]$Value
    )
    try {
        $binding = "System.Reflection.BindingFlags" -as [type]
        $builtInProperties = $Document.BuiltInDocumentProperties
        $property = [System.__ComObject].invokemember("item", $binding::GetProperty, $null, $BuiltinProperties, $DocProperty)
        [System.__ComObject].invokemember("value", $binding::SetProperty, $null, $property, $Value)
    }
    catch {
        Write-Warning "Failed to set $DocProperty to $Value"
    }
}

Function FindWordDocumentEnd
{
	#return focus to main document    
	$Script:Doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument
	#move to the end of the current document
	$Script:Selection.EndKey($wdStory,$wdMove) | Out-Null
}

Function SetupWord
{
	Write-Verbose "$(Get-Date -Format G): Setting up Word"
    
	If(!$AddDateTime)
	{
		[string]$Script:WordFileName = "$($Script:pwdpath)\$($OutputFileName).docx"
		If($PDF)
		{
			[string]$Script:PDFFileName = "$($Script:pwdpath)\$($OutputFileName).pdf"
		}
	}
	ElseIf($AddDateTime)
	{
		[string]$Script:WordFileName = "$($Script:pwdpath)\$($OutputFileName)_$(Get-Date -f yyyy-MM-dd_HHmm).docx"
		If($PDF)
		{
			[string]$Script:PDFFileName = "$($Script:pwdpath)\$($OutputFileName)_$(Get-Date -f yyyy-MM-dd_HHmm).pdf"
		}
	}

	# Setup word for output
	Write-Verbose "$(Get-Date -Format G): Create Word comObject."
	$Script:Word = New-Object -comobject "Word.Application" -EA 0 4>$Null
	
	If(!$? -or $Null -eq $Script:Word)
	{
		Write-Warning "The Word object could not be created. You may need to repair your Word installation."
		$ErrorActionPreference = $SaveEAPreference
		Write-Error "
		`n`n
		`t`t
		The Word object could not be created. You may need to repair your Word installation.
		`n`n
		`t`t
		Script cannot Continue.
		`n`n"
		Exit
	}

	Write-Verbose "$(Get-Date -Format G): Determine Word language value"
	If( ( validStateProp $Script:Word Language Value__ ) )
	{
		[int]$Script:WordLanguageValue = [int]$Script:Word.Language.Value__
	}
	Else
	{
		[int]$Script:WordLanguageValue = [int]$Script:Word.Language
	}

	If(!($Script:WordLanguageValue -gt -1))
	{
		$ErrorActionPreference = $SaveEAPreference
		Write-Error "
		`n`n
		`t`t
		Unable to determine the Word language value. You may need to repair your Word installation.
		`n`n
		`t`t
		Script cannot Continue.
		`n`n
		"
		AbortScript
	}
	Write-Verbose "$(Get-Date -Format G): Word language value is $($Script:WordLanguageValue)"
	
	$Script:WordCultureCode = GetCulture $Script:WordLanguageValue
	
	SetWordHashTable $Script:WordCultureCode
	
	[int]$Script:WordVersion = [int]$Script:Word.Version
	If($Script:WordVersion -eq $wdWord2016)
	{
		$Script:WordProduct = "Word 2016"
	}
	ElseIf($Script:WordVersion -eq $wdWord2013)
	{
		$Script:WordProduct = "Word 2013"
	}
	ElseIf($Script:WordVersion -eq $wdWord2010)
	{
		$Script:WordProduct = "Word 2010"
	}
	ElseIf($Script:WordVersion -eq $wdWord2007)
	{
		$ErrorActionPreference = $SaveEAPreference
		Write-Error "
		`n`n
		`t`t
		Microsoft Word 2007 is no longer supported.`n`n`t`tScript will end.
		`n`n
		"
		AbortScript
	}
	ElseIf($Script:WordVersion -eq 0)
	{
		Write-Error "
		`n`n
		`t`t
		The Word Version is 0. You should run a full online repair of your Office installation.
		`n`n
		`t`t
		Script cannot Continue.
		`n`n
		"
		Exit
	}
	Else
	{
		$ErrorActionPreference = $SaveEAPreference
		Write-Error "
		`n`n
		`t`t
		You are running an untested or unsupported version of Microsoft Word.
		`n`n
		`t`tScript will end.`n`n`t`tPlease send info on your version of Word to webster@carlwebster.com
		`n`n
		"
		AbortScript
	}

	#only validate CompanyName if the field is blank
	If([String]::IsNullOrEmpty($CompanyName))
	{
		Write-Verbose "$(Get-Date -Format G): Company name is blank. Retrieve company name from registry."
		$TmpName = ValidateCompanyName
		
		If([String]::IsNullOrEmpty($TmpName))
		{
			Write-Host "
		Company Name is blank so Cover Page will not show a Company Name.
		Check HKCU:\Software\Microsoft\Office\Common\UserInfo for Company or CompanyName value.
		You may want to use the -CompanyName parameter if you need a Company Name on the cover page.
			" -Foreground White
			$Script:CoName = $TmpName
		}
		Else
		{
			$Script:CoName = $TmpName
			Write-Verbose "$(Get-Date -Format G): Updated company name to $($Script:CoName)"
		}
	}
	Else
	{
		$Script:CoName = $CompanyName
	}

	If($Script:WordCultureCode -ne "en-")
	{
		Write-Verbose "$(Get-Date -Format G): Check Default Cover Page for $($WordCultureCode)"
		[bool]$CPChanged = $False
		Switch ($Script:WordCultureCode)
		{
			'ca-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Línia lateral"
						$CPChanged = $True
					}
				}

			'da-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Sidelinje"
						$CPChanged = $True
					}
				}

			'de-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Randlinie"
						$CPChanged = $True
					}
				}

			'es-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Línea lateral"
						$CPChanged = $True
					}
				}

			'fi-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Sivussa"
						$CPChanged = $True
					}
				}

			'fr-'	{
					If($CoverPage -eq "Sideline")
					{
						If($Script:WordVersion -eq $wdWord2013 -or $Script:WordVersion -eq $wdWord2016)
						{
							$CoverPage = "Lignes latérales"
							$CPChanged = $True
						}
						Else
						{
							$CoverPage = "Ligne latérale"
							$CPChanged = $True
						}
					}
				}

			'nb-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Sidelinje"
						$CPChanged = $True
					}
				}

			'nl-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Terzijde"
						$CPChanged = $True
					}
				}

			'pt-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Linha Lateral"
						$CPChanged = $True
					}
				}

			'sv-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Sidlinje"
						$CPChanged = $True
					}
				}

			'zh-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "边线型"
						$CPChanged = $True
					}
				}
		}

		If($CPChanged)
		{
			Write-Verbose "$(Get-Date -Format G): Changed Default Cover Page from Sideline to $($CoverPage)"
		}
	}

	Write-Verbose "$(Get-Date -Format G): Validate cover page $($CoverPage) for culture code $($Script:WordCultureCode)"
	[bool]$ValidCP = $False
	
	$ValidCP = ValidateCoverPage $Script:WordVersion $CoverPage $Script:WordCultureCode
	
	If(!$ValidCP)
	{
		$ErrorActionPreference = $SaveEAPreference
		Write-Verbose "$(Get-Date -Format G): Word language value $($Script:WordLanguageValue)"
		Write-Verbose "$(Get-Date -Format G): Culture code $($Script:WordCultureCode)"
		Write-Error "
		`n`n
		`t`t
		For $($Script:WordProduct), $($CoverPage) is not a valid Cover Page option.
		`n`n
		`t`t
		Script cannot Continue.
		`n`n
		"
		AbortScript
	}

	$Script:Word.Visible = $False

	#http://jdhitsolutions.com/blog/2012/05/san-diego-2012-powershell-deep-dive-slides-and-demos/
	#using Jeff's Demo-WordReport.ps1 file for examples
	Write-Verbose "$(Get-Date -Format G): Load Word Templates"

	[bool]$Script:CoverPagesExist = $False
	[bool]$BuildingBlocksExist = $False

	$Script:Word.Templates.LoadBuildingBlocks()
	#word 2010/2013/2016
	$BuildingBlocksCollection = $Script:Word.Templates | Where-Object{$_.name -eq "Built-In Building Blocks.dotx"}

	Write-Verbose "$(Get-Date -Format G): Attempt to load cover page $($CoverPage)"
	$part = $Null

	$BuildingBlocksCollection | 
	ForEach-Object {
		If ($_.BuildingBlockEntries.Item($CoverPage).Name -eq $CoverPage) 
		{
			$BuildingBlocks = $_
		}
	}        

	If($Null -ne $BuildingBlocks)
	{
		$BuildingBlocksExist = $True

		Try 
		{
			$part = $BuildingBlocks.BuildingBlockEntries.Item($CoverPage)
		}

		Catch
		{
			$part = $Null
		}

		If($Null -ne $part)
		{
			$Script:CoverPagesExist = $True
		}
	}

	If(!$Script:CoverPagesExist)
	{
		Write-Verbose "$(Get-Date -Format G): Cover Pages are not installed or the Cover Page $($CoverPage) does not exist."
		Write-Host "Cover Pages are not installed or the Cover Page $($CoverPage) does not exist." -Foreground White
		Write-Host "This report will not have a Cover Page." -Foreground White
	}

	Write-Verbose "$(Get-Date -Format G): Create empty word doc"
	$Script:Doc = $Script:Word.Documents.Add()
	If($Null -eq $Script:Doc)
	{
		Write-Verbose "$(Get-Date -Format G): "
		$ErrorActionPreference = $SaveEAPreference
		Write-Error "
		`n`n
	An empty Word document could not be created. You may need to repair your Word installation.
		`n`n
	Script cannot Continue.
		`n`n"
		AbortScript
	}

	$Script:Selection = $Script:Word.Selection
	If($Null -eq $Script:Selection)
	{
		Write-Verbose "$(Get-Date -Format G): "
		$ErrorActionPreference = $SaveEAPreference
		Write-Error "
		`n`n
	An unknown error happened selecting the entire Word document for default formatting options.
		`n`n
	Script cannot Continue.
		`n`n"
		AbortScript
	}

	#set Default tab stops to 1/2 inch (this line is not from Jeff Hicks)
	#36 = .50"
	$Script:Word.ActiveDocument.DefaultTabStop = 36

	#Disable Spell and Grammar Check to resolve issue and improve performance (from Pat Coughlin)
	Write-Verbose "$(Get-Date -Format G): Disable grammar and spell checking"
	#bug reported 1-Apr-2014 by Tim Mangan
	#save current options first before turning them off
	$Script:CurrentGrammarOption = $Script:Word.Options.CheckGrammarAsYouType
	$Script:CurrentSpellingOption = $Script:Word.Options.CheckSpellingAsYouType
	$Script:Word.Options.CheckGrammarAsYouType = $False
	$Script:Word.Options.CheckSpellingAsYouType = $False

	If($BuildingBlocksExist)
	{
		#insert new page, getting ready for table of contents
		Write-Verbose "$(Get-Date -Format G): Insert new page, getting ready for table of contents"
		$part.Insert($Script:Selection.Range,$True) | Out-Null
		$Script:Selection.InsertNewPage()

		#table of contents
		Write-Verbose "$(Get-Date -Format G): Table of Contents - $($Script:MyHash.Word_TableOfContents)"
		$toc = $BuildingBlocks.BuildingBlockEntries.Item($Script:MyHash.Word_TableOfContents)
		If($Null -eq $toc)
		{
			Write-Verbose "$(Get-Date -Format G): "
			Write-Host "Table of Content - $($Script:MyHash.Word_TableOfContents) could not be retrieved." -Foreground White
			Write-Host "This report will not have a Table of Contents." -Foreground White
		}
		Else
		{
			$toc.insert($Script:Selection.Range,$True) | Out-Null
		}
	}
	Else
	{
		Write-Host "Table of Contents are not installed." -Foreground White
		Write-Host "Table of Contents are not installed so this report will not have a Table of Contents." -Foreground White
	}

	#set the footer
	Write-Verbose "$(Get-Date -Format G): Set the footer"
	[string]$footertext = "Report created by $username"

	#get the footer
	Write-Verbose "$(Get-Date -Format G): Get the footer and format font"
	$Script:Doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekPrimaryFooter
	#get the footer and format font
	$footers = $Script:Doc.Sections.Last.Footers
	ForEach ($footer in $footers) 
	{
		If($footer.exists) 
		{
			$footer.range.Font.name = "Calibri"
			$footer.range.Font.size = 8
			$footer.range.Font.Italic = $True
			$footer.range.Font.Bold = $True
		}
	} #end ForEach
	Write-Verbose "$(Get-Date -Format G): Footer text"
	$Script:Selection.HeaderFooter.Range.Text = $footerText

	#add page numbering
	Write-Verbose "$(Get-Date -Format G): Add page numbering"
	$Script:Selection.HeaderFooter.PageNumbers.Add($wdAlignPageNumberRight) | Out-Null

	FindWordDocumentEnd
	Write-Verbose "$(Get-Date -Format G):"
	#end of Jeff Hicks 
}

Function UpdateDocumentProperties
{
	Param([string]$AbstractTitle, [string]$SubjectTitle)
	#updated 12-Nov-2017 with additional cover page fields
	#Update document properties
	If($MSWORD -or $PDF)
	{
		If($Script:CoverPagesExist)
		{
			Write-Verbose "$(Get-Date -Format G): Set Cover Page Properties"
			#8-Jun-2017 put these 4 items in alpha order
            Set-DocumentProperty -Document $Script:Doc -DocProperty Author -Value $UserName
            Set-DocumentProperty -Document $Script:Doc -DocProperty Company -Value $Script:CoName
            Set-DocumentProperty -Document $Script:Doc -DocProperty Subject -Value $SubjectTitle
            Set-DocumentProperty -Document $Script:Doc -DocProperty Title -Value $Script:title

			#Get the Coverpage XML part
			$cp = $Script:Doc.CustomXMLParts | Where-Object {$_.NamespaceURI -match "coverPageProps$"}

			#get the abstract XML part
			$ab = $cp.documentelement.ChildNodes | Where-Object {$_.basename -eq "Abstract"}
			#set the text
			If([String]::IsNullOrEmpty($Script:CoName))
			{
				[string]$abstract = $AbstractTitle
			}
			Else
			{
				[string]$abstract = "$($AbstractTitle) for $($Script:CoName)"
			}
			$ab.Text = $abstract

			#added 8-Jun-2017
			$ab = $cp.documentelement.ChildNodes | Where-Object {$_.basename -eq "CompanyAddress"}
			#set the text
			[string]$abstract = $CompanyAddress
			$ab.Text = $abstract

			#added 8-Jun-2017
			$ab = $cp.documentelement.ChildNodes | Where-Object {$_.basename -eq "CompanyEmail"}
			#set the text
			[string]$abstract = $CompanyEmail
			$ab.Text = $abstract

			#added 8-Jun-2017
			$ab = $cp.documentelement.ChildNodes | Where-Object {$_.basename -eq "CompanyFax"}
			#set the text
			[string]$abstract = $CompanyFax
			$ab.Text = $abstract

			#added 8-Jun-2017
			$ab = $cp.documentelement.ChildNodes | Where-Object {$_.basename -eq "CompanyPhone"}
			#set the text
			[string]$abstract = $CompanyPhone
			$ab.Text = $abstract

			$ab = $cp.documentelement.ChildNodes | Where-Object {$_.basename -eq "PublishDate"}
			#set the text
			[string]$abstract = (Get-Date -Format d).ToString()
			$ab.Text = $abstract

			Write-Verbose "$(Get-Date -Format G): Update the Table of Contents"
			#update the Table of Contents
			$Script:Doc.TablesOfContents.item(1).Update()
			$cp = $Null
			$ab = $Null
			$abstract = $Null
		}
	}
}
#endregion

#region registry functions
#http://stackoverflow.com/questions/5648931/test-if-registry-value-exists
# This Function just gets $True or $False
Function Test-RegistryValue($path, $name)
{
	$key = Get-Item -LiteralPath $path -EA 0
	$key -and $Null -ne $key.GetValue($name, $Null)
}

# Gets the specified local registry value or $Null if it is missing
Function Get-LocalRegistryValue($path, $name)
{
	$key = Get-Item -LiteralPath $path -EA 0
	If($key)
	{
		$key.GetValue($name, $Null)
	}
	Else
	{
		$Null
	}
}

Function Get-RegistryValue
{
	# Gets the specified registry value or $Null if it is missing
	[CmdletBinding()]
	Param([string]$path, [string]$name, [string]$ComputerName)
	If($ComputerName -eq $env:computername -or $ComputerName -eq "LocalHost")
	{
		$key = Get-Item -LiteralPath $path -EA 0
		If($key)
		{
			Return $key.GetValue($name, $Null)
		}
		Else
		{
			Return $Null
		}
	}
	Else
	{
		#path needed here is different for remote registry access
		$path1 = $path.SubString(6)
		$path2 = $path1.Replace('\','\\')
		$Reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $ComputerName)
		$RegKey= $Reg.OpenSubKey($path2)
		$Results = $RegKey.GetValue($name)
		If($Null -ne $Results)
		{
			Return $Results
		}
		Else
		{
			Return $Null
		}
	}
}
#endregion

#region word, text and html line output functions
Function line
#function created by Michael B. Smith, Exchange MVP
#@essentialexch on Twitter
#https://essential.exchange/blog
#for creating the formatted text report
#created March 2011
#updated March 2014
# updated March 2019 to use StringBuilder (about 100 times more efficient than simple strings)
{
	Param
	(
		[Int]    $tabs = 0, 
		[String] $name = '', 
		[String] $value = '', 
		[String] $newline = [System.Environment]::NewLine, 
		[Switch] $nonewline
	)

	while( $tabs -gt 0 )
	{
		#V1.17 - switch to using a StringBuilder for $global:Output
		$null = $global:Output.Append( "`t" )
		$tabs--
	}

	If( $nonewline )
	{
		#V1.17 - switch to using a StringBuilder for $global:Output
		$null = $global:Output.Append( $name + $value )
	}
	Else
	{
		#V1.17 - switch to using a StringBuilder for $global:Output
		$null = $global:Output.AppendLine( $name + $value )
	}
}
	
Function WriteWordLine
#Function created by Ryan Revord
#@rsrevord on Twitter
#Function created to make output to Word easy in this script
#updated 27-Mar-2014 to include font name, font size, italics and bold options
{
	Param([int]$style=0, 
	[int]$tabs = 0, 
	[string]$name = '', 
	[string]$value = '', 
	[string]$fontName=$Null,
	[int]$fontSize=0,
	[bool]$italics=$False,
	[bool]$boldface=$False,
	[Switch]$nonewline)
	
	#Build output style
	[string]$output = ""
	Switch ($style)
	{
		0 {$Script:Selection.Style = $Script:MyHash.Word_NoSpacing; Break}
		1 {$Script:Selection.Style = $Script:MyHash.Word_Heading1; Break}
		2 {$Script:Selection.Style = $Script:MyHash.Word_Heading2; Break}
		3 {$Script:Selection.Style = $Script:MyHash.Word_Heading3; Break}
		4 {$Script:Selection.Style = $Script:MyHash.Word_Heading4; Break}
		Default {$Script:Selection.Style = $Script:MyHash.Word_NoSpacing; Break}
	}
	
	#build # of tabs
	While($tabs -gt 0)
	{ 
		$output += "`t"; $tabs--; 
	}
 
	If(![String]::IsNullOrEmpty($fontName)) 
	{
		$Script:Selection.Font.name = $fontName
	} 

	If($fontSize -ne 0) 
	{
		$Script:Selection.Font.size = $fontSize
	} 
 
	If($italics -eq $True) 
	{
		$Script:Selection.Font.Italic = $True
	} 
 
	If($boldface -eq $True) 
	{
		$Script:Selection.Font.Bold = $True
	} 

	#output the rest of the parameters.
	$output += $name + $value
	$Script:Selection.TypeText($output)
 
	#test for new WriteWordLine 0.
	If($nonewline)
	{
		# Do nothing.
	} 
	Else 
	{
		$Script:Selection.TypeParagraph()
	}
}

#***********************************************************************************************************
# WriteHTMLLine
#***********************************************************************************************************

<#
.Synopsis
	Writes a line of output for HTML output
.DESCRIPTION
	This Function formats an HTML line
.USAGE
	WriteHTMLLine <Style> <Tabs> <Name> <Value> <Font Name> <Font Size> <Options>

	0 for Font Size denotes using the default font size of 2 or 10 point

.EXAMPLE
	WriteHTMLLine 0 0 " "

	Writes a blank line with no style or tab stops, obviously none needed.

.EXAMPLE
	WriteHTMLLine 0 1 "This is a regular line of text indented 1 tab stops"

	Writes a line with 1 tab stop.

.EXAMPLE
	WriteHTMLLine 0 0 "This is a regular line of text in the default font in italics" "" $Null 0 $htmlitalics

	Writes a line omitting font and font size and setting the italics attribute

.EXAMPLE
	WriteHTMLLine 0 0 "This is a regular line of text in the default font in bold" "" $Null 0 $htmlBold

	Writes a line omitting font and font size and setting the bold attribute

.EXAMPLE
	WriteHTMLLine 0 0 "This is a regular line of text in the default font in bold italics" "" $Null 0 ($htmlBold -bor $htmlitalics)

	Writes a line omitting font and font size and setting both italics and bold options

.EXAMPLE	
	WriteHTMLLine 0 0 "This is a regular line of text in the default font in 10 point" "" $Null 2  # 10 point font

	Writes a line using 10 point font

.EXAMPLE
	WriteHTMLLine 0 0 "This is a regular line of text in Courier New font" "" "Courier New" 0 

	Writes a line using Courier New Font and 0 font point size (default = 2 if set to 0)

.EXAMPLE	
	WriteHTMLLine 0 0 "This is a regular line of RED text indented 0 tab stops with the computer name as data in 10 point Courier New bold italics: " $env:computername "Courier New" 2 ($htmlBold -bor $htmlred -bor $htmlitalics)

	Writes a line using Courier New Font with first and second string values to be used, also uses 10 point font with bold, italics and red color options set.

.NOTES

	Font Size - Unlike word, there is a limited set of font sizes that can be used in HTML.  They are:
		0 - default which actually gives it a 2 or 10 point.
		1 - 7.5 point font size
		2 - 10 point
		3 - 13.5 point
		4 - 15 point
		5 - 18 point
		6 - 24 point
		7 - 36 point
	Any number larger than 7 defaults to 7

	Style - Refers to the headers that are used with output and resemble the headers in word, 
	HTML supports headers h1-h6 and h1-h4 are more commonly used.  Unlike word, H1 will not 
	give you a blue colored font, you will have to set that yourself.

	Colors and Bold/Italics Flags are:

		htmlbold       
		htmlitalics    
		htmlred        
		htmlcyan        
		htmlblue       
		htmldarkblue   
		htmllightblue   
		htmlpurple      
		htmlyellow      
		htmllime       
		htmlmagenta     
		htmlwhite       
		htmlsilver      
		htmlgray       
		htmlolive       
		htmlorange      
		htmlmaroon      
		htmlgreen       
		htmlblack       
#>

#V3.00
# to suppress $crlf in HTML documents, replace this with '' (empty string)
# but this was added to make the HTML readable
$crlf = [System.Environment]::NewLine

Function WriteHTMLLine
#Function created by Ken Avram
#Function created to make output to HTML easy in this script
#headings fixed 12-Oct-2016 by Webster
#errors with $HTMLStyle fixed 7-Dec-2017 by Webster
# re-implemented/re-based for v3.00 by Michael B. Smith
{
	Param
	(
		[Int]    $style    = 0, 
		[Int]    $tabs     = 0, 
		[String] $name     = '', 
		[String] $value    = '', 
		[String] $fontName = $null,
		[Int]    $fontSize = 1,
		[Int]    $options  = $htmlblack
	)

	#V3.00 - FIXME - long story short, this Function was wrong and had been wrong for a long time. 
	## The Function generated invalid HTML, and ignored fontname and fontsize parameters. I fixed
	## those items, but that made the report unreadable, because all of the formatting had been based
	## on this Function not working properly.

	## here is a typical H1 previously generated:
	## <h1>///&nbsp;&nbsp;Forest Information&nbsp;&nbsp;\\\<font face='Calibri' color='#000000' size='1'></h1></font>

	## fixing the Function generated this (unreadably small):
	## <h1><font face='Calibri' color='#000000' size='1'>///&nbsp;&nbsp;Forest Information&nbsp;&nbsp;\\\</font></h1>

	## So I took all the fixes out. This routine now generates valid HTML, but the fontName, fontSize,
	## and options parameters are ignored; so the routine generates equivalent output as before. I took
	## the fixes out instead of fixing all the call sites, because there are 225 call sites! If you are
	## willing to update all the call sites, you can easily re-instate the fixes. They have only been
	## commented out with '##' below.

	## If( [String]::IsNullOrEmpty( $fontName ) )
	## {
	##	$fontName = 'Calibri'
	## }
	## If( $fontSize -le 0 )
	## {
	##	$fontSize = 1
	## }

	## ## output data is stored here
	## [String] $output = ''
	[System.Text.StringBuilder] $sb = New-Object System.Text.StringBuilder( 1024 )

	If( [String]::IsNullOrEmpty( $name ) )	
	{
		## $HTMLBody = '<p></p>'
		$null = $sb.Append( '<p></p>' )
	}
	Else
	{
		## #V3.00
		[Bool] $ital = $options -band $htmlitalics
		[Bool] $bold = $options -band $htmlBold
		## $color = $global:htmlColor[ $options -band 0xffffc ]

		## ## build the HTML output string
##		$HTMLBody = ''
##		If( $ital ) { $HTMLBody += '<i>' }
##		If( $bold ) { $HTMLBody += '<b>' } 
		If( $ital ) { $null = $sb.Append( '<i>' ) }
		If( $bold ) { $null = $sb.Append( '<b>' ) } 

		Switch( $style )
		{
			1 { $HTMLOpen = '<h1>'; $HTMLClose = '</h1>'; Break }
			2 { $HTMLOpen = '<h2>'; $HTMLClose = '</h2>'; Break }
			3 { $HTMLOpen = '<h3>'; $HTMLClose = '</h3>'; Break }
			4 { $HTMLOpen = '<h4>'; $HTMLClose = '</h4>'; Break }
			Default { $HTMLOpen = ''; $HTMLClose = ''; Break }
		}

		## $HTMLBody += $HTMLOpen
		$null = $sb.Append( $HTMLOpen )

		## If($HTMLClose -eq '')
		## {
		##	$HTMLBody += "<br><font face='" + $fontName + "' " + "color='" + $color + "' size='"  + $fontSize + "'>"
		## }
		## Else
		## {
		##	$HTMLBody += "<font face='" + $fontName + "' " + "color='" + $color + "' size='"  + $fontSize + "'>"
		## }
		
##		While( $tabs -gt 0 )
##		{ 
##			$output += '&nbsp;&nbsp;&nbsp;&nbsp;'
##			$tabs--
##		}
		## output the rest of the parameters.
##		$output += $name + $value
		## $HTMLBody += $output
		$null = $sb.Append( ( '&nbsp;&nbsp;&nbsp;&nbsp;' * $tabs ) + $name + $value )

		## $HTMLBody += '</font>'
##		If( $HTMLClose -eq '' ) { $HTMLBody += '<br>'     }
##		Else                    { $HTMLBody += $HTMLClose }

##		If( $ital ) { $HTMLBody += '</i>' }
##		If( $bold ) { $HTMLBody += '</b>' } 

##		If( $HTMLClose -eq '' ) { $HTMLBody += '<br />' }

		If( $HTMLClose -eq '' ) { $null = $sb.Append( '<br>' )     }
		Else                    { $null = $sb.Append( $HTMLClose ) }

		If( $ital ) { $null = $sb.Append( '</i>' ) }
		If( $bold ) { $null = $sb.Append( '</b>' ) } 

		If( $HTMLClose -eq '' ) { $null = $sb.Append( '<br />' ) }
	}
	##$HTMLBody += $crlf
	$null = $sb.AppendLine( '' )

	Out-File -FilePath $Script:HTMLFileName -Append -InputObject $sb.ToString() 4>$Null
}
#endregion

#region HTML table functions
#***********************************************************************************************************
# AddHTMLTable - Called from FormatHTMLTable Function
# Created by Ken Avram
# modified by Jake Rutski
# re-implemented by Michael B. Smith for v2.00. Also made the documentation match reality.
#***********************************************************************************************************
Function AddHTMLTable
{
	Param
	(
		[String]   $fontName  = 'Calibri',
		[Int]      $fontSize  = 2,
		[Int]      $colCount  = 0,
		[Int]      $rowCount  = 0,
		[Object[]] $rowInfo   = $null,
		[Object[]] $fixedInfo = $null
	)
	#V3.00 - Use StringBuilder - MBS
	## In the normal case, tables are only a few dozen cells. But in the case
	## of Sites, OUs, and Users - there may be many hundreds of thousands of 
	## cells. Using normal strings is too slow.

	#V3.00
	## If( $ExtraSpecialVerbose )
	## {
	##	$global:rowInfo1 = $rowInfo
	## }
<#
	If( $SuperVerbose )
	{
		wv "AddHTMLTable: fontName '$fontName', fontsize $fontSize, colCount $colCount, rowCount $rowCount"
		If( $null -ne $rowInfo -and $rowInfo.Count -gt 0 )
		{
			wv "AddHTMLTable: rowInfo has $( $rowInfo.Count ) elements"
			If( $ExtraSpecialVerbose )
			{
				wv "AddHTMLTable: rowInfo length $( $rowInfo.Length )"
				For( $ii = 0; $ii -lt $rowInfo.Length; $ii++ )
				{
					$row = $rowInfo[ $ii ]
					wv "AddHTMLTable: index $ii, type $( $row.GetType().FullName ), length $( $row.Length )"
					For( $yyy = 0; $yyy -lt $row.Length; $yyy++ )
					{
						wv "AddHTMLTable: index $ii, yyy = $yyy, val = '$( $row[ $yyy ] )'"
					}
					wv "AddHTMLTable: done"
				}
			}
		}
		Else
		{
			wv "AddHTMLTable: rowInfo is empty"
		}
		If( $null -ne $fixedInfo -and $fixedInfo.Count -gt 0 )
		{
			wv "AddHTMLTable: fixedInfo has $( $fixedInfo.Count ) elements"
		}
		Else
		{
			wv "AddHTMLTable: fixedInfo is empty"
		}
	}
#>

	$fwLength = If( $null -ne $fixedInfo ) { $fixedInfo.Count } else { 0 }

	##$htmlbody = ''
	[System.Text.StringBuilder] $sb = New-Object System.Text.StringBuilder( 8192 )

	If( $rowInfo -and $rowInfo.Length -lt $rowCount )
	{
##		$oldCount = $rowCount
		$rowCount = $rowInfo.Length
##		If( $SuperVerbose )
##		{
##			wv "AddHTMLTable: updated rowCount to $rowCount from $oldCount, based on rowInfo.Length"
##		}
	}

	For( $rowCountIndex = 0; $rowCountIndex -lt $rowCount; $rowCountIndex++ )
	{
		$null = $sb.AppendLine( '<tr>' )
		## $htmlbody += '<tr>'
		## $htmlbody += $crlf #V3.00 - make the HTML readable

		## each row of rowInfo is an array
		## each row consists of tuples: an item of text followed by an item of formatting data
<#		
		$row = $rowInfo[ $rowCountIndex ]
		If( $ExtraSpecialVerbose )
		{
			wv "!!!!! AddHTMLTable: rowCountIndex = $rowCountIndex, row.Length = $( $row.Length ), row gettype = $( $row.GetType().FullName )"
			wv "!!!!! AddHTMLTable: colCount $colCount"
			wv "!!!!! AddHTMLTable: row[0].Length $( $row[0].Length )"
			wv "!!!!! AddHTMLTable: row[0].GetType $( $row[0].GetType().FullName )"
			$subRow = $row
			If( $subRow -is [Array] -and $subRow[ 0 ] -is [Array] )
			{
				$subRow = $subRow[ 0 ]
				wv "!!!!! AddHTMLTable: deref subRow.Length $( $subRow.Length ), subRow.GetType $( $subRow.GetType().FullName )"
			}

			For( $columnIndex = 0; $columnIndex -lt $subRow.Length; $columnIndex += 2 )
			{
				$item = $subRow[ $columnIndex ]
				wv "!!!!! AddHTMLTable: item.GetType $( $item.GetType().FullName )"
				## If( !( $item -is [String] ) -and $item -is [Array] )
##				If( $item -is [Array] -and $item[ 0 ] -is [Array] )				
##				{
##					$item = $item[ 0 ]
##					wv "!!!!! AddHTMLTable: dereferenced item.GetType $( $item.GetType().FullName )"
##				}
				wv "!!!!! AddHTMLTable: rowCountIndex = $rowCountIndex, columnIndex = $columnIndex, val '$item'"
			}
			wv "!!!!! AddHTMLTable: done"
		}
#>

		## reset
		$row = $rowInfo[ $rowCountIndex ]

		$subRow = $row
		If( $subRow -is [Array] -and $subRow[ 0 ] -is [Array] )
		{
			$subRow = $subRow[ 0 ]
			## wv "***** AddHTMLTable: deref rowCountIndex $rowCountIndex, subRow.Length $( $subRow.Length ), subRow.GetType $( $subRow.GetType().FullName )"
		}

		$subRowLength = $subRow.Count
		For( $columnIndex = 0; $columnIndex -lt $colCount; $columnIndex += 2 )
		{
			$item = If( $columnIndex -lt $subRowLength ) { $subRow[ $columnIndex ] } Else { 0 }
			## If( !( $item -is [String] ) -and $item -is [Array] )
##			If( $item -is [Array] -and $item[ 0 ] -is [Array] )
##			{
##				$item = $item[ 0 ]
##			}

			$text   = If( $item ) { $item.ToString() } Else { '' }
			$format = If( ( $columnIndex + 1 ) -lt $subRowLength ) { $subRow[ $columnIndex + 1 ] } Else { 0 }
			## item, text, and format ALWAYS have values, even if empty values
			$color  = $global:htmlColor[ $format -band 0xffffc ]
			[Bool] $bold = $format -band $htmlBold
			[Bool] $ital = $format -band $htmlitalics
<#			
			If( $ExtraSpecialVerbose )
			{
				wv "***** columnIndex $columnIndex, subRow.Length $( $subRow.Length ), item GetType $( $item.GetType().FullName ), item '$item'"
				wv "***** format $format, color $color, text '$text'"
				wv "***** format gettype $( $format.GetType().Fullname ), text gettype $( $text.GetType().Fullname )"
			}
#>

			If( $fwLength -eq 0 )
			{
				$null = $sb.Append( "<td style=""background-color:$( $color )""><font face='$( $fontName )' size='$( $fontSize )'>" )
				##$htmlbody += "<td style=""background-color:$( $color )""><font face='$( $fontName )' size='$( $fontSize )'>"
			}
			Else
			{
				$null = $sb.Append( "<td style=""width:$( $fixedInfo[ $columnIndex / 2 ] ); background-color:$( $color )""><font face='$( $fontName )' size='$( $fontSize )'>" )
				##$htmlbody += "<td style=""width:$( $fixedInfo[ $columnIndex / 2 ] ); background-color:$( $color )""><font face='$( $fontName )' size='$( $fontSize )'>"
			}

			##If( $bold ) { $htmlbody += '<b>' }
			##If( $ital ) { $htmlbody += '<i>' }
			If( $bold ) { $null = $sb.Append( '<b>' ) }
			If( $ital ) { $null = $sb.Append( '<i>' ) }

			If( $text -eq ' ' -or $text.length -eq 0)
			{
				##$htmlbody += '&nbsp;&nbsp;&nbsp;'
				$null = $sb.Append( '&nbsp;&nbsp;&nbsp;' )
			}
			Else
			{
				For($inx = 0; $inx -lt $text.length; $inx++ )
				{
					If( $text[ $inx ] -eq ' ' )
					{
						##$htmlbody += '&nbsp;'
						$null = $sb.Append( '&nbsp;' )
					}
					Else
					{
						Break
					}
				}
				##$htmlbody += $text
				$null = $sb.Append( $text )
			}

##			If( $bold ) { $htmlbody += '</b>' }
##			If( $ital ) { $htmlbody += '</i>' }
			If( $bold ) { $null = $sb.Append( '</b>' ) }
			If( $ital ) { $null = $sb.Append( '</i>' ) }

			$null = $sb.AppendLine( '</font></td>' )
##			$htmlbody += '</font></td>'
##			$htmlbody += $crlf
		}

		$null = $sb.AppendLine( '</tr>' )
##		$htmlbody += '</tr>'
##		$htmlbody += $crlf
	}

##	If( $ExtraSpecialVerbose )
##	{
##		$global:rowInfo = $rowInfo
##		wv "!!!!! AddHTMLTable: HTML = '$htmlbody'"
##	}

	Out-File -FilePath $Script:HTMLFileName -Append -InputObject $sb.ToString() 4>$Null 
}

#***********************************************************************************************************
# FormatHTMLTable 
# Created by Ken Avram
# modified by Jake Rutski
# reworked by Michael B. Smith for v2.23
#***********************************************************************************************************

<#
.Synopsis
	Format table for a HTML output document.
.DESCRIPTION
	This function formats a table for HTML from multiple arrays of strings.
.PARAMETER noBorder
	If set to $true, a table will be generated without a border (border = '0'). Otherwise the table will be generated
	with a border (border = '1').
.PARAMETER noHeadCols
	This parameter should be used when generating tables which do not have a separate array containing column headers
	(columnArray is not specified). Set this parameter equal to the number of columns in the table.
.PARAMETER rowArray
	This parameter contains the row data array for the table.
.PARAMETER columnArray
	This parameter contains column header data for the table.
.PARAMETER fixedWidth
	This parameter contains widths for columns in pixel format ("100px") to override auto column widths
	The variable should contain a width for each column you wish to override the auto-size setting
	For example: $fixedWidth = @("100px","110px","120px","130px","140px")
.PARAMETER tableHeader
	A string containing the header for the table (printed at the top of the table, left justified). The
	default is a blank string.
.PARAMETER tableWidth
	The width of the table in pixels, or 'auto'. The default is 'auto'.
.PARAMETER fontName
	The name of the font to use in the table. The default is 'Calibri'.
.PARAMETER fontSize
	The size of the font to use in the table. The default is 2. Note that this is the HTML size, not the pixel size.

.USAGE
	FormatHTMLTable <Table Header> <Table Width> <Font Name> <Font Size>

.EXAMPLE
	FormatHTMLTable "Table Heading" "auto" "Calibri" 3

	This example formats a table and writes it out into an html file.  All of the parameters are optional
	defaults are used if not supplied.

	for <Table format>, the default is auto which will autofit the text into the columns and adjust to the longest text in that column.  You can also use percentage i.e. 25%
	which will take only 25% of the line and will auto word wrap the text to the next line in the column.  Also, instead of using a percentage, you can use pixels i.e. 400px.

	FormatHTMLTable "Table Heading" "auto" -rowArray $rowData -columnArray $columnData

	This example creates an HTML table with a heading of 'Table Heading', auto column spacing, column header data from $columnData and row data from $rowData

	FormatHTMLTable "Table Heading" -rowArray $rowData -noHeadCols 3

	This example creates an HTML table with a heading of 'Table Heading', auto column spacing, no header, and row data from $rowData

	FormatHTMLTable "Table Heading" -rowArray $rowData -fixedWidth $fixedColumns

	This example creates an HTML table with a heading of 'Table Heading, no header, row data from $rowData, and fixed columns defined by $fixedColumns

.NOTES
	In order to use the formatted table it first has to be loaded with data.  Examples below will show how to load the table:

	First, initialize the table array

	$rowdata = @()

	Then Load the array.  If you are using column headers then load those into the column headers array, otherwise the first line of the table goes into the column headers array
	and the second and subsequent lines go into the $rowdata table as shown below:

	$columnHeaders = @('Display Name',$htmlsb,'Status',$htmlsb,'Startup Type',$htmlsb)

	The first column is the actual name to display, the second are the attributes of the column i.e. color anded with bold or italics.  For the anding, parens are required or it will
	not format correctly.

	This is following by adding rowdata as shown below.  As more columns are added the columns will auto adjust to fit the size of the page.

	$rowdata = @()
	$columnHeaders = @("User Name",$htmlsb,$UserName,$htmlwhite)
	$rowdata += @(,('Save as PDF',$htmlsb,$PDF.ToString(),$htmlwhite))
	$rowdata += @(,('Save as TEXT',$htmlsb,$TEXT.ToString(),$htmlwhite))
	$rowdata += @(,('Save as WORD',$htmlsb,$MSWORD.ToString(),$htmlwhite))
	$rowdata += @(,('Save as HTML',$htmlsb,$HTML.ToString(),$htmlwhite))
	$rowdata += @(,('Add DateTime',$htmlsb,$AddDateTime.ToString(),$htmlwhite))
	$rowdata += @(,('Hardware Inventory',$htmlsb,$Hardware.ToString(),$htmlwhite))
	$rowdata += @(,('Computer Name',$htmlsb,$ComputerName,$htmlwhite))
	$rowdata += @(,('Filename1',$htmlsb,$Script:FileName1,$htmlwhite))
	$rowdata += @(,('OS Detected',$htmlsb,$Script:RunningOS,$htmlwhite))
	$rowdata += @(,('PSUICulture',$htmlsb,$PSCulture,$htmlwhite))
	$rowdata += @(,('PoSH version',$htmlsb,$Host.Version.ToString(),$htmlwhite))
	FormatHTMLTable "Example of Horizontal AutoFitContents HTML Table" -rowArray $rowdata

	The 'rowArray' paramater is mandatory to build the table, but it is not set as such in the function - if nothing is passed, the table will be empty.

	Colors and Bold/Italics Flags are shown below:

		htmlbold       
		htmlitalics    
		htmlred        
		htmlcyan        
		htmlblue       
		htmldarkblue   
		htmllightblue   
		htmlpurple      
		htmlyellow      
		htmllime       
		htmlmagenta     
		htmlwhite       
		htmlsilver      
		htmlgray       
		htmlolive       
		htmlorange      
		htmlmaroon      
		htmlgreen       
		htmlblack     
#>

Function FormatHTMLTable
{
	Param
	(
		[String]   $tableheader = '',
		[String]   $tablewidth  = 'auto',
		[String]   $fontName    = 'Calibri',
		[Int]      $fontSize    = 2,
		[Switch]   $noBorder    = $false,
		[Int]      $noHeadCols  = 1,
		[Object[]] $rowArray    = $null,
		[Object[]] $fixedWidth  = $null,
		[Object[]] $columnArray = $null
	)

	## FIXME - the help text for this Function is wacky wrong - MBS
	## FIXME - Use StringBuilder - MBS - this only builds the table header - benefit relatively small
<#
	If( $SuperVerbose )
	{
		wv "FormatHTMLTable: fontname '$fontname', size $fontSize, tableheader '$tableheader'"
		wv "FormatHTMLTable: noborder $noborder, noheadcols $noheadcols"
		If( $rowarray -and $rowarray.count -gt 0 )
		{
			wv "FormatHTMLTable: rowarray has $( $rowarray.count ) elements"
		}
		Else
		{
			wv "FormatHTMLTable: rowarray is empty"
		}
		If( $columnarray -and $columnarray.count -gt 0 )
		{
			wv "FormatHTMLTable: columnarray has $( $columnarray.count ) elements"
		}
		Else
		{
			wv "FormatHTMLTable: columnarray is empty"
		}
		If( $fixedwidth -and $fixedwidth.count -gt 0 )
		{
			wv "FormatHTMLTable: fixedwidth has $( $fixedwidth.count ) elements"
		}
		Else
		{
			wv "FormatHTMLTable: fixedwidth is empty"
		}
	}
#>

	$HTMLBody = ''
	If( $tableheader.Length -gt 0 )
	{
		$HTMLBody += "<b><font face='" + $fontname + "' size='" + ($fontsize + 1) + "'>" + $tableheader + "</font></b>" + $crlf
	}

	$fwSize = If( $null -eq $fixedWidth ) { 0 } else { $fixedWidth.Count }

	If( $null -eq $columnArray -or $columnArray.Length -eq 0)
	{
		$NumCols = $noHeadCols + 1
	}  # means we have no column headers, just a table
	Else
	{
		$NumCols = $columnArray.Length
	}  # need to add one for the color attrib

	If( $null -eq $rowArray )
	{
		$NumRows = 1
	}
	Else
	{
		$NumRows = $rowArray.length + 1
	}

	If( $noBorder )
	{
		$HTMLBody += "<table border='0' width='" + $tablewidth + "'>"
	}
	Else
	{
		$HTMLBody += "<table border='1' width='" + $tablewidth + "'>"
	}
	$HTMLBody += $crlf

	If( $columnArray -and $columnArray.Length -gt 0 )
	{
		$HTMLBody += '<tr>' + $crlf

		For( $columnIndex = 0; $columnIndex -lt $NumCols; $columnindex += 2 )
		{
			#V3.00
			$val = $columnArray[ $columnIndex + 1 ]
			$tmp = $global:htmlColor[ $val -band 0xffffc ]
			[Bool] $bold = $val -band $htmlBold
			[Bool] $ital = $val -band $htmlitalics

			If( $fwSize -eq 0 )
			{
				$HTMLBody += "<td style=""background-color:$($tmp)""><font face='$($fontName)' size='$($fontSize)'>"
			}
			Else
			{
				$HTMLBody += "<td style=""width:$($fixedWidth[$columnIndex / 2]); background-color:$($tmp)""><font face='$($fontName)' size='$($fontSize)'>"
			}

			If( $bold ) { $HTMLBody += '<b>' }
			If( $ital ) { $HTMLBody += '<i>' }

			$array = $columnArray[ $columnIndex ]
			If( $array )
			{
				If( $array -eq ' ' -or $array.Length -eq 0 )
				{
					$HTMLBody += '&nbsp;&nbsp;&nbsp;'
				}
				Else
				{
					For( $i = 0; $i -lt $array.Length; $i += 2 )
					{
						If( $array[ $i ] -eq ' ' )
						{
							$HTMLBody += '&nbsp;'
						}
						Else
						{
							Break
						}
					}
					$HTMLBody += $array
				}
			}
			Else
			{
				$HTMLBody += '&nbsp;&nbsp;&nbsp;'
			}
			
			If( $bold ) { $HTMLBody += '</b>' }
			If( $ital ) { $HTMLBody += '</i>' }

			$HTMLBody += '</font></td>'
			$HTMLBody += $crlf
		}

		$HTMLBody += '</tr>' + $crlf
	}

	#V3.00
	Out-File -FilePath $Script:HTMLFileName -Append -InputObject $HTMLBody 4>$Null 
	$HTMLBody = ''

	##$rowindex = 2
	If( $rowArray )
	{
<#
		If( $ExtraSpecialVerbose )
		{
			wv "***** FormatHTMLTable: rowarray length $( $rowArray.Length )"
			For( $ii = 0; $ii -lt $rowArray.Length; $ii++ )
			{
				$row = $rowArray[ $ii ]
				wv "***** FormatHTMLTable: index $ii, type $( $row.GetType().FullName ), length $( $row.Length )"
				For( $yyy = 0; $yyy -lt $row.Length; $yyy++ )
				{
					wv "***** FormatHTMLTable: index $ii, yyy = $yyy, val = '$( $row[ $yyy ] )'"
				}
				wv "***** done"
			}
			wv "***** FormatHTMLTable: rowCount $NumRows"
		}
#>

		AddHTMLTable -fontName $fontName -fontSize $fontSize `
			-colCount $numCols -rowCount $NumRows `
			-rowInfo $rowArray -fixedInfo $fixedWidth
		##$rowArray = @()
		$rowArray = $null
		$HTMLBody = '</table>'
	}
	Else
	{
		$HTMLBody += '</table>'
	}

	Out-File -FilePath $Script:HTMLFileName -Append -InputObject $HTMLBody 4>$Null 
}
#endregion

#region other HTML functions
<#
#***********************************************************************************************************
# CheckHTMLColor - Called from AddHTMLTable WriteHTMLLine and FormatHTMLTable
#***********************************************************************************************************
Function CheckHTMLColor
{
	Param($hash)

	#V2.23 -- this is really slow. several ways to fixit. so fixit. MBS
	#V2.23 - obsolete. replaced by using $global:htmlColor lookup table
	If($hash -band $htmlwhite)
	{
		Return $htmlwhitemask
	}
	If($hash -band $htmlred)
	{
		Return $htmlredmask
	}
	If($hash -band $htmlcyan)
	{
		Return $htmlcyanmask
	}
	If($hash -band $htmlblue)
	{
		Return $htmlbluemask
	}
	If($hash -band $htmldarkblue)
	{
		Return $htmldarkbluemask
	}
	If($hash -band $htmllightblue)
	{
		Return $htmllightbluemask
	}
	If($hash -band $htmlpurple)
	{
		Return $htmlpurplemask
	}
	If($hash -band $htmlyellow)
	{
		Return $htmlyellowmask
	}
	If($hash -band $htmllime)
	{
		Return $htmllimemask
	}
	If($hash -band $htmlmagenta)
	{
		Return $htmlmagentamask
	}
	If($hash -band $htmlsilver)
	{
		Return $htmlsilvermask
	}
	If($hash -band $htmlgray)
	{
		Return $htmlgraymask
	}
	If($hash -band $htmlblack)
	{
		Return $htmlblackmask
	}
	If($hash -band $htmlorange)
	{
		Return $htmlorangemask
	}
	If($hash -band $htmlmaroon)
	{
		Return $htmlmaroonmask
	}
	If($hash -band $htmlgreen)
	{
		Return $htmlgreenmask
	}
	If($hash -band $htmlolive)
	{
		Return $htmlolivemask
	}
}
#>

Function SetupHTML
{
	Write-Verbose "$(Get-Date -Format G): Setting up HTML"
	If(!$AddDateTime)
	{
		[string]$Script:HTMLFileName = "$($Script:pwdpath)\$($OutputFileName).html"
	}
	ElseIf($AddDateTime)
	{
		[string]$Script:HTMLFileName = "$($Script:pwdpath)\$($OutputFileName)_$(Get-Date -f yyyy-MM-dd_HHmm).html"
	}

	$htmlhead = "<html><head><meta http-equiv='Content-Language' content='da'><title>" + $Script:Title + "</title></head><body>"
	Out-File -FilePath $Script:HTMLFileName -Force -InputObject $HTMLHead 4>$Null
}
#endregion

#region Iain's Word table functions

<#
.Synopsis
	Add a table to a Microsoft Word document
.DESCRIPTION
	This function adds a table to a Microsoft Word document from either an array of
	Hashtables or an array of PSCustomObjects.

	Using this function is quicker than setting each table cell individually but can
	only utilise the built-in MS Word table autoformats. Individual tables cells can
	be altered after the table has been appended to the document (a table reference
	is returned).
.EXAMPLE
	AddWordTable -Hashtable $HashtableArray

	This example adds table to the MS Word document, utilising all key/value pairs in
	the array of hashtables. Column headers will display the key names as defined.
	Note: the columns might not be displayed in the order that they were defined. To
	ensure columns are displayed in the required order utilise the -Columns parameter.
.EXAMPLE
	AddWordTable -Hashtable $HashtableArray -List

	This example adds table to the MS Word document, utilising all key/value pairs in
	the array of hashtables. No column headers will be added, in a ListView format.
	Note: the columns might not be displayed in the order that they were defined. To
	ensure columns are displayed in the required order utilise the -Columns parameter.
.EXAMPLE
	AddWordTable -CustomObject $PSCustomObjectArray

	This example adds table to the MS Word document, utilising all note property names
	the array of PSCustomObjects. Column headers will display the note property names.
	Note: the columns might not be displayed in the order that they were defined. To
	ensure columns are displayed in the required order utilise the -Columns parameter.
.EXAMPLE
	AddWordTable -Hashtable $HashtableArray -Columns FirstName,LastName,EmailAddress

	This example adds a table to the MS Word document, but only using the specified
	key names: FirstName, LastName and EmailAddress. If other keys are present in the
	array of Hashtables they will be ignored.
.EXAMPLE
	AddWordTable -CustomObject $PSCustomObjectArray -Columns FirstName,LastName,EmailAddress -Headers "First Name","Last Name","Email Address"

	This example adds a table to the MS Word document, but only using the specified
	PSCustomObject note properties: FirstName, LastName and EmailAddress. If other note
	properties are present in the array of PSCustomObjects they will be ignored. The
	display names for each specified column header has been overridden to display a
	custom header. Note: the order of the header names must match the specified columns.
#>

Function AddWordTable
{
	[CmdletBinding()]
	Param
	(
		# Array of Hashtable (including table headers)
		[Parameter(Mandatory=$True, ValueFromPipelineByPropertyName=$True, ParameterSetName='Hashtable', Position=0)]
		[ValidateNotNullOrEmpty()] [System.Collections.Hashtable[]] $Hashtable,
		# Array of PSCustomObjects
		[Parameter(Mandatory=$True, ValueFromPipelineByPropertyName=$True, ParameterSetName='CustomObject', Position=0)]
		[ValidateNotNullOrEmpty()] [PSCustomObject[]] $CustomObject,
		# Array of Hashtable key names or PSCustomObject property names to include, in display order.
		# If not supplied then all Hashtable keys or all PSCustomObject properties will be displayed.
		[Parameter(ValueFromPipelineByPropertyName=$True)] [AllowNull()] [string[]] $Columns = $Null,
		# Array of custom table header strings in display order.
		[Parameter(ValueFromPipelineByPropertyName=$True)] [AllowNull()] [string[]] $Headers = $Null,
		# AutoFit table behavior.
		[Parameter(ValueFromPipelineByPropertyName=$True)] [AllowNull()] [int] $AutoFit = -1,
		# List view (no headers)
		[Switch] $List,
		# Grid lines
		[Switch] $NoGridLines,
		[Switch] $NoInternalGridLines,
		# Built-in Word table formatting style constant
		# Would recommend only $wdTableFormatContempory for normal usage (possibly $wdTableFormatList5 for List view)
		[Parameter(ValueFromPipelineByPropertyName=$True)] [int] $Format = 0
	)

	Begin 
	{
		Write-Debug ("Using parameter set '{0}'" -f $PSCmdlet.ParameterSetName);
		## Check if -Columns wasn't specified but -Headers were (saves some additional parameter sets!)
		If(($Null -eq $Columns) -and ($Null -ne $Headers)) 
		{
			Write-Warning "No columns specified and therefore, specified headers will be ignored.";
			$Columns = $Null;
		}
		ElseIf(($Null -ne $Columns) -and ($Null -ne $Headers)) 
		{
			## Check if number of specified -Columns matches number of specified -Headers
			If($Columns.Length -ne $Headers.Length) 
			{
				Write-Error "The specified number of columns does not match the specified number of headers.";
			}
		} ## end Elseif
	} ## end Begin

	Process
	{
		## Build the Word table data string to be converted to a range and then a table later.
		[System.Text.StringBuilder] $WordRangeString = New-Object System.Text.StringBuilder;

		Switch ($PSCmdlet.ParameterSetName) 
		{
			'CustomObject' 
			{
				If($Null -eq $Columns) 
				{
					## Build the available columns from all availble PSCustomObject note properties
					[string[]] $Columns = @();
					## Add each NoteProperty name to the array
					ForEach($Property in ($CustomObject | Get-Member -MemberType NoteProperty)) 
					{ 
						$Columns += $Property.Name; 
					}
				}

				## Add the table headers from -Headers or -Columns (except when in -List(view)
				If(-not $List) 
				{
					Write-Debug ("$(Get-Date -Format G): `t`tBuilding table headers");
					If($Null -ne $Headers) 
					{
                        [ref] $Null = $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $Headers));
					}
					Else 
					{ 
                        [ref] $Null = $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $Columns));
					}
				}

				## Iterate through each PSCustomObject
				Write-Debug ("$(Get-Date -Format G): `t`tBuilding table rows");
				ForEach($Object in $CustomObject) 
				{
					$OrderedValues = @();
					## Add each row item in the specified order
					ForEach($Column in $Columns) 
					{ 
						$OrderedValues += $Object.$Column; 
					}
					## Use the ordered list to add each column in specified order
					[ref] $Null = $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $OrderedValues));
				} ## end foreach
				Write-Debug ("$(Get-Date -Format G): `t`t`tAdded '{0}' table rows" -f ($CustomObject.Count));
			} ## end CustomObject

			Default 
			{   ## Hashtable
				If($Null -eq $Columns) 
				{
					## Build the available columns from all available hashtable keys. Hopefully
					## all Hashtables have the same keys (they should for a table).
					$Columns = $Hashtable[0].Keys;
				}

				## Add the table headers from -Headers or -Columns (except when in -List(view)
				If(-not $List) 
				{
					Write-Debug ("$(Get-Date -Format G): `t`tBuilding table headers");
					If($Null -ne $Headers) 
					{ 
						[ref] $Null = $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $Headers));
					}
					Else 
					{
						[ref] $Null = $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $Columns));
					}
				}
                
				## Iterate through each Hashtable
				Write-Debug ("$(Get-Date -Format G): `t`tBuilding table rows");
				ForEach($Hash in $Hashtable) 
				{
					$OrderedValues = @();
					## Add each row item in the specified order
					ForEach($Column in $Columns) 
					{ 
						$OrderedValues += $Hash.$Column; 
					}
					## Use the ordered list to add each column in specified order
					[ref] $Null = $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $OrderedValues));
				} ## end foreach

				Write-Debug ("$(Get-Date -Format G): `t`t`tAdded '{0}' table rows" -f $Hashtable.Count);
			} ## end default
		} ## end switch

		## Create a MS Word range and set its text to our tab-delimited, concatenated string
		Write-Debug ("$(Get-Date -Format G): `t`tBuilding table range");
		$WordRange = $Script:Doc.Application.Selection.Range;
		$WordRange.Text = $WordRangeString.ToString();

		## Create hash table of named arguments to pass to the ConvertToTable method
		$ConvertToTableArguments = @{ Separator = [Microsoft.Office.Interop.Word.WdTableFieldSeparator]::wdSeparateByTabs; }

		## Negative built-in styles are not supported by the ConvertToTable method
		If($Format -ge 0) 
		{
			$ConvertToTableArguments.Add("Format", $Format);
			$ConvertToTableArguments.Add("ApplyBorders", $True);
			$ConvertToTableArguments.Add("ApplyShading", $True);
			$ConvertToTableArguments.Add("ApplyFont", $True);
			$ConvertToTableArguments.Add("ApplyColor", $True);
			If(!$List) 
			{ 
				$ConvertToTableArguments.Add("ApplyHeadingRows", $True); 
			}
			$ConvertToTableArguments.Add("ApplyLastRow", $True);
			$ConvertToTableArguments.Add("ApplyFirstColumn", $True);
			$ConvertToTableArguments.Add("ApplyLastColumn", $True);
		}

		## Invoke ConvertToTable method - with named arguments - to convert Word range to a table
		## See http://msdn.microsoft.com/en-us/library/office/aa171893(v=office.11).aspx
		Write-Debug ("$(Get-Date -Format G): `t`tConverting range to table");
		## Store the table reference just in case we need to set alternate row coloring
		$WordTable = $WordRange.GetType().InvokeMember(
			"ConvertToTable",                               # Method name
			[System.Reflection.BindingFlags]::InvokeMethod, # Flags
			$Null,                                          # Binder
			$WordRange,                                     # Target (self!)
			([Object[]]($ConvertToTableArguments.Values)),  ## Named argument values
			$Null,                                          # Modifiers
			$Null,                                          # Culture
			([String[]]($ConvertToTableArguments.Keys))     ## Named argument names
		);

		## Implement grid lines (will wipe out any existing formatting
		If($Format -lt 0) 
		{
			Write-Debug ("$(Get-Date -Format G): `t`tSetting table format");
			$WordTable.Style = $Format;
		}

		## Set the table autofit behavior
		If($AutoFit -ne -1) 
		{ 
			$WordTable.AutoFitBehavior($AutoFit); 
		}

		If(!$List)
		{
			#the next line causes the heading row to flow across page Breaks
			$WordTable.Rows.First.Headingformat = $wdHeadingFormatTrue;
		}

		If(!$NoGridLines) 
		{
			$WordTable.Borders.InsideLineStyle = $wdLineStyleSingle;
			$WordTable.Borders.OutsideLineStyle = $wdLineStyleSingle;
		}
		If($NoGridLines) 
		{
			$WordTable.Borders.InsideLineStyle = $wdLineStyleNone;
			$WordTable.Borders.OutsideLineStyle = $wdLineStyleNone;
		}
		If($NoInternalGridLines) 
		{
			$WordTable.Borders.InsideLineStyle = $wdLineStyleNone;
			$WordTable.Borders.OutsideLineStyle = $wdLineStyleSingle;
		}

		Return $WordTable;

	} ## end Process
}

<#
.Synopsis
	Sets the format of one or more Word table cells
.DESCRIPTION
	This function sets the format of one or more table cells, either from a collection
	of Word COM object cell references, an individual Word COM object cell reference or
	a hashtable containing Row and Column information.

	The font name, font size, bold, italic , underline and shading values can be used.
.EXAMPLE
	SetWordCellFormat -Hashtable $Coordinates -Table $TableReference -Bold

	This example sets all text to bold that is contained within the $TableReference
	Word table, using an array of hashtables. Each hashtable contain a pair of co-
	ordinates that is used to select the required cells. Note: the hashtable must
	contain the .Row and .Column key names. For example:
	@ { Row = 7; Column = 3 } to set the cell at row 7 and column 3 to bold.
.EXAMPLE
	$RowCollection = $Table.Rows.First.Cells
	SetWordCellFormat -Collection $RowCollection -Bold -Size 10

	This example sets all text to size 8 and bold for all cells that are contained
	within the first row of the table.
	Note: the $Table.Rows.First.Cells returns a collection of Word COM cells objects
	that are in the first table row.
.EXAMPLE
	$ColumnCollection = $Table.Columns.Item(2).Cells
	SetWordCellFormat -Collection $ColumnCollection -BackgroundColor 255

	This example sets the background (shading) of all cells in the table's second
	column to red.
	Note: the $Table.Columns.Item(2).Cells returns a collection of Word COM cells objects
	that are in the table's second column.
.EXAMPLE
	SetWordCellFormat -Cell $Table.Cell(17,3) -Font "Tahoma" -Color 16711680

	This example sets the font to Tahoma and the text color to blue for the cell located
	in the table's 17th row and 3rd column.
	Note: the $Table.Cell(17,3) returns a single Word COM cells object.
#>

Function SetWordCellFormat 
{
	[CmdletBinding(DefaultParameterSetName='Collection')]
	Param (
		# Word COM object cell collection reference
		[Parameter(Mandatory=$true, ValueFromPipeline=$true, ParameterSetName='Collection', Position=0)] [ValidateNotNullOrEmpty()] $Collection,
		# Word COM object individual cell reference
		[Parameter(Mandatory=$true, ParameterSetName='Cell', Position=0)] [ValidateNotNullOrEmpty()] $Cell,
		# Hashtable of cell co-ordinates
		[Parameter(Mandatory=$true, ParameterSetName='Hashtable', Position=0)] [ValidateNotNullOrEmpty()] [System.Collections.Hashtable[]] $Coordinates,
		# Word COM object table reference
		[Parameter(Mandatory=$true, ParameterSetName='Hashtable', Position=1)] [ValidateNotNullOrEmpty()] $Table,
		# Font name
		[Parameter()] [AllowNull()] [string] $Font = $Null,
		# Font color
		[Parameter()] [AllowNull()] $Color = $Null,
		# Font size
		[Parameter()] [ValidateNotNullOrEmpty()] [int] $Size = 0,
		# Cell background color
		[Parameter()] [AllowNull()] [int]$BackgroundColor = $Null,
		# Force solid background color
		[Switch] $Solid,
		[Switch] $Bold,
		[Switch] $Italic,
		[Switch] $Underline
	)

	Begin 
	{
		Write-Debug ("Using parameter set '{0}'." -f $PSCmdlet.ParameterSetName);
	}

	Process 
	{
		Switch ($PSCmdlet.ParameterSetName) 
		{
			'Collection' {
				ForEach($Cell in $Collection) 
				{
					If($Null -ne $BackgroundColor) { $Cell.Shading.BackgroundPatternColor = $BackgroundColor; }
					If($Bold) { $Cell.Range.Font.Bold = $true; }
					If($Italic) { $Cell.Range.Font.Italic = $true; }
					If($Underline) { $Cell.Range.Font.Underline = 1; }
					If($Null -ne $Font) { $Cell.Range.Font.Name = $Font; }
					If($Null -ne $Color) { $Cell.Range.Font.Color = $Color; }
					If($Size -ne 0) { $Cell.Range.Font.Size = $Size; }
					If($Solid) { $Cell.Shading.Texture = 0; } ## wdTextureNone
				} # end ForEach
			} # end Collection
			'Cell' 
			{
				If($Bold) { $Cell.Range.Font.Bold = $true; }
				If($Italic) { $Cell.Range.Font.Italic = $true; }
				If($Underline) { $Cell.Range.Font.Underline = 1; }
				If($Null -ne $Font) { $Cell.Range.Font.Name = $Font; }
				If($Null -ne $Color) { $Cell.Range.Font.Color = $Color; }
				If($Size -ne 0) { $Cell.Range.Font.Size = $Size; }
				If($Null -ne $BackgroundColor) { $Cell.Shading.BackgroundPatternColor = $BackgroundColor; }
				If($Solid) { $Cell.Shading.Texture = 0; } ## wdTextureNone
			} # end Cell
			'Hashtable' 
			{
				ForEach($Coordinate in $Coordinates) 
				{
					$Cell = $Table.Cell($Coordinate.Row, $Coordinate.Column);
					If($Bold) { $Cell.Range.Font.Bold = $true; }
					If($Italic) { $Cell.Range.Font.Italic = $true; }
					If($Underline) { $Cell.Range.Font.Underline = 1; }
					If($Null -ne $Font) { $Cell.Range.Font.Name = $Font; }
					If($Null -ne $Color) { $Cell.Range.Font.Color = $Color; }
					If($Size -ne 0) { $Cell.Range.Font.Size = $Size; }
					If($Null -ne $BackgroundColor) { $Cell.Shading.BackgroundPatternColor = $BackgroundColor; }
					If($Solid) { $Cell.Shading.Texture = 0; } ## wdTextureNone
				}
			} # end Hashtable
		} # end Switch
	} # end process
}

<#
.Synopsis
	Sets alternate row colors in a Word table
.DESCRIPTION
	This function sets the format of alternate rows within a Word table using the
	specified $BackgroundColor. This function is expensive (in performance terms) as
	it recursively sets the format on alternate rows. It would be better to pick one
	of the predefined table formats (if one exists)? Obviously the more rows, the
	longer it takes :'(

	Note: this function is called by the AddWordTable function if an alternate row
	format is specified.
.EXAMPLE
	SetWordTableAlternateRowColor -Table $TableReference -BackgroundColor 255

	This example sets every-other table (starting with the first) row and sets the
	background color to red (wdColorRed).
.EXAMPLE
	SetWordTableAlternateRowColor -Table $TableReference -BackgroundColor 39423 -Seed Second

	This example sets every other table (starting with the second) row and sets the
	background color to light orange (weColorLightOrange).
#>

Function SetWordTableAlternateRowColor 
{
	[CmdletBinding()]
	Param (
		# Word COM object table reference
		[Parameter(Mandatory=$true, ValueFromPipeline=$true, Position=0)] [ValidateNotNullOrEmpty()] $Table,
		# Alternate row background color
		[Parameter(Mandatory=$true, Position=1)] [ValidateNotNull()] [int] $BackgroundColor,
		# Alternate row starting seed
		[Parameter(ValueFromPipelineByPropertyName=$true, Position=2)] [ValidateSet('First','Second')] [string] $Seed = 'First'
	)

	Process 
	{
		$StartDateTime = Get-Date;
		Write-Debug ("{0}: `t`tSetting alternate table row colors.." -f $StartDateTime);

		## Determine the row seed (only really need to check for 'Second' and default to 'First' otherwise
		If($Seed.ToLower() -eq 'second') 
		{ 
			$StartRowIndex = 2; 
		}
		Else 
		{ 
			$StartRowIndex = 1; 
		}

		For($AlternateRowIndex = $StartRowIndex; $AlternateRowIndex -lt $Table.Rows.Count; $AlternateRowIndex += 2) 
		{ 
			$Table.Rows.Item($AlternateRowIndex).Shading.BackgroundPatternColor = $BackgroundColor;
		}

		## I've put verbose calls in here we can see how expensive this functionality actually is.
		$EndDateTime = Get-Date;
		$ExecutionTime = New-TimeSpan -Start $StartDateTime -End $EndDateTime;
		Write-Debug ("{0}: `t`tDone setting alternate row style color in '{1}' seconds" -f $EndDateTime, $ExecutionTime.TotalSeconds);
	}
}
#endregion

#region general script functions
Function validStateProp( [object] $object, [string] $topLevel, [string] $secondLevel )
{
	#function created 8-jan-2014 by Michael B. Smith
	If( $object )
	{
		If((Get-Member -Name $topLevel -InputObject $object))
		{
			If((Get-Member -Name $secondLevel -InputObject $object.$topLevel))
			{
				Return $True
			}
		}
	}
	Return $False
}

Function validObject( [object] $object, [string] $topLevel )
{
	#function created 8-jan-2014 by Michael B. Smith
	If( $object )
	{
		If((Get-Member -Name $topLevel -InputObject $object))
		{
			Return $True
		}
	}
	Return $False
}

Function ShowScriptOptions
{
	Write-Verbose "$(Get-Date -Format G): "
	Write-Verbose "$(Get-Date -Format G): "
	Write-Verbose "$(Get-Date -Format G): Add DateTime    : $AddDateTime"
	Write-Verbose "$(Get-Date -Format G): All DNS Servers : $AllDNSServers"
	If($MSWORD -or $PDF)
	{
		Write-Verbose "$(Get-Date -Format G): Company Name    : $Script:CoName"
	}
	Write-Verbose "$(Get-Date -Format G): Computer Name   : $ComputerName"
	If($MSWORD -or $PDF)
	{
		Write-Verbose "$(Get-Date -Format G): Company Address : $($CompanyAddress)"
		Write-Verbose "$(Get-Date -Format G): Company Email   : $($CompanyEmail)"
		Write-Verbose "$(Get-Date -Format G): Company Fax     : $($CompanyFax)"
		Write-Verbose "$(Get-Date -Format G): Company Phone   : $($CompanyPhone)"
		Write-Verbose "$(Get-Date -Format G): Cover Page      : $CoverPage"
	}
	Write-Verbose "$(Get-Date -Format G): Details         : $Details"
	Write-Verbose "$(Get-Date -Format G): Dev             : $Dev"
	If($Dev)
	{
		Write-Verbose "$(Get-Date -Format G): DevErrorFile    : $Script:DevErrorFile"
	}
	If($MSWord)
	{
		Write-Verbose "$(Get-Date -Format G): Word FileName   : $($Script:WordFileName)"
	}
	If($HTML)
	{
		Write-Verbose "$(Get-Date -Format G): HTML FileName   : $($Script:HTMLFileName)"
	} 
	If($PDF)
	{
		Write-Verbose "$(Get-Date -Format G): PDF FileName    : $($Script:PDFFileName)"
	}
	If($Text)
	{
		Write-Verbose "$(Get-Date -Format G): Text FileName   : $($Script:TextFileName)"
	}
	Write-Verbose "$(Get-Date -Format G): Folder          : $Folder"
	Write-Verbose "$(Get-Date -Format G): From            : $From"
	Write-Verbose "$(Get-Date -Format G): Log             : $($Log)"
	Write-Verbose "$(Get-Date -Format G): Save As HTML    : $HTML"
	Write-Verbose "$(Get-Date -Format G): Save As PDF     : $PDF"
	Write-Verbose "$(Get-Date -Format G): Save As Text    : $Text"
	Write-Verbose "$(Get-Date -Format G): Save As Word    : $MSWord"
	Write-Verbose "$(Get-Date -Format G): Script Info     : $ScriptInfo"
	Write-Verbose "$(Get-Date -Format G): Smtp Port       : $SmtpPort"
	Write-Verbose "$(Get-Date -Format G): Smtp Server     : $SmtpServer"
	Write-Verbose "$(Get-Date -Format G): Title           : $Script:Title"
	Write-Verbose "$(Get-Date -Format G): To              : $To"
	Write-Verbose "$(Get-Date -Format G): Use SSL         : $UseSSL"
	If($MSWORD -or $PDF)
	{
		Write-Verbose "$(Get-Date -Format G): Username        : $UserName"
	}
	Write-Verbose "$(Get-Date -Format G): "
	Write-Verbose "$(Get-Date -Format G): OS Detected     : $Script:RunningOS"
	Write-Verbose "$(Get-Date -Format G): PSUICulture     : $PSUICulture"
	Write-Verbose "$(Get-Date -Format G): PSCulture       : $PSCulture"
	If($MSWORD -or $PDF)
	{
		Write-Verbose "$(Get-Date -Format G): Word version    : $WordProduct"
		Write-Verbose "$(Get-Date -Format G): Word language   : $Script:WordLanguageValue"
	}
	Write-Verbose "$(Get-Date -Format G): PoSH version    : $($Host.Version)"
	Write-Verbose "$(Get-Date -Format G): "
	Write-Verbose "$(Get-Date -Format G): Script start  : $Script:StartTime"
	Write-Verbose "$(Get-Date -Format G): "
	Write-Verbose "$(Get-Date -Format G): "
}

Function SaveandCloseDocumentandShutdownWord
{
	#bug fix 1-Apr-2014
	#reset Grammar and Spelling options back to their original settings
	$Script:Word.Options.CheckGrammarAsYouType = $Script:CurrentGrammarOption
	$Script:Word.Options.CheckSpellingAsYouType = $Script:CurrentSpellingOption

	Write-Verbose "$(Get-Date -Format G): Save and Close document and Shutdown Word"
	If($Script:WordVersion -eq $wdWord2010)
	{
		#the $saveFormat below passes StrictMode 2
		#I found this at the following link
		#http://msdn.microsoft.com/en-us/library/microsoft.office.interop.word.wdsaveformat(v=office.14).aspx
		If($PDF)
		{
			Write-Verbose "$(Get-Date -Format G): Saving as DOCX file first before saving to PDF"
		}
		Else
		{
			Write-Verbose "$(Get-Date -Format G): Saving DOCX file"
		}
		Write-Verbose "$(Get-Date -Format G): Running $($Script:WordProduct) and detected operating system $($Script:RunningOS)"
		$saveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], "wdFormatDocumentDefault")
		$Script:Doc.SaveAs([REF]$Script:WordFileName, [ref]$SaveFormat)
		If($PDF)
		{
			Write-Verbose "$(Get-Date -Format G): Now saving as PDF"
			$saveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], "wdFormatPDF")
			$Script:Doc.SaveAs([REF]$Script:PDFFileName, [ref]$saveFormat)
		}
	}
	ElseIf($Script:WordVersion -eq $wdWord2013 -or $Script:WordVersion -eq $wdWord2016)
	{
		If($PDF)
		{
			Write-Verbose "$(Get-Date -Format G): Saving as DOCX file first before saving to PDF"
		}
		Else
		{
			Write-Verbose "$(Get-Date -Format G): Saving DOCX file"
		}
		Write-Verbose "$(Get-Date -Format G): Running $($Script:WordProduct) and detected operating system $($Script:RunningOS)"
		$Script:Doc.SaveAs2([REF]$Script:WordFileName, [ref]$wdFormatDocumentDefault)
		If($PDF)
		{
			Write-Verbose "$(Get-Date -Format G): Now saving as PDF"
			$Script:Doc.SaveAs([REF]$Script:PDFFileName, [ref]$wdFormatPDF)
		}
	}

	Write-Verbose "$(Get-Date -Format G): Closing Word"
	$Script:Doc.Close()
	$Script:Word.Quit()
	Write-Verbose "$(Get-Date -Format G): System Cleanup"
	[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Script:Word) | Out-Null
	If(Test-Path variable:global:word)
	{
		Remove-Variable -Name word -Scope Global 4>$Null
	}
	$SaveFormat = $Null
	[gc]::collect() 
	[gc]::WaitForPendingFinalizers()
	
	#is the winword process still running? kill it

	#find out our session (usually "1" except on TS/RDC or Citrix)
	$SessionID = (Get-Process -PID $PID).SessionId

	#Find out if winword is running in our session
	$wordprocess = $Null
	$wordprocess = ((Get-Process 'WinWord' -ea 0) | Where-Object {$_.SessionId -eq $SessionID}).Id
	If($null -ne $wordprocess -and $wordprocess -gt 0)
	{
		Write-Verbose "$(Get-Date -Format G): WinWord process is still running. Attempting to stop WinWord process # $($wordprocess)"
		Stop-Process $wordprocess -EA 0
	}
}

Function SetupText
{
	Write-Verbose "$(Get-Date -Format G): Setting up Text"

	[System.Text.StringBuilder] $global:Output = New-Object System.Text.StringBuilder( 16384 )

	If(!$AddDateTime)
	{
		[string]$Script:TextFileName = "$($Script:pwdpath)\$($OutputFileName).txt"
	}
	ElseIf($AddDateTime)
	{
		[string]$Script:TextFileName = "$($Script:pwdpath)\$($OutputFileName)_$(Get-Date -f yyyy-MM-dd_HHmm).txt"
	}
}

Function SaveandCloseTextDocument
{
	Write-Verbose "$(Get-Date -Format G): Saving Text file"
	Write-Output $global:Output.ToString() | Out-File $Script:TextFileName 4>$Null
}

Function SaveandCloseHTMLDocument
{
	Write-Verbose "$(Get-Date -Format G): Saving HTML file"
	Out-File -FilePath $Script:HTMLFileName -Append -InputObject "<p></p></body></html>" 4>$Null
}

Function SetFilenames
{
	Param([string]$OutputFileName)
	
	If($MSWord -or $PDF)
	{
		CheckWordPreReq
		
		SetupWord
	}
	If($Text)
	{
		SetupText
	}
	If($HTML)
	{
		SetupHTML
	}
	ShowScriptOptions
}

Function testPort
{
	Param
	(
	[String] $computer,
	[Int[]]  $ports,
	[Int]    $timeOut,
	[Bool]   $quiet = $false
	)

	If( $result = $computer -as [System.Net.IpAddress] )
	{
		## we got passed an IP address, not a DNS name. Resolve-DnsName doesn't just
		## pass it through, but instead returns a PTR record. I consider it broken,
		## but it is what it is.
		$success = testPortsOnOneIP $computer $ports $timeOut $result.AddressFamily $quiet

		Return $success
	}

	$results = Resolve-DnsName -Name $computer -Type A_AAAA -EA 0 4>$Null

	$success = $false

	ForEach( $result in $results )
	{
		$type = $result.Type.ToString()
		If( $type -ne 'A' -and $type -ne 'AAAA' )
		{
			Continue
		}

		$ip = $result.IPAddress
		If( $type -eq 'AAAA' )
		{
			If( -not ( canRoute $ip ) )
			{
				Continue
			}

			$family = [System.Net.Sockets.AddressFamily]::InterNetworkv6
		}
		Else
		{
			$family = [System.Net.Sockets.AddressFamily]::InterNetwork
		}

		$success = $success -or ( testPortsOnOneIP $ip $ports $timeOut $family $quiet )
	}

	$results = $null

	$success
}

Function testPortsOnOneIP
{
	Param
	(
		[String] $ip,
		[Int[]]  $ports,
		[Int]    $timeOut,
		[System.Net.Sockets.AddressFamily] $family,
		[Bool]   $quiet
	)

	$success = $false

	ForEach( $port in $ports )
	{
		$tcpclient = New-Object System.Net.Sockets.TcpClient( $family )

		$async = $tcpclient.BeginConnect( $ip, $port, $null, $null )
		$wait  = $async.AsyncWaitHandle.WaitOne( $timeOut, $false )
		If( !$wait )
		{
			$tcpclient.Close()
			Continue
		}
		Else
		{
			$error.Clear()
			$null = $tcpclient.EndConnect( $async )
			If( $error -and $error.Count -gt 0 )
			{
			}
			Else
			{
				$success = $true
			}
			$tcpclient.Close()
		}

		$wait      = $null
		$async     = $null
		$tcpclient = $null

		If( $success )
		{
			## break
		}
	}

	$success
}

Function TestComputerName
{
	Param([string]$Cname)

	$DNSPort = 53
	$DNSTimeout = 300	#milliseconds

	If(![String]::IsNullOrEmpty($CName)) 
	{
		#get computer name
		#first test to make sure the computer is reachable
		Write-Verbose "$(Get-Date -Format G): Testing to see if $CName is online, reachable, and a DNS server"
		If(TestPort $CName $DNSPort $DNSTimeout)
		{
			Write-Verbose "$(Get-Date -Format G): Server $CName is online and responding on port 53"
		}
		Else
		{
			Write-Output "$(Get-Date -Format G): Computer $CName is either offline or not a DNS server (port 53)" | Out-File $Script:BadDNSErrorFile -Append 4>$Null
			Return "BAD"
		}
	}

	#if computer name is an IP address, get host name from DNS
	#http://blogs.technet.com/b/gary/archive/2009/08/29/resolve-ip-addresses-to-hostname-using-powershell.aspx
	#help from Michael B. Smith
	$ip = $CName -as [System.Net.IpAddress]
	If($ip)
	{
		$Result = [System.Net.Dns]::gethostentry($ip)
		
		If($? -and $Null -ne $Result)
		{
			$CName = $Result.HostName
			Write-Verbose "$(Get-Date -Format G): Computer name has been changed from $ip to $CName"
			Write-Verbose "$(Get-Date -Format G): Testing to see if $CName is a DNS Server"
			$results = Get-DNSServer -ComputerName $CName -EA 0 2>$Null 3>$Null 4>$Null
			If($? -and $Null -ne $results)
			{
				#the computer is a dns server
				Write-Verbose "$(Get-Date -Format G): Computer $CName is a DNS Server"
				Return $CName
			}
			ElseIf(!$? -or $Null -eq $results)
			{
				#the computer is not a dns server
				Write-Verbose "$(Get-Date -Format G): Computer $CName is not a DNS Server or the Trust Points node is missing from the DNS console"
				Write-Output "$(Get-Date -Format G): Computer $CName is not a DNS Server or the Trust Points node is missing from the DNS console" | Out-File $Script:BadDNSErrorFile -Append 4>$Null
				Return "BAD"
			}
		}
		Else
		{
			Write-Verbose "$(Get-Date -Format G): Unable to resolve $CName to a hostname"
			Write-Output "$(Get-Date -Format G): Unable to resolve $CName to a hostname" | Out-File $Script:BadDNSErrorFile -Append 4>$Null
			Return "BAD"
		}
	}
	Else
	{
		Write-Verbose "$(Get-Date -Format G): Testing to see if $CName is a DNS Server"
		$results = Get-DNSServer -ComputerName $CName -EA 0 2>$Null 3>$Null 4>$Null
		If($? -and $Null -ne $results)
		{
			#the computer is a dns server
			Write-Verbose "$(Get-Date -Format G): Computer $CName is a DNS Server"
			Return $CName
		}
		ElseIf(!$? -or $Null -eq $results)
		{
			#the computer is not a dns server
			Write-Verbose "$(Get-Date -Format G): Computer $CName is not a DNS Server or the Trust Points node is missing from the DNS console"
			Write-Output "$(Get-Date -Format G): Computer $CName is not a DNS Server or the Trust Points node is missing from the DNS console" | Out-File $Script:BadDNSErrorFile -Append 4>$Null
			Return "BAD"
		}
	}

	Write-Verbose "$(Get-Date -Format G): "
	Return $CName
}

Function ProcessDocumentOutput
{
	If($MSWORD -or $PDF)
	{
		SaveandCloseDocumentandShutdownWord
	}
	If($Text)
	{
		SaveandCloseTextDocument
	}
	If($HTML)
	{
		SaveandCloseHTMLDocument
	}

	$GotFile = $False

	If($MSWord)
	{
		If(Test-Path "$($Script:WordFileName)")
		{
			Write-Verbose "$(Get-Date -Format G): $($Script:WordFileName) is ready for use"
			$GotFile = $True
		}
		Else
		{
			Write-Warning "$(Get-Date -Format G): Unable to save the output file, $($Script:WordFileName)"
			Write-Error "Unable to save the output file, $($Script:WordFileName)"
		}
	}
	If($PDF)
	{
		If(Test-Path "$($Script:PDFFileName)")
		{
			Write-Verbose "$(Get-Date -Format G): $($Script:PDFFileName) is ready for use"
			$GotFile = $True
		}
		Else
		{
			Write-Warning "$(Get-Date -Format G): Unable to save the output file, $($Script:PDFFileName)"
			Write-Error "Unable to save the output file, $($Script:PDFFileName)"
		}
	}
	If($Text)
	{
		If(Test-Path "$($Script:TextFileName)")
		{
			Write-Verbose "$(Get-Date -Format G): $($Script:TextFileName) is ready for use"
			$GotFile = $True
		}
		Else
		{
			Write-Warning "$(Get-Date -Format G): Unable to save the output file, $($Script:TextFileName)"
			Write-Error "Unable to save the output file, $($Script:TextFileName)"
		}
	}
	If($HTML)
	{
		If(Test-Path "$($Script:HTMLFileName)")
		{
			Write-Verbose "$(Get-Date -Format G): $($Script:HTMLFileName) is ready for use"
			$GotFile = $True
		}
		Else
		{
			Write-Warning "$(Get-Date -Format G): Unable to save the output file, $($Script:HTMLFileName)"
			Write-Error "Unable to save the output file, $($Script:HTMLFileName)"
		}
	}
	
	#email output file if requested
	If($GotFile -and ![System.String]::IsNullOrEmpty( $SmtpServer ))
	{
		$emailattachments = @()
		If($MSWord)
		{
			$emailAttachments += $Script:WordFileName
		}
		If($PDF)
		{
			$emailAttachments += $Script:PDFFileName
		}
		If($Text)
		{
			$emailAttachments += $Script:TextFileName
		}
		If($HTML)
		{
			$emailAttachments += $Script:HTMLFileName
		}
		SendEmail $emailAttachments
	}
}

Function AbortScript
{
	If($MSWord -or $PDF)
	{
		$Script:Word.quit()
		Write-Verbose "$(Get-Date -Format G): System Cleanup"
		[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Script:Word) | Out-Null
		If(Test-Path variable:global:word)
		{
			Remove-Variable -Name word -Scope Global
		}
	}
	Write-Verbose "$(Get-Date -Format G): Script has been aborted"
	$ErrorActionPreference = $SaveEAPreference
	Exit
}
#endregion

#region script setup function
Function ProcessScriptStart
{
	$script:startTime = Get-Date

	$Script:DNSServerNames = @()
	If($AllDNSServers -eq $False)
	{
		$CName = TestComputerName $ComputerName
		If($CName -ne "BAD")
		{
			$Script:DNSServerNames += $CName
		}
		Else
		{
			$ErrorActionPreference = $SaveEAPreference
			Write-Error "
			`n`n
			`t`t
			Computer $ComputerName is offline or is not a DNS server (port 53).
			`n`n
			`t`t
			Script cannot continue.
			`n`n
			"
			Exit
		}
	}
	Else
	{
		Write-Verbose "$(Get-Date -Format G): Retrieving all DNS servers in domain"
		$ComputerName = "All DNS Servers"
		
		$ALLServers = dsquery * forestroot -filter "(servicePrincipalName=DNS*)"
		
		If($Null -eq $AllServers)
		{
			#oops no DNS servers (which shouldn't happen in AD)
			Write-Error "
			`n`n
			`t`t
			Unable to retrieve any AD DNS servers.
			`n`n
			`t`t
			Script cannot continue.
			`n`n
			"
			Exit
		}
		Else
		{
			[int]$cnt = 0
			If($AllServers -is [array])
			{
				$cnt = $AllServers.Count
				Write-Verbose "$(Get-Date -Format G): $($cnt) DNS servers were found"
			}
			Else
			{
				$cnt = 1
				Write-Verbose "$(Get-Date -Format G): $($cnt) DNS server was found"
			}
			
			$Script:BadDNSErrorFile = "$Script:pwdPath\BadDNSServers_$(Get-Date -f yyyy-MM-dd_HHmm) for the Domain $Script:RptDomain.txt"

			ForEach($Server in $AllServers)
			{
				$TmpArray = $Server.Split("=").Split(",")
				$DNSServer = $TmpArray[1]
				
				$Result = TestComputerName $DNSServer
				
				If($Result -ne "BAD")
				{
					$Script:DNSServerNames += $Result
				}
			}
			
			$Script:DNSServerNames = $Script:DNSServerNames | Sort-Object
			
			Write-Verbose "$(Get-Date -Format G): $($Script:DNSServerNames.Count) DNS servers will be processed"
			Write-Verbose "$(Get-Date -Format G): "
		}
	}
}

Function ProcessScriptEnd
{
	Write-Verbose "$(Get-Date -Format G): Script has completed"
	Write-Verbose "$(Get-Date -Format G): "

	#http://poshtips.com/measuring-elapsed-time-in-powershell/
	Write-Verbose "$(Get-Date -Format G): Script started: $Script:StartTime"
	Write-Verbose "$(Get-Date -Format G): Script ended: $(Get-Date)"
	$runtime = $(Get-Date) - $Script:StartTime
	$Str = [string]::format("{0} days, {1} hours, {2} minutes, {3}.{4} seconds", `
		$runtime.Days, `
		$runtime.Hours, `
		$runtime.Minutes, `
		$runtime.Seconds,
		$runtime.Milliseconds)
	Write-Verbose "$(Get-Date -Format G): Elapsed time: $Str"

	If($Dev)
	{
		If($SmtpServer -eq "")
		{
			Out-File -FilePath $Script:DevErrorFile -InputObject $error 4>$Null
		}
		Else
		{
			Out-File -FilePath $Script:DevErrorFile -InputObject $error -Append 4>$Null
		}
	}

	If($ScriptInfo)
	{
		$SIFile = "$Script:pwdPath\DNSInventoryScriptInfo_$(Get-Date -f yyyy-MM-dd_HHmm) for the Domain $Script:RptDomain.txt"
		Out-File -FilePath $SIFile -InputObject "" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Add DateTime       : $AddDateTime" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "All DNS Servers    : $AllDNSServers" 4>$Null
		If($MSWORD -or $PDF)
		{
			Out-File -FilePath $SIFile -Append -InputObject "Company Name       : $Script:CoName" 4>$Null		
		}
		Out-File -FilePath $SIFile -Append -InputObject "ComputerName       : $ComputerName" 4>$Null		
		If($MSWORD -or $PDF)
		{
			Out-File -FilePath $SIFile -Append -InputObject "Company Address    : $CompanyAddress" 4>$Null		
			Out-File -FilePath $SIFile -Append -InputObject "Company Email      : $CompanyEmail" 4>$Null		
			Out-File -FilePath $SIFile -Append -InputObject "Company Fax        : $CompanyFax" 4>$Null		
			Out-File -FilePath $SIFile -Append -InputObject "Company Phone      : $CompanyPhone" 4>$Null		
			Out-File -FilePath $SIFile -Append -InputObject "Cover Page         : $CoverPage" 4>$Null
		}
		Out-File -FilePath $SIFile -Append -InputObject "Details            : $Details" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Dev                : $Dev" 4>$Null
		If($Dev)
		{
			Out-File -FilePath $SIFile -Append -InputObject "DevErrorFile       : $Script:DevErrorFile" 4>$Null
		}
		If($MSWord)
		{
			Out-File -FilePath $SIFile -Append -InputObject "Word FileName      : $($Script:WordFileName)" 4>$Null
		}
		If($HTML)
		{
			Out-File -FilePath $SIFile -Append -InputObject "HTML FileName      : $($Script:HTMLFileName)" 4>$Null
		}
		If($PDF)
		{
			Out-File -FilePath $SIFile -Append -InputObject "PDF Filename       : $($Script:PDFFileName)" 4>$Null
		}
		If($Text)
		{
			Out-File -FilePath $SIFile -Append -InputObject "Text FileName      : $($Script:TextFileName)" 4>$Null
		}
		Out-File -FilePath $SIFile -Append -InputObject "Folder             : $Folder" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "From               : $From" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Log                : $Log" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Save As HTML       : $HTML" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Save As PDF        : $PDF" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Save As TEXT       : $TEXT" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Save As WORD       : $MSWORD" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Script Info        : $ScriptInfo" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Smtp Port          : $SmtpPort" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Smtp Server        : $SmtpServer" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "To                 : $To" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Use SSL            : $UseSSL" 4>$Null
		If($MSWORD -or $PDF)
		{
			Out-File -FilePath $SIFile -Append -InputObject "Username           : $UserName" 4>$Null
		}
		Out-File -FilePath $SIFile -Append -InputObject "" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "OS Detected        : $RunningOS" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "PSUICulture        : $PSUICulture" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "PSCulture          : $PSCulture" 4>$Null
		If($MSWORD -or $PDF)
		{
			Out-File -FilePath $SIFile -Append -InputObject "Word version       : $Script:WordProduct" 4>$Null
			Out-File -FilePath $SIFile -Append -InputObject "Word language      : $Script:WordLanguageValue" 4>$Null
		}
		Out-File -FilePath $SIFile -Append -InputObject "PoSH version       : $($Host.Version)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Script start       : $Script:StartTime" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Elapsed time       : $Str" 4>$Null
	}
	
	#V1.09 added
	#stop transcript logging
	If($Log -eq $True) 
	{
		If($Script:StartLog -eq $true) 
		{
			try 
			{
				Stop-Transcript | Out-Null
				Write-Verbose "$(Get-Date -Format G): $Script:LogPath is ready for use"
			} 
			catch 
			{
				Write-Verbose "$(Get-Date -Format G): Transcript/log stop failed"
			}
		}
	}
	$runtime = $Null
	$Str = $Null
	$ErrorActionPreference = $SaveEAPreference
}
#endregion

#region email function
Function SendEmail
{
	Param([array]$Attachments)
	Write-Verbose "$(Get-Date -Format G): Prepare to email"

	$emailAttachment = $Attachments
	$emailSubject = $Script:Title
	$emailBody = @"
Hello, <br />
<br />
$Script:Title is attached.

"@ 

	If($Dev)
	{
		Out-File -FilePath $Script:DevErrorFile -InputObject $error 4>$Null
	}

	$error.Clear()
	
	If($From -Like "anonymous@*")
	{
		#https://serverfault.com/questions/543052/sending-unauthenticated-mail-through-ms-exchange-with-powershell-windows-server
		$anonUsername = "anonymous"
		$anonPassword = ConvertTo-SecureString -String "anonymous" -AsPlainText -Force
		$anonCredentials = New-Object System.Management.Automation.PSCredential($anonUsername,$anonPassword)

		If($UseSSL)
		{
			Send-MailMessage -Attachments $emailAttachment -Body $emailBody -BodyAsHtml -From $From `
			-Port $SmtpPort -SmtpServer $SmtpServer -Subject $emailSubject -To $To `
			-UseSSL -credential $anonCredentials *>$Null 
		}
		Else
		{
			Send-MailMessage -Attachments $emailAttachment -Body $emailBody -BodyAsHtml -From $From `
			-Port $SmtpPort -SmtpServer $SmtpServer -Subject $emailSubject -To $To `
			-credential $anonCredentials *>$Null 
		}
		
		If($?)
		{
			Write-Verbose "$(Get-Date -Format G): Email successfully sent using anonymous credentials"
		}
		ElseIf(!$?)
		{
			$e = $error[0]

			Write-Verbose "$(Get-Date -Format G): Email was not sent:"
			Write-Warning "$(Get-Date -Format G): Exception: $e.Exception" 
		}
	}
	Else
	{
		If($UseSSL)
		{
			Write-Verbose "$(Get-Date -Format G): Trying to send email using current user's credentials with SSL"
			Send-MailMessage -Attachments $emailAttachment -Body $emailBody -BodyAsHtml -From $From `
			-Port $SmtpPort -SmtpServer $SmtpServer -Subject $emailSubject -To $To `
			-UseSSL *>$Null
		}
		Else
		{
			Write-Verbose  "$(Get-Date -Format G): Trying to send email using current user's credentials without SSL"
			Send-MailMessage -Attachments $emailAttachment -Body $emailBody -BodyAsHtml -From $From `
			-Port $SmtpPort -SmtpServer $SmtpServer -Subject $emailSubject -To $To *>$Null
		}

		If(!$?)
		{
			$e = $error[0]
			
			#error 5.7.57 is O365 and error 5.7.0 is gmail
			If($null -ne $e.Exception -and $e.Exception.ToString().Contains("5.7"))
			{
				#The server response was: 5.7.xx SMTP; Client was not authenticated to send anonymous mail during MAIL FROM
				Write-Verbose "$(Get-Date -Format G): Current user's credentials failed. Ask for usable credentials."

				If($Dev)
				{
					Out-File -FilePath $Script:DevErrorFile -InputObject $error -Append 4>$Null
				}

				$error.Clear()

				$emailCredentials = Get-Credential -UserName $From -Message "Enter the password to send email"

				If($UseSSL)
				{
					Send-MailMessage -Attachments $emailAttachment -Body $emailBody -BodyAsHtml -From $From `
					-Port $SmtpPort -SmtpServer $SmtpServer -Subject $emailSubject -To $To `
					-UseSSL -credential $emailCredentials *>$Null 
				}
				Else
				{
					Send-MailMessage -Attachments $emailAttachment -Body $emailBody -BodyAsHtml -From $From `
					-Port $SmtpPort -SmtpServer $SmtpServer -Subject $emailSubject -To $To `
					-credential $emailCredentials *>$Null 
				}

				If($?)
				{
					Write-Verbose "$(Get-Date -Format G): Email successfully sent using new credentials"
				}
				ElseIf(!$?)
				{
					$e = $error[0]

					Write-Verbose "$(Get-Date -Format G): Email was not sent:"
					Write-Warning "$(Get-Date -Format G): Exception: $e.Exception" 
				}
			}
			Else
			{
				Write-Verbose "$(Get-Date -Format G): Email was not sent:"
				Write-Warning "$(Get-Date -Format G): Exception: $e.Exception" 
			}
		}
	}
}
#endregion

#region ProcessDNSServer
Function ProcessDNSServer
{
	Param([string] $DNSServerName)
	#V1.20, add support for the AllDNSServers parameter	
	
	Write-Verbose "$(Get-Date -Format G): Processing DNS Server"
	Write-Verbose "$(Get-Date -Format G): `tRetrieving DNS Server Information using Server $DNSServerName"
	
	$Script:DNSServerData = Get-DNSServer -ComputerName $DNSServerName -EA 0 2>$Null 3>$Null 4>$Null
	
	$DNSServerSettings    = $Script:DNSServerData.ServerSetting
	$DNSForwarders        = $Script:DNSServerData.ServerForwarder
	$DNSServerRecursion   = $Script:DNSServerData.ServerRecursion
	$DNSServerCache       = $Script:DNSServerData.ServerCache
	$DNSServerScavenging  = $Script:DNSServerData.ServerScavenging
	$DNSRootHints         = $Script:DNSServerData.ServerRootHint
	$DNSServerDiagnostics = $Script:DNSServerData.ServerDiagnostics
	
	OutputDNSServer $DNSServerSettings `
		$DNSForwarders `
		$DNSServerRecursion `
		$DNSServerCache `
		$DNSServerScavenging `
		$DNSRootHints `
		$DNSServerDiagnostics `
		$DNSServerName
}

Function OutputDNSServer
{
	Param(
		[object] $ServerSettings, 
		[object] $DNSForwarders, 
		[object] $ServerRecursion, 
		[object] $ServerCache, 
		[object] $ServerScavenging, 
		[object] $RootHints, 
		[object] $ServerDiagnostics, 
		[string] $DNSServerName)
	#V1.20, add support for the AllDNSServers parameter	
	
	#$RootHints = $RootHints | Sort-Object $RootHints.NameServer.RecordData.NameServer
	
	#V1.11 Thanks to MBS, Root Hint servers are now sorted
	#$global:saveRootHints = $RootHints
	#V1.21 MBS fixed cases where no root hint servers were in the report or all the root hint servers were duplicated
	$RHs = $( foreach( $r in $RootHints ) {
			if( $r -is [Array] )
			{
				foreach( $i in $r )
				{
					[PsCustomObject] @{
						NameServer = $i.ipaddress.hostname; 
						IPAddresss = if( $i.IPAddress.RecordType -eq 'AAAA' ) { $i.IPAddress.RecordData.IPv6Address } else { $i.IPAddress.recorddata.IPv4Address } 
					}	
				}
			}
			else
			{
				if( $r.IPAddress -is [Array] )
				{
					foreach( $i in $r.IPAddress )
					{
						[PsCustomObject] @{
							NameServer = $i.hostname; 
							IPAddresss = if( $i.RecordType -eq 'AAAA' ) { $i.RecordData.IPv6Address } else { $i.RecordData.IPv4Address } 
						}	
					}	
				}
				else
				{
					[PsCustomObject] @{
						NameServer = $r.ipaddress.hostname; 
						IPAddresss = if( $r.IPAddress.RecordType -eq 'AAAA' ) { $r.IPAddress.RecordData.IPv6Address } else { $r.IPAddress.recorddata.IPv4Address } 
					}
				}
			}
		}
	) | Sort-Object -Property NameServer

	Write-Verbose "$(Get-Date -Format G): `t`tOutput DNS Server Settings for $DNSServerName"
	$txt = "DNS Server Properties for $DNSServerName"
	If($MSWord -or $PDF)
	{
		$Selection.InsertNewPage()
		WriteWordLine 1 0 $txt
	}
	If($Text)
	{
		Line 0 $txt
	}
	If($HTML)
	{
		WriteHTMLLine 1 0 $txt
	}
	
	#Interfaces tab
	Write-Verbose "$(Get-Date -Format G): `t`t`tInterfaces"

	#coutesy of MBS
	#if the value does not exist, then All IP Addresses is selected
	$AllIPs = Get-RegistryValue "HKLM:\SYSTEM\CurrentControlSet\Services\DNS\Parameters" "ListenAddresses" $DNSServerName
	
	If($MSWord -or $PDF)
	{
		WriteWordLine 2 0 "Interfaces"
		[System.Collections.Hashtable[]] $ScriptInformation = @()
		If($Null -eq $AllIPs)
		{
			$ScriptInformation += @{ Data = "Listen on"; Value = "All IP addresses"; }
		}
		Else
		{
			$First = $True
			ForEach($ipa in $ServerSettings.ListeningIPAddress)
			{
				If($First)
				{
					$ScriptInformation += @{ Data = "Only the following IP addresses"; Value = $ipa; }
				}
				Else
				{
					$ScriptInformation += @{ Data = ""; Value = $ipa; }
				}
				$First = $False
			}
		}
		$Table = AddWordTable -Hashtable $ScriptInformation `
		-Columns Data,Value `
		-List `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 200;
		$Table.Columns.Item(2).Width = 150;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
	}
	If($Text)
	{
		Line 0 "Interfaces"
		If($Null -eq $AllIPs)
		{
			Line 1 "Listen on: "
			Line 2 "All IP addresses"
		}
		Else
		{
			Line 1 "Listen on: "
			Line 2 "Only the following IP addresses: " 
			ForEach($IP in $ServerSettings.ListeningIPAddress)
			{
				Line 2 $IP
			}
		}
		Line 0 ""
	}
	If($HTML)
	{
		WriteHTMLLine 2 0 "Interfaces"
		$rowdata = @()
		If($Null -eq $AllIPs)
		{
			$columnHeaders = @("Listen on",($htmlsilver -bor $htmlbold),"All IP addresses",$htmlwhite)
		}
		Else
		{
			$First = $True
			ForEach($ipa in $ServerSettings.ListeningIPAddress)
			{
				If($First)
				{
					$columnHeaders = @("Only the following IP addresses",($htmlsilver -bor $htmlbold),$ipa,$htmlwhite)
				}
				Else
				{
					$rowdata += @(,('',($htmlsilver -bor $htmlbold),$ipa,$htmlwhite))
				}
				$First = $False
			}
		}

		$msg = ""
		$columnWidths = @("200","150")
		FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "350"
		WriteHTMLLine 0 0 " "
	}

	#Forwarders tab
	Write-Verbose "$(Get-Date -Format G): `t`t`tForwarders"
	If($DNSForwarders.UseRootHint)
	{
		$UseRootHints = "Yes"
	}
	Else
	{
		$UseRootHints = "No"
	}

	If($MSWord -or $PDF)
	{
		WriteWordLine 2 0 "Forwarders"
		[System.Collections.Hashtable[]] $FwdWordTable = @();
	}
	If($Text)
	{
		Line 0 "Forwarders"
	}
	If($HTML)
	{
		WriteHTMLLine 2 0 "Forwarders"
		$rowdata = @()
	}
	
	ForEach($IP in $DNSForwarders.IPAddress.IPAddressToString)
	{
		$Resolved = ResolveIPtoFQDN $IP

		If($MSWord -or $PDF)
		{
			$WordTableRowHash = @{ 
			IPAddress = $IP;
			ServerFQDN = $Resolved;
			}

			$FwdWordTable += $WordTableRowHash;
		}
		If($Text)
		{
			Line 1 "IP Address`t: " $IP
			Line 1 "Server FQDN`t: " $Resolved
			Line 0 ""
		}
		If($HTML)
		{
			$rowdata += @(,(
			$Resolved,$htmlwhite,
			$IP,$htmlwhite))
		}
	}

	If($MSWord -or $PDF)
	{
		$Table = AddWordTable -Hashtable $FwdWordTable `
		-Columns ServerFQDN, IPAddress `
		-Headers "Server FQDN", "IP Address" `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 200;
		$Table.Columns.Item(2).Width = 150;
		
		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""

	}
	If($HTML)
	{
		$columnHeaders = @(
		'Server FQDN',($htmlsilver -bor $htmlbold),
		'IP Address',($htmlsilver -bor $htmlbold))

		$msg = ""
		$columnWidths = @("200","150")
		FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "350"
		WriteHTMLLine 0 0 " "
	}

	If($MSWord -or $PDF)
	{
		[System.Collections.Hashtable[]] $ScriptInformation = @()
		$ScriptInformation += @{ Data = "Number of seconds before forward queries time out"; Value = $DNSForwarders.Timeout; }
		$ScriptInformation += @{ Data = "Use root hints if no forwarders are available"; Value = $UseRootHints; }
		$Table = AddWordTable -Hashtable $ScriptInformation `
		-Columns Data,Value `
		-List `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 200;
		$Table.Columns.Item(2).Width = 150;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
	}
	If($Text)
	{
		Line 0 ""
		Line 1 "Number of seconds before forward queries time out: " $DNSForwarders.Timeout
		Line 1 "Use root hints if no forwarders are available`t : " $UseRootHints
		Line 0 ""
	}
	If($HTML)
	{
		$rowdata = @()
		$columnHeaders = @("Number of seconds before forward queries time out",($htmlsilver -bor $htmlbold),$DNSForwarders.Timeout.ToString(),$htmlwhite)
		$rowdata += @(,('Use root hints if no forwarders are available',($htmlsilver -bor $htmlbold),$UseRootHints,$htmlwhite))

		$msg = ""
		$columnWidths = @("200","150")
		FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "350"
		WriteHTMLLine 0 0 " "
	}
	#Advanced tab
	Write-Verbose "$(Get-Date -Format G): `t`t`tAdvanced"
	
	$ServerVersion = "$($ServerSettings.MajorVersion).$($ServerSettings.MinorVersion).$($ServerSettings.BuildNumber) (0x{0:X})" -f $ServerSettings.BuildNumber

	If($ServerRecursion.Enable)
	{
		$Recursion = "Not Selected"
	}
	Else
	{
		$Recursion = "Selected"
	}
	
	If($ServerSettings.BindSecondaries)
	{
		$Bind = "Selected"
	}
	Else
	{
		$Bind = "Not Selected"
	}

	If($ServerSettings.StrictFileParsing)
	{
		$FailOnLoad = "Selected"
	}
	Else
	{
		$FailOnLoad = "Not Selected"
	}
	
	If($ServerSettings.RoundRobin)
	{
		$RoundRobin = "Selected"
	}
	Else
	{
		$RoundRobin = "Not Selected"
	}

	If($ServerSettings.LocalNetPriority)
	{
		$NetMask = "Selected"
	}
	Else
	{
		$NetMask = "Not Selected"
	}
	
	If($ServerRecursion.SecureResponse -and $ServerCache.EnablePollutionProtection)
	{
		$Pollution = "Selected"
	}
	Else
	{
		$Pollution = "Not Selected"
	}
	
	If($ServerSettings.EnableDnsSec )
	{
		$DNSSEC = "Selected"
	}
	Else
	{
		$DNSSEC = "Not Selected"
	}
	
	Switch ($ServerSettings.NameCheckFlag)
	{
		0 {$NameCheck = "Strict RFC (ANSI)"; Break}
		1 {$NameCheck = "Non RFC (ANSI)"; Break}
		2 {$NameCheck = "Multibyte (UTF8)"; Break}
		3 {$NameCheck = "All names"; Break}
		Default {$NameCheck = "Unknown: NameCheckFlag Value is $($ServerSettings.NameCheckFlag)"}
	}
	
	Switch ($ServerSettings.BootMethod)
	{
		3 {$LoadZone = "From Active Directory and registry"; Break}
		2 {$LoadZone = "From registry"; Break}
		Default {$LoadZone = "Unknown: BootMethod Value is $($ServerSettings.BootMethod)"; Break}
	}
	
	If($ServerScavenging.ScavengingInterval.days -gt 0 -or  $ServerScavenging.ScavengingInterval.hours -gt 0)
	{
		$EnableScavenging = "Selected"
		If($ServerScavenging.ScavengingInterval.days -gt 0)
		{
			$ScavengingInterval = "$($ServerScavenging.ScavengingInterval.days) days"
		}
		ElseIf($ServerScavenging.ScavengingInterval.hours -gt 0)
		{
			$ScavengingInterval = "$($ServerScavenging.ScavengingInterval.hours) hours"
		}
	}
	Else
	{
		$EnableScavenging = "Not Selected"
	}
	
	If($MSWord -or $PDF)
	{
		WriteWordLine 2 0 "Advanced"
		[System.Collections.Hashtable[]] $ScriptInformation = @()
		$ScriptInformation += @{ Data = "Server version number"; Value = $ServerVersion; }
		$Table = AddWordTable -Hashtable $ScriptInformation `
		-Columns Data,Value `
		-List `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 200;
		$Table.Columns.Item(2).Width = 150;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
		$ScriptInformation = @()
		
		WriteWordLine 0 0 "Server options:"
		[System.Collections.Hashtable[]] $ScriptInformation = @()
		$ScriptInformation += @{ Data = "Disable recursion (also disables forwarders)"; Value = $Recursion; }
		$ScriptInformation += @{ Data = "Enable BIND secondaries"; Value = $Bind; }
		$ScriptInformation += @{ Data = "Fail on load if bad zone data"; Value = $FailOnLoad; }
		$ScriptInformation += @{ Data = "Enable round robin"; Value = $RoundRobin; }
		$ScriptInformation += @{ Data = "Enable netmask ordering"; Value = $NetMask; }
		$ScriptInformation += @{ Data = "Secure cache against pollution"; Value = $Pollution; }
		$ScriptInformation += @{ Data = "Enable DNSSec validation for remote responses"; Value = $DNSSEC; }
		$ScriptInformation += @{ Data = "Name checking"; Value = $NameCheck; }
		$ScriptInformation += @{ Data = "Load zone data on startup"; Value = $LoadZone; }
		$ScriptInformation += @{ Data = "Enable automatic scavenging of stale records"; Value = $EnableScavenging; }
		If($EnableScavenging -eq "Selected")
		{
			$ScriptInformation += @{ Data = "Scavenging period"; Value = $ScavengingInterval; }
		}
		$Table = AddWordTable -Hashtable $ScriptInformation `
		-Columns Data,Value `
		-List `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 200;
		$Table.Columns.Item(2).Width = 150;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
	}
	If($Text)
	{
		Line 0 "Advanced"
		Line 1 "Server version number: " $ServerVersion
		Line 0 ""
		Line 0 "Server options:"
		Line 1 "Disable recursion (also disables forwarders)`t: " $Recursion
		Line 1 "Enable BIND secondaries`t`t`t`t: " $Bind
		Line 1 "Fail on load if bad zone data`t`t`t: " $FailOnLoad
		Line 1 "Enable round robin`t`t`t`t: " $RoundRobin
		Line 1 "Enable netmask ordering`t`t`t`t: " $NetMask
		Line 1 "Secure cache against pollution`t`t`t: " $Pollution
		Line 1 "Enable DNSSec validation for remote responses`t: " $DNSSEC
		Line 0 ""
		Line 1 "Name checking`t`t`t`t`t: " $NameCheck
		Line 1 "Load zone data on startup`t`t`t: " $LoadZone
		Line 0 ""
		Line 1 "Enable automatic scavenging of stale records`t: " $EnableScavenging
		If($EnableScavenging -eq "Selected")
		{
			Line 1 "Scavenging period`t`t`t`t: " $ScavengingInterval
		}
		Line 0 ""
	}
	If($HTML)
	{
		WriteHTMLLine 2 0 "Advanced"
		$rowdata = @()
		$columnHeaders = @("Server version number",($htmlsilver -bor $htmlbold),$ServerVersion,$htmlwhite)

		$msg = ""
		$columnWidths = @("200","150")
		FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "350"
		WriteHTMLLine 0 0 " "

		WriteHTMLLine 0 0 "Server options:"
		$rowdata = @()
		$columnHeaders = @("Disable recursion (also disables forwarders)",($htmlsilver -bor $htmlbold),$Recursion,$htmlwhite)
		$rowdata += @(,('Enable BIND secondaries',($htmlsilver -bor $htmlbold),$Bind,$htmlwhite))
		$rowdata += @(,('Fail on load if bad zone data',($htmlsilver -bor $htmlbold),$FailOnLoad,$htmlwhite))
		$rowdata += @(,('Enable round robin',($htmlsilver -bor $htmlbold),$RoundRobin,$htmlwhite))
		$rowdata += @(,('Enable netmask ordering',($htmlsilver -bor $htmlbold),$NetMask,$htmlwhite))
		$rowdata += @(,('Secure cache against pollution',($htmlsilver -bor $htmlbold),$Pollution,$htmlwhite))
		$rowdata += @(,('Enable DNSSec validation for remote responses',($htmlsilver -bor $htmlbold),$DNSSEC,$htmlwhite))
		$rowdata += @(,('Name checking',($htmlsilver -bor $htmlbold),$NameCheck,$htmlwhite))
		$rowdata += @(,('Load zone data on startup',($htmlsilver -bor $htmlbold),$LoadZone,$htmlwhite))
		$rowdata += @(,('Enable automatic scavenging of stale records',($htmlsilver -bor $htmlbold),$EnableScavenging,$htmlwhite))
		If($EnableScavenging -eq "Selected")
		{
			$rowdata += @(,('Scavenging period',($htmlsilver -bor $htmlbold),$ScavengingInterval,$htmlwhite))
		}

		$msg = ""
		$columnWidths = @("200","150")
		FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "350"
		WriteHTMLLine 0 0 " "
	}
	
	#Root Hints tab
	Write-Verbose "$(Get-Date -Format G): `t`t`tRoot Hints"

	#V1.11 Thanks to MBS, Root Hint servers are now sorted and processed more efficiently
	
	If($MSWord -or $PDF)
	{
		WriteWordLine 2 0 "Root Hints"
		WriteWordLine 3 0 "Name servers:"
		[System.Collections.Hashtable[]] $RootWordTable = @();
	}
	If($Text)
	{
		Line 0 "Root Hints"
		Line 1 "Name servers:"
		Line 1 "Server Fully Qualified Domain Name  IP Address"
		Line 1 "-------------------------------------------------------"
		#       a.root-servers.net.                 255.255.255.255
		#                                           2001:503:ba3e::2:30
	}
	If($HTML)
	{
		WriteHTMLLine 2 0 "Root Hints"
		WriteHTMLLine 3 0 "Name servers:"
		$rowdata = @()
	}

	ForEach($RH in $RHs)
	{
		$cnt = 0
		$ip = $Null
		$PrvIP = $Null
		If($rh.NameServer -is [array])
		{
			$nameServer = $rh.NameServer[0]
		}
		Else
		{
			$nameServer = $rh.NameServer
		}
		$ipAddresses = $rh.IPAddresss
		ForEach( $ipAddress in $ipAddresses )
		{
			$cnt++
			$ip = $IPAddress.IPAddressToString
			
			If($PrvIP -ne $ip)
			{
				If($cnt -eq 1)
				{
					If($MSWord -or $PDF)
					{
						$WordTableRowHash = @{ 
						ServerFQDN = $NameServer;
						IPAddress = $ip;
						}
					}
					If($Text)
					{
						Line 1 "" $NameServer -NoNewLine
						Line 1 "`t    " $ip
					}
					If($HTML)
					{
						$rowdata += @(,(
						$NameServer,$htmlwhite,
						$ip,$htmlwhite))
					}
				}
				Else
				{
					If($MSWord -or $PDF)
					{
						$WordTableRowHash = @{ 
						ServerFQDN = "";
						IPAddress = $ip;
						}
					}
					If($Text)
					{
						Line 4 "    " $ip
					}
					If($HTML)
					{
						$rowdata += @(,(
						"",$htmlwhite,
						$ip,$htmlwhite))
					}
				}
			}

			If($MSWord -or $PDF)
			{
				$RootWordTable += $WordTableRowHash;
			}
			$PrvIP = $ip
		}
	}

	If($MSWord -or $PDF)
	{
		$Table = AddWordTable -Hashtable $RootWordTable `
		-Columns ServerFQDN, IPAddress `
		-Headers "Server Fully Qualified Domain Name (FQDN)", "IP Address" `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 200;
		$Table.Columns.Item(2).Width = 150;
		
		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
	}
	If($Text)
	{
		Line 0 ""
	}
	If($HTML)
	{
		$columnHeaders = @(
		'Server Fully Qualified Domain Name (FQDN)',($htmlsilver -bor $htmlbold),
		'IP Address',($htmlsilver -bor $htmlbold))

		$msg = ""
		$columnWidths = @("200","150")
		FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "350"
		WriteHTMLLine 0 0 " "
	}
	
	#Event Logging
	Write-Verbose "$(Get-Date -Format G): `t`t`tEvent Logging"
	
	Switch ($ServerDiagnostics.EventLogLevel)
	{
		0 {$LogLevel = "No events"; Break}
		1 {$LogLevel = "Errors only"; Break}
		2 {$LogLevel = "Errors and warnings"; Break}
		4 {$LogLevel = "All events"; Break}	#my value is 7, everyone else appears to be 4
		7 {$LogLevel = "All events"; Break}	#leaving as separate stmts for now just in case
		Default {$LogLevel = "Unknown: EventLogLevel Value is $($ServerDiagnostics.EventLogLevel)"; Break}
	}
	
	If($MSWord -or $PDF)
	{
		WriteWordLine 2 0 "Event Logging"
		[System.Collections.Hashtable[]] $ScriptInformation = @()
		$ScriptInformation += @{ Data = "Log the following events"; Value = $LogLevel; }
		$Table = AddWordTable -Hashtable $ScriptInformation `
		-Columns Data,Value `
		-List `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 200;
		$Table.Columns.Item(2).Width = 150;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
	}
	If($Text)
	{
		Line 0 "Event Logging"
		Line 1 "Log the following events: " $LogLevel
		Line 0 ""
	}
	If($HTML)
	{
		WriteHTMLLine 2 0 "Event Logging"
		$rowdata = @()
		$columnHeaders = @("Log the following events",($htmlsilver -bor $htmlbold),$LogLevel,$htmlwhite)

		$msg = ""
		$columnWidths = @("200","150")
		FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "350"
		WriteHTMLLine 0 0 " "
	}

}

Function ResolveIPtoFQDN
{
	Param([string]$cname)

	Write-Verbose "$(Get-Date -Format G): `t`t`t`tAttempting to resolve $cname"
	
	$ip = $CName -as [System.Net.IpAddress]
	
	If($ip)
	{
		$Result = [System.Net.Dns]::gethostentry($ip)
		
		If($? -and $Null -ne $Result)
		{
			$CName = $Result.HostName
		}
		Else
		{
			$CName = 'Unable to resolve'
		}
	}
	Return $CName
}
#endregion

#region ProcessForwardLookupZones
Function ProcessForwardLookupZones
{
	Param([string] $DNSServerName)
	
	#V1.20, add support for the AllDNSServers parameter	
	
	Write-Verbose "$(Get-Date -Format G): Processing Forward Lookup Zones"

	$txt = "Forward Lookup Zones"
	If($MSWord -or $PDF)
	{
		$Selection.InsertNewPage()
		WriteWordLine 1 0 $txt
	}
	If($Text)
	{
		Line 0 $txt
	}
	If($HTML)
	{
		WriteHTMLLine 1 0 $txt
	}

	$First = $True
	$DNSZones = $Script:DNSServerData.ServerZone | Where-Object {
	$_.IsReverseLookupZone -eq $False -and (
	$_.ZoneType -eq "Primary" -and 
	$_.ZoneName -ne "TrustAnchors" -or 
	$_.ZoneType -eq "Stub" -or 
	$_.ZoneType -eq "Secondary")}
	
	ForEach($DNSZone in $DNSZones)
	{
		If(!$First)
		{
			If($MSWord -or $PDF)
			{
				$Selection.InsertNewPage()
			}
		}
		OutputLookupZone "Forward" $DNSZone $DNSServerName
		If($Details)
		{
			ProcessLookupZoneDetails "Forward" $DNSZone $DNSServerName
		}
		$First = $False
	}
}
#endregion

#region process lookzone data
Function OutputLookupZone
{
	Param([string] $zType, [object] $DNSZone, [string] $DNSServerName)

	#V1.20, add support for the AllDNSServers parameter	
	
	Write-Verbose "$(Get-Date -Format G): `tProcessing $($DNSZone.ZoneName)"
	
	#General tab
	Write-Verbose "$(Get-Date -Format G): `t`tGeneral"
	
	#set all the variable to N/A first since some of the variables/properties do not exist for all zones and zone types
	
	$Status            = "N/A"
	$ZoneType          = "N/A"
	$Replication       = "N/A"
	$DynamicUpdate     = "N/A"
	$NorefreshInterval = "N/A"
	$RefreshInterval   = "N/A"
	$EnableScavenging  = "N/A"
	
	If($DNSZone.IsPaused -eq $False)
	{
		$Status = "Running"
	}
	Else
	{
		$Status = "Paused"
	}
	
	If($DNSZone.ZoneType -eq "Primary" -and $DNSZone.IsDsIntegrated -eq $True)
	{
		$ZoneType = "Active Directory-Integrated"
	}
	ElseIf($DNSZone.ZoneType -eq "Primary" -and $DNSZone.IsDsIntegrated -eq $False)
	{
		$ZoneType = "Primary"
	}
	ElseIf($DNSZone.ZoneType -eq "Secondary" -and $DNSZone.IsDsIntegrated -eq $False)
	{
		$ZoneType = "Secondary"
	}
	ElseIf($DNSZone.ZoneType -eq "Stub")
	{
		$ZoneType = "Stub"
	}
	
	Switch ($DNSZone.ReplicationScope)
	{
		"Forest" {$Replication = "All DNS servers in this forest"; Break}
		"Domain" {$Replication = "All DNS servers in this domain"; Break}
		"Legacy" {$Replication = "All domain controllers in this domain (for Windows 2000 compatibility"; Break}
		"None" {$Replication = "Not an Active-Directory-Integrated zone"; Break}
		Default {$Replication = "Unknown: $($DNSZone.ReplicationScope)"; Break}
	}
	
	If( ( validObject $DNSZone DynamicUpdate ) )
	{
		Switch ($DNSZone.DynamicUpdate)
		{
			"Secure" {$DynamicUpdate = "Secure only"; Break}
			"NonsecureAndSecure" {$DynamicUpdate = "Nonsecure and secure"; Break}
			"None" {$DynamicUpdate = "None"; Break}
			Default {$DynamicUpdate = "Unknown: $($DNSZone.DynamicUpdate)"; Break}
		}
	}
	
	If($DNSZone.ZoneType -eq "Primary")
	{
		$ZoneAging = Get-DnsServerZoneAging -Name $DNSZone.ZoneName -ComputerName $DNSServerName -EA 0
		
		If($Null -ne $ZoneAging)
		{
			If($ZoneAging.AgingEnabled)
			{
				$EnableScavenging = "Selected"
				If($ZoneAging.NoRefreshInterval.days -gt 0)
				{
					$NorefreshInterval = "$($ZoneAging.NoRefreshInterval.days) days"
				}
				ElseIf($ZoneAging.NoRefreshInterval.hours -gt 0)
				{
					$NorefreshInterval = "$($ZoneAging.NoRefreshInterval.hours) hours"
				}
				If($ZoneAging.RefreshInterval.days -gt 0)
				{
					$RefreshInterval = "$($ZoneAging.RefreshInterval.days) days"
				}
				ElseIf($ZoneAging.RefreshInterval.hours -gt 0)
				{
					$RefreshInterval = "$($ZoneAging.RefreshInterval.hours) hours"
				}
			}
			Else
			{
				$EnableScavenging = "Not Selected"
			}
		}
		Else
		{
			$EnableScavenging = "Unknown"
		}
		
		$ScavengeServers = @()
		
		If($ZoneAging.ScavengeServers -is [array])
		{
			ForEach($Item in $ZoneAging.ScavengeServers)
			{
				$ScavengeServers += $ZoneAging.ScavengeServers.IPAddressToString
			}
		}
		ElseIf($Null -ne $ZoneAging.ScavengeServers)
		{
			$ScavengeServers += $ZoneAging.ScavengeServers.IPAddressToString
		}
		
		If($ScavengeServers.Count -eq 0)
		{
			$ScavengeServers += "Not Configured"
		}
	}
	
	If($MSWord -or $PDF)
	{
		WriteWordLine 2 0 "$($DNSZone.ZoneName) Properties"
		WriteWordLine 3 0 "General"

		[System.Collections.Hashtable[]] $ScriptInformation = @()
		$ScriptInformation += @{ Data = "Status"; Value = $Status; }
		$ScriptInformation += @{ Data = "Type"; Value = $ZoneType; }
		$ScriptInformation += @{ Data = "Replication"; Value = $Replication; }
		If($Null -ne $DNSZone.ZoneFile)
		{
			$ScriptInformation += @{ Data = "Zone file name"; Value = $DNSZone.ZoneFile; }
		}
		ElseIf($Null -eq $DNSZone.ZoneFile -and $DNSZone.IsDsIntegrated)
		{
			$ScriptInformation += @{ Data = "Data is stored in Active Directory"; Value = "Yes"; }
		}
		$ScriptInformation += @{ Data = "Dynamic updates"; Value = $DynamicUpdate; }
		If($DNSZone.ZoneType -eq "Primary")
		{
			$ScriptInformation += @{ Data = "Scavenge stale resource records"; Value = $EnableScavenging; }
			If($EnableScavenging -eq "Selected")
			{
				$ScriptInformation += @{ Data = "No-refresh interval"; Value = $NorefreshInterval; }
				$ScriptInformation += @{ Data = "Refresh interval"; Value = $RefreshInterval; }
			}
			$ScriptInformation += @{ Data = "Scavenge servers"; Value = $ScavengeServers[0]; }
			
			$cnt = -1
			ForEach($ScavengeServer in $ScavengeServers)
			{
				$cnt++
				
				If($cnt -gt 0)
				{
					$ScriptInformation += @{ Data = ""; Value = $ScavengeServer; }
				}
			}
		}
		$Table = AddWordTable -Hashtable $ScriptInformation `
		-Columns Data,Value `
		-List `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 200;
		$Table.Columns.Item(2).Width = 200;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
	}
	If($Text)
	{
		Line 1 "$($DNSZone.ZoneName) Properties"
		Line 2 "General"
		Line 3 "Status`t`t`t`t: " $Status
		Line 3 "Type`t`t`t`t: " $ZoneType
		Line 3 "Replication`t`t`t: " $Replication
		If($Null -ne $DNSZone.ZoneFile)
		{
			Line 3 "Zone file name`t`t`t: " $DNSZone.ZoneFile
		}
		ElseIf($Null -eq $DNSZone.ZoneFile -and $DNSZone.IsDsIntegrated)
		{
			Line 3 "Data stored in Active Directory`t: " "Yes"
		}
		Line 3 "Dynamic updates`t`t`t: " $DynamicUpdate
		If($DNSZone.ZoneType -eq "Primary")
		{
			Line 3 "Scavenge stale resource records`t: " $EnableScavenging
			If($EnableScavenging -eq "Selected")
			{
				Line 3 "No-refresh interval`t`t: " $NorefreshInterval
				Line 3 "Refresh interval`t`t: " $RefreshInterval
			}
			Line 3 "Scavenge servers`t`t: " $ScavengeServers[0]
			
			$cnt = -1
			ForEach($ScavengeServer in $ScavengeServers)
			{
				$cnt++
				
				If($cnt -gt 0)
				{
					Line 7 "  " $ScavengeServer
				}
			}
		}
		Line 0 ""
	}
	If($HTML)
	{
		WriteHTMLLine 2 0 "$($DNSZone.ZoneName) Properties"
		WriteHTMLLine 3 0 "General"

		$rowdata = @()
		$columnHeaders = @("Status",($htmlsilver -bor $htmlbold),$Status,$htmlwhite)
		$rowdata += @(,('Type',($htmlsilver -bor $htmlbold),$ZoneType,$htmlwhite))
		$rowdata += @(,('Replication',($htmlsilver -bor $htmlbold),$Replication,$htmlwhite))
		If($Null -ne $DNSZone.ZoneFile)
		{
			$rowdata += @(,('Zone file name',($htmlsilver -bor $htmlbold),$DNSZone.ZoneFile,$htmlwhite))
		}
		ElseIf($Null -eq $DNSZone.ZoneFile -and $DNSZone.IsDsIntegrated)
		{
			$rowdata += @(,('Data is stored in Active Directory',($htmlsilver -bor $htmlbold),"Yes",$htmlwhite))
		}
		$rowdata += @(,('Dynamic updates',($htmlsilver -bor $htmlbold),$DynamicUpdate,$htmlwhite))
		If($DNSZone.ZoneType -eq "Primary")
		{
			$rowdata += @(,('Scavenge stale resource records',($htmlsilver -bor $htmlbold),$EnableScavenging,$htmlwhite))
			If($EnableScavenging -eq "Selected")
			{
				$rowdata += @(,('No-refresh interval',($htmlsilver -bor $htmlbold),$NorefreshInterval,$htmlwhite))
				$rowdata += @(,('Refresh interval',($htmlsilver -bor $htmlbold),$RefreshInterval,$htmlwhite))
			}
			$rowdata += @(,('Scavenge servers',($htmlsilver -bor $htmlbold),$ScavengeServers[0],$htmlwhite))
			
			$cnt = -1
			ForEach($ScavengeServer in $ScavengeServers)
			{
				$cnt++
				
				If($cnt -gt 0)
				{
					$rowdata += @(,(' ',($htmlsilver -bor $htmlbold),$ScavengeServer,$htmlwhite))
				}
			}
		}

		$msg = ""
		$columnWidths = @("200","200")
		FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "400"
		WriteHTMLLine 0 0 " "
	}

	#Start of Authority (SOA) tab
	Write-Verbose "$(Get-Date -Format G): `t`tStart of Authority (SOA)"

	$Results = Get-DnsServerResourceRecord -zonename $DNSZone.ZoneName -rrtype soa -ComputerName $DNSServerName -EA 0

	If($? -and $Null -ne $Results)
	{
		$SOA = $Results[0]
		
		If($SOA.RecordData.RefreshInterval.Days -gt 0)
		{
			$RefreshInterval = "$($SOA.RecordData.RefreshInterval.Days) days"
		}
		ElseIf($SOA.RecordData.RefreshInterval.Hours -gt 0)
		{
			$RefreshInterval = "$($SOA.RecordData.RefreshInterval.Hours) hours"
		}
		ElseIf($SOA.RecordData.RefreshInterval.Minutes -gt 0)
		{
			$RefreshInterval = "$($SOA.RecordData.RefreshInterval.Minutes) minutes"
		}
		ElseIf($SOA.RecordData.RefreshInterval.Seconds -gt 0)
		{
			$RefreshInterval = "$($SOA.RecordData.RefreshInterval.Seconds) seconds"
		}
		Else
		{
			$RefreshInterval = "Unknown"
		}
		
		If($SOA.RecordData.RetryDelay.Days -gt 0)
		{
			$RetryDelay = "$($SOA.RecordData.RetryDelay.Days) days"
		}
		ElseIf($SOA.RecordData.RetryDelay.Hours -gt 0)
		{
			$RetryDelay = "$($SOA.RecordData.RetryDelay.Hours) hours"
		}
		ElseIf($SOA.RecordData.RetryDelay.Minutes -gt 0)
		{
			$RetryDelay = "$($SOA.RecordData.RetryDelay.Minutes) minutes"
		}
		ElseIf($SOA.RecordData.RetryDelay.Seconds -gt 0)
		{
			$RetryDelay = "$($SOA.RecordData.RetryDelay.Seconds) seconds"
		}
		Else
		{
			$RetryDelay = "Unknown"
		}
		
		If($SOA.RecordData.ExpireLimit.Days -gt 0)
		{
			$ExpireLimit = "$($SOA.RecordData.ExpireLimit.Days) days"
		}
		ElseIf($SOA.RecordData.ExpireLimit.Hours -gt 0)
		{
			$ExpireLimit = "$($SOA.RecordData.ExpireLimit.Hours) hours"
		}
		ElseIf($SOA.RecordData.ExpireLimit.Minutes -gt 0)
		{
			$ExpireLimit = "$($SOA.RecordData.ExpireLimit.Minutes) minutes"
		}
		ElseIf($SOA.RecordData.ExpireLimit.Seconds -gt 0)
		{
			$ExpireLimit = "$($SOA.RecordData.ExpireLimit.Seconds) seconds"
		}
		Else
		{
			$ExpireLimit = "Unknown"
		}
		
		If($SOA.RecordData.MinimumTimeToLive.Days -gt 0)
		{
			$MinimumTTL = "$($SOA.RecordData.MinimumTimeToLive.Days) days"
		}
		ElseIf($SOA.RecordData.MinimumTimeToLive.Hours -gt 0)
		{
			$MinimumTTL = "$($SOA.RecordData.MinimumTimeToLive.Hours) hours"
		}
		ElseIf($SOA.RecordData.MinimumTimeToLive.Minutes -gt 0)
		{
			$MinimumTTL = "$($SOA.RecordData.MinimumTimeToLive.Minutes) minutes"
		}
		ElseIf($SOA.RecordData.MinimumTimeToLive.Seconds -gt 0)
		{
			$MinimumTTL = "$($SOA.RecordData.MinimumTimeToLive.Seconds) seconds"
		}
		Else
		{
			$MinimumTTL = "Unknown"
		}
		
		If($MSWord -or $PDF)
		{
			WriteWordLine 3 0 "Start of Authority (SOA)"
			[System.Collections.Hashtable[]] $ScriptInformation = @()
			$ScriptInformation += @{ Data = "Serial number"; Value = $SOA.RecordData.SerialNumber.ToString(); }
			$ScriptInformation += @{ Data = "Primary server"; Value = $SOA.RecordData.PrimaryServer; }
			$ScriptInformation += @{ Data = "Responsible person"; Value = $SOA.RecordData.ResponsiblePerson; }
			$ScriptInformation += @{ Data = "Refresh interval"; Value = $RefreshInterval; }
			$ScriptInformation += @{ Data = "Retry interval"; Value = $RetryDelay; }
			$ScriptInformation += @{ Data = "Expires after"; Value = $ExpireLimit; }
			$ScriptInformation += @{ Data = "Minimum (default) TTL"; Value = $MinimumTTL; }
			$ScriptInformation += @{ Data = "TTL for this record"; Value = $SOA.TimeToLive.ToString(); }
			$Table = AddWordTable -Hashtable $ScriptInformation `
			-Columns Data,Value `
			-List `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitFixed;

			SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Columns.Item(1).Width = 200;
			$Table.Columns.Item(2).Width = 200;

			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
			WriteWordLine 0 0 ""
		}
		If($Text)
		{
			Line 2 "Start of Authority (SOA)"
			Line 3 "Serial number`t`t`t: " $SOA.RecordData.SerialNumber.ToString()
			Line 3 "Primary server`t`t`t: " $SOA.RecordData.PrimaryServer
			Line 3 "Responsible person`t`t: " $SOA.RecordData.ResponsiblePerson
			Line 3 "Refresh interval`t`t: " $RefreshInterval
			Line 3 "Retry interval`t`t`t: " $RetryDelay
			Line 3 "Expires after`t`t`t: " $ExpireLimit
			Line 3 "Minimum (default) TTL`t`t: " $MinimumTTL
			Line 3 "TTL for this record`t`t: " $SOA.TimeToLive.ToString()
			Line 0 ""
		}
		If($HTML)
		{
			WriteHTMLLine 3 0 "Start of Authority (SOA)"
			$rowdata = @()
			$columnHeaders = @("Serial number",($htmlsilver -bor $htmlbold),$SOA.RecordData.SerialNumber.ToString(),$htmlwhite)
			$rowdata += @(,('Primary server',($htmlsilver -bor $htmlbold),$SOA.RecordData.PrimaryServer,$htmlwhite))
			$rowdata += @(,('Responsible person',($htmlsilver -bor $htmlbold),$SOA.RecordData.ResponsiblePerson,$htmlwhite))
			$rowdata += @(,('Refresh interval',($htmlsilver -bor $htmlbold),$RefreshInterval,$htmlwhite))
			$rowdata += @(,('Retry interval',($htmlsilver -bor $htmlbold),$RetryDelay,$htmlwhite))
			$rowdata += @(,('Expires after',($htmlsilver -bor $htmlbold),$ExpireLimit,$htmlwhite))
			$rowdata += @(,('Minimum (default) TTL',($htmlsilver -bor $htmlbold),$MinimumTTL,$htmlwhite))
			$rowdata += @(,('TTL for this record',($htmlsilver -bor $htmlbold),$SOA.TimeToLive.ToString(),$htmlwhite))

			$msg = ""
			$columnWidths = @("200","200")
			FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "400"
			WriteHTMLLine 0 0 " "
		}
	}
	Else
	{
		$txt1 = "Start of Authority (SOA)"
		$txt2 = "Start of Authority data could not be retrieved"
		If($MSWord -or $PDF)
		{
			WriteWordLine 3 0 $txt1
			WriteWordLine 0 0 $txt2
			WriteWordLine 0 0 ""
		}
		If($Text)
		{
			Line 1 $txt1
			Line 0 $txt2
			Line 0 ""
		}
		If($HTML)
		{
			WriteHTMLLine 3 0 $txt1
			WriteHTMLLine 0 0 $txt2
			WriteHTMLLine 0 0 " "
		}
	}
	
	#Name Servers tab
	Write-Verbose "$(Get-Date -Format G): `t`tName Servers"
	$NameServers = Get-DnsServerResourceRecord -zonename $DNSZone.ZoneName -rrtype ns -node -ComputerName $DNSServerName -EA 0

	If($? -and $Null -ne $NameServers)
	{
		#Sort name servers added in V1.11
		$NameServers = $NameServers.RecordData.NameServer
		$NameServers = $NameServers | Sort-Object
		
		If($MSWord -or $PDF)
		{
			WriteWordLine 3 0 "Name Servers"
			$NSWordTable = New-Object System.Collections.ArrayList
			## Create an array of hashtables to store references of cells that we wish to highlight after the table has been added
			$HighlightedCells = New-Object System.Collections.ArrayList
			## Seed the $Services row index from the second row
			[int] $CurrentServiceIndex = 2;
		}
		If($Text)
		{
			Line 2 "Name Servers:"
		}
		If($HTML)
		{
			WriteHTMLLine 3 0 "Name Servers"
			$rowdata = @()
		}

		ForEach($NS in $NameServers)
		{
			# fixed in V1.11 $ipAddress = ([System.Net.Dns]::gethostentry($NS.RecordData.NameServer)).AddressList.IPAddressToString
			
			Try
			{
				$ipAddress = ([System.Net.Dns]::gethostentry($NS)).AddressList.IPAddressToString
			}
			
			Catch
			{
				$ipAddress = "Unable to retrieve an IP Address"
			}
			
			If($?)
			{
				If($ipAddress -is [array])
				{
					$cnt = -1
					
					ForEach($ip in $ipAddress)
					{
						$cnt++
						
						If($MSWord -or $PDF)
						{
							If([String]::IsNullOrEmpty($ip))	#added in V1.12
							{
								$ip = "Unable to retrieve an IP Address"
								$HighlightedCells.Add(@{ Row = $CurrentServiceIndex; Column = 1; }) > $Null
								$HighlightedCells.Add(@{ Row = $CurrentServiceIndex; Column = 2; }) > $Null
							}
							$CurrentServiceIndex++;
							
							If($cnt -eq 0)
							{
								$WordTableRowHash = @{ 
								ServerFQDN = $NS	#removed in V1.11 .RecordData.NameServer;
								IPAddress = $ip;
								}
							}
							Else
							{
								$WordTableRowHash = @{ 
								ServerFQDN = $NS	#removed in V1.11 .RecordData.NameServer;
								IPAddress = $ip;
								}
							}
						}
						If($Text)
						{
							If([String]::IsNullOrEmpty($ip))	#added in V1.12
							{
								$ip = "***Unable to retrieve an IP Address***"
							}

							If($cnt -eq 0)
							{
								Line 3 "Server FQDN`t`t`t: " $NS
								Line 4 "IP Address`t`t: " $ip
							}
							Else
							{
								Line 7 "  " $ip
							}
						}
						If($HTML)
						{
							If([String]::IsNullOrEmpty($ip))	#added in V1.12
							{
								$ip = "Unable to retrieve an IP Address"
								$HTMLHighlightedCells = $htmlred
							}
							Else
							{
								$HTMLHighlightedCells = $htmlwhite
							}

							If($cnt -eq 0)
							{
								$rowdata += @(,(
								$NS,$HTMLHighlightedCells,	#removed in V1.11 .RecordData.NameServer;
								$ip,$HTMLHighlightedCells))
							}
							Else
							{
								$rowdata += @(,(
								$NS,$HTMLHighlightedCells,	#removed in V1.11 .RecordData.NameServer;
								$ip,$HTMLHighlightedCells))
							}
						}
					}
					
					If($Text)
					{
						Line 0 ""
					}
				}
				Else
				{
					If($MSWord -or $PDF)
					{
						If([String]::IsNullOrEmpty($ipAddress))	#added in V1.12
						{
							$ipAddress = "Unable to retrieve an IP Address"
							$HighlightedCells.Add(@{ Row = $CurrentServiceIndex; Column = 1; }) > $Null
							$HighlightedCells.Add(@{ Row = $CurrentServiceIndex; Column = 2; }) > $Null
						}
						$CurrentServiceIndex++;
							
						$WordTableRowHash = @{ 
						ServerFQDN = $NS	#removed in V1.11 .RecordData.NameServer;
						IPAddress = $ipAddress;
						}
					}
					If($Text)
					{
						If([String]::IsNullOrEmpty($ipAddress))	#added in V1.12
						{
							$ipAddress = "***Unable to retrieve an IP Address***"
						}

						Line 3 "Server FQDN`t`t`t: " $NS
						Line 4 "IP Address`t`t: " $ipAddress
						Line 0 ""
					}
					If($HTML)
					{
						If([String]::IsNullOrEmpty($ipAddress))	#added in V1.12
						{
							$ipAddress = "Unable to retrieve an IP Address"
							$HTMLHighlightedCells = $htmlred
						}
						Else
						{
							$HTMLHighlightedCells = $htmlwhite
						}

						$rowdata += @(,(
						$NS,$HTMLHighlightedCells,	#removed in V1.11 .RecordData.NameServer;
						$ipAddress,$HTMLHighlightedCells))
					}
				}
			}
			ElseIf(-not $?)
			{
				If($MSWord -or $PDF)
				{
					$ipAddress = "Unable to retrieve an IP Address"
					$HighlightedCells.Add(@{ Row = $CurrentServiceIndex; Column = 1; }) > $Null
					$HighlightedCells.Add(@{ Row = $CurrentServiceIndex; Column = 2; }) > $Null
					$CurrentServiceIndex++;
						
					$WordTableRowHash = @{ 
					ServerFQDN = $NS	#removed in V1.11 .RecordData.NameServer;
					IPAddress = $ipAddress;
					}
				}
				If($Text)
				{
					$ipAddress = "***Unable to retrieve an IP Address***"

					Line 3 "Server FQDN`t`t`t: " $NS
					Line 4 "IP Address`t`t: " $ipAddress
					Line 0 ""
				}
				If($HTML)
				{
					$ipAddress = "Unable to retrieve an IP Address"
					$HTMLHighlightedCells = $htmlred

					$rowdata += @(,(
					$NS,$HTMLHighlightedCells,	#removed in V1.11 .RecordData.NameServer;
					$ipAddress,$HTMLHighlightedCells))
				}
			}

			If($MSWord -or $PDF)
			{
				## Add the hash to the array
				$NSWordTable += $WordTableRowHash;
			}
		}
		If($MSWord -or $PDF)
		{
			$Table = AddWordTable -Hashtable $NSWordTable `
			-Columns ServerFQDN, IPAddress `
			-Headers "Server Fully Qualified Domain Name (FQDN)", "IP Address" `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitFixed;

			SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			## IB - Set the required highlighted cells
			If($HighlightedCells.Count -gt 0)
			{
				SetWordCellFormat -Coordinates $HighlightedCells -Table $Table -Bold -BackgroundColor $wdColorRed -Solid;
			}
			
			$Table.Columns.Item(1).Width = 200;
			$Table.Columns.Item(2).Width = 200;
			
			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
			WriteWordLine 0 0 ""
		}
		If($Text)
		{
			Line 0 ""
		}
		If($HTML)
		{
			$columnHeaders = @(
			'Server Fully Qualified Domain Name (FQDN)',($htmlsilver -bor $htmlbold),
			'IP Address',($htmlsilver -bor $htmlbold))

			$msg = ""
			$columnWidths = @("200","200")
			FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "400"
			WriteHTMLLine 0 0 " "
		}
	}
	Else
	{
		$txt1 = "Name Servers"
		$txt2 = "Name Servers data could not be retrieved"
		If($MSWord -or $PDF)
		{
			WriteWordLine 3 0 $txt1
			WriteWordLine 0 0 $txt2
			WriteWordLine 0 0 ""
		}
		If($Text)
		{
			Line 1 $txt1
			Line 0 $txt2
			Line 0 ""
		}
		If($HTML)
		{
			WriteHTMLLine 3 0 $txt1
			WriteHTMLLine 0 0 $txt2
			WriteHTMLLine 0 0 " "
		}
	}

	If($zType -eq "Forward")
	{
		#WINS tab
		Write-Verbose "$(Get-Date -Format G): `t`tWINS"
		If( ( validObject $DNSZone IsWinsEnabled ) )
		{
			If($DNSZone.IsWinsEnabled)
			{
				$WINSEnabled = "Selected"
				
				$WINS = Get-DnsServerResourceRecord -zonename $DNSZone.ZoneName -rrtype wins -ComputerName $DNSServerName -EA 0
				
				If($? -and $Null -ne $WINS)
				{
					If($WINS.RecordData.Replicate)
					{
						$WINSReplicate = "Selected"
					}
					Else
					{
						$WINSReplicate = "Not selected"
					}

					$ip = @()
					ForEach($ipAddress in $WINS.RecordData.WinsServers)
					{
						$ip += "$ipAddress"
					}
					
					If($WINS.RecordData.CacheTimeout.Days -gt 0)
					{
						$CacheTimeout = "$($WINS.RecordData.CacheTimeout.Days) days"
					}
					ElseIf($WINS.RecordData.CacheTimeout.Hours -gt 0)
					{
						$CacheTimeout = "$($WINS.RecordData.CacheTimeout.Hours) hours"
					}
					ElseIf($WINS.RecordData.CacheTimeout.Minutes -gt 0)
					{
						$CacheTimeout = "$($WINS.RecordData.CacheTimeout.Minutes) minutes"
					}
					ElseIf($WINS.RecordData.CacheTimeout.Seconds -gt 0)
					{
						$CacheTimeout = "$($WINS.RecordData.CacheTimeout.Seconds) seconds"
					}
					Else
					{
						$CacheTimeout = "Unknown"
					}

					If($WINS.RecordData.LookupTimeout.Days -gt 0)
					{
						$LookupTimeout = "$($WINS.RecordData.LookupTimeout.Days) days"
					}
					ElseIf($WINS.RecordData.LookupTimeout.Hours -gt 0)
					{
						$LookupTimeout = "$($WINS.RecordData.LookupTimeout.Hours) hours"
					}
					ElseIf($WINS.RecordData.LookupTimeout.Minutes -gt 0)
					{
						$LookupTimeout = "$($WINS.RecordData.LookupTimeout.Minutes) minutes"
					}
					ElseIf($WINS.RecordData.LookupTimeout.Seconds -gt 0)
					{
						$LookupTimeout = "$($WINS.RecordData.LookupTimeout.Seconds) seconds"
					}
					Else
					{
						$LookupTimeout = "Unknown"
					}

					If($MSWord -or $PDF)
					{
						WriteWordLine 3 0 "WINS"
						[System.Collections.Hashtable[]] $ScriptInformation = @()
						$ScriptInformation += @{ Data = "Use WINS forward lookup"; Value = $WINSEnabled; }
						$ScriptInformation += @{ Data = "Do not replicate this record"; Value = $WINSReplicate; }
						$First = $True
						ForEach($ipa in $ip)
						{
							If($First)
							{
								$ScriptInformation += @{ Data = "IP address"; Value = $ipa; }
							}
							Else
							{
								$ScriptInformation += @{ Data = ""; Value = $ipa; }
							}
							$First = $False
						}
						$ScriptInformation += @{ Data = "Cache time-out"; Value = $CacheTimeout; }
						$ScriptInformation += @{ Data = "Lookup time-out"; Value = $LookupTimeout; }
						$Table = AddWordTable -Hashtable $ScriptInformation `
						-Columns Data,Value `
						-List `
						-Format $wdTableGrid `
						-AutoFit $wdAutoFitFixed;

						SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

						$Table.Columns.Item(1).Width = 200;
						$Table.Columns.Item(2).Width = 200;

						$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

						FindWordDocumentEnd
						$Table = $Null
						WriteWordLine 0 0 ""
					}
					If($Text)
					{
						Line 2 "WINS"
						Line 3 "Use WINS forward lookup`t`t: " $WINSEnabled
						Line 3 "Do not replicate this record`t: " $WINSReplicate
						$First = $True
						ForEach($ipa in $ip)
						{
							If($First)
							{
								Line 3 "IP address`t`t`t: " $ipa
							}
							Else
							{
								Line 7 "  " $ipa
							}
							$First = $False
						}
						Line 3 "Cache time-out`t`t`t: " $CacheTimeout
						Line 3 "Lookup time-out`t`t`t: " $LookupTimeout
						Line 0 ""
					}
					If($HTML)
					{
						WriteHTMLLine 3 0 "WINS"
						$rowdata = @()
						$columnHeaders = @("Use WINS forward lookup",($htmlsilver -bor $htmlbold),$WINSEnabled,$htmlwhite)
						$rowdata += @(,('Do not replicate this record',($htmlsilver -bor $htmlbold),$WINSReplicate,$htmlwhite))
						$First = $True
						ForEach($ipa in $ip)
						{
							If($First)
							{
								$rowdata += @(,('IP address',($htmlsilver -bor $htmlbold),$ipa,$htmlwhite))
							}
							Else
							{
								$rowdata += @(,('',($htmlsilver -bor $htmlbold),$ipa,$htmlwhite))
							}
							$First = $False
						}
						$rowdata += @(,('Cache time-out',($htmlsilver -bor $htmlbold),$CacheTimeout,$htmlwhite))
						$rowdata += @(,('Lookup time-out',($htmlsilver -bor $htmlbold),$LookupTimeout,$htmlwhite))

						$msg = ""
						$columnWidths = @("200","200")
						FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "400"
						WriteHTMLLine 0 0 " "
					}
				}
				Else
				{
					$txt1 = "WINS"
					$txt2 = "Use WINS forward lookup: $WINSEnabled"
					$txt3 = "Unable to retrieve WINS details"
					If($MSWord -or $PDF)
					{
						WriteWordLine 3 0 $txt1
						WriteWordLine 0 0 $txt2
						WriteWordLine 0 0 $txt3
						WriteWordLine 0 0 ""
					}
					If($Text)
					{
						Line 1 $txt1
						Line 2 $txt2
						Line 0 $txt3
						Line 0 ""
					}
					If($HTML)
					{
						WriteHTMLLine 3 0 $txt1
						WriteHTMLLine 0 0 $txt2
						WriteHTMLLine 0 0 $txt3
						WriteHTMLLine 0 0 " "
					}
				}
			}
			Else
			{
				$WINSEnabled = "Not selected"
				If($MSWord -or $PDF)
				{
					WriteWordLine 3 0 "WINS"
					[System.Collections.Hashtable[]] $ScriptInformation = @()
					$ScriptInformation += @{ Data = "Use WINS forward lookup"; Value = $WINSEnabled; }
					$Table = AddWordTable -Hashtable $ScriptInformation `
					-Columns Data,Value `
					-List `
					-Format $wdTableGrid `
					-AutoFit $wdAutoFitFixed;

					SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

					$Table.Columns.Item(1).Width = 200;
					$Table.Columns.Item(2).Width = 200;

					$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

					FindWordDocumentEnd
					$Table = $Null
					WriteWordLine 0 0 ""
				}
				If($Text)
				{
					Line 2 "WINS"
					Line 3 "Use WINS forward lookup`t`t: " $WINSEnabled
					Line 0 ""
				}
				If($HTML)
				{
					WriteHTMLLine 3 0 "WINS"
					$rowdata = @()
					$columnHeaders = @("Use WINS forward lookup",($htmlsilver -bor $htmlbold),$WINSEnabled,$htmlwhite)

					$msg = ""
					$columnWidths = @("200","200")
					FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "400"
					WriteHTMLLine 0 0 " "
				}
			}
		}
	}
	ElseIf($zType -eq "Reverse")
	{
		#WINS-R tab
		Write-Verbose "$(Get-Date -Format G): `t`tWINS-R"

		If( ( validObject $DNSZone IsWinsEnabled ) )
		{
			If($DNSZone.IsWinsEnabled)
			{
				$WINSEnabled = "Selected"
				
				$WINS = Get-DnsServerResourceRecord -zonename $DNSZone.ZoneName -rrtype winsr -ComputerName $DNSServerName -EA 0
				
				If($? -and $Null -ne $WINS)
				{
					If($WINS.RecordData.Replicate)
					{
						$WINSReplicate = "Selected"
					}
					Else
					{
						$WINSReplicate = "Not selected"
					}

					If($WINS.RecordData.CacheTimeout.Days -gt 0)
					{
						$CacheTimeout = "$($WINS.RecordData.CacheTimeout.Days) days"
					}
					ElseIf($WINS.RecordData.CacheTimeout.Hours -gt 0)
					{
						$CacheTimeout = "$($WINS.RecordData.CacheTimeout.Hours) hours"
					}
					ElseIf($WINS.RecordData.CacheTimeout.Minutes -gt 0)
					{
						$CacheTimeout = "$($WINS.RecordData.CacheTimeout.Minutes) minutes"
					}
					ElseIf($WINS.RecordData.CacheTimeout.Seconds -gt 0)
					{
						$CacheTimeout = "$($WINS.RecordData.CacheTimeout.Seconds) seconds"
					}
					Else
					{
						$CacheTimeout = "Unknown"
					}

					If($WINS.RecordData.LookupTimeout.Days -gt 0)
					{
						$LookupTimeout = "$($WINS.RecordData.LookupTimeout.Days) days"
					}
					ElseIf($WINS.RecordData.LookupTimeout.Hours -gt 0)
					{
						$LookupTimeout = "$($WINS.RecordData.LookupTimeout.Hours) hours"
					}
					ElseIf($WINS.RecordData.LookupTimeout.Minutes -gt 0)
					{
						$LookupTimeout = "$($WINS.RecordData.LookupTimeout.Minutes) minutes"
					}
					ElseIf($WINS.RecordData.LookupTimeout.Seconds -gt 0)
					{
						$LookupTimeout = "$($WINS.RecordData.LookupTimeout.Seconds) seconds"
					}
					Else
					{
						$LookupTimeout = "Unknown"
					}

					If($MSWord -or $PDF)
					{
						WriteWordLine 3 0 "WINS-R"
						[System.Collections.Hashtable[]] $ScriptInformation = @()
						$ScriptInformation += @{ Data = "Use WINS-R lookup"; Value = $WINSEnabled; }
						$ScriptInformation += @{ Data = "Do not replicate this record"; Value = $WINS.RecordData.Replicate.ToString(); }
						$ScriptInformation += @{ Data = "Domain name to append to returned name"; Value = $WINS.RecordData.ResultDomain; }
						$ScriptInformation += @{ Data = "Cache time-out"; Value = $CacheTimeout; }
						$ScriptInformation += @{ Data = "Lookup time-out"; Value = $LookupTimeout; }
						$ScriptInformation += @{ Data = "Submit DNS domain as NetBIOS scope"; Value = $WINSReplicate; }
						$Table = AddWordTable -Hashtable $ScriptInformation `
						-Columns Data,Value `
						-List `
						-Format $wdTableGrid `
						-AutoFit $wdAutoFitFixed;

						SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

						$Table.Columns.Item(1).Width = 200;
						$Table.Columns.Item(2).Width = 200;

						$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

						FindWordDocumentEnd
						$Table = $Null
						WriteWordLine 0 0 ""
					}
					If($Text)
					{
						Line 2 "WINS-R"
						Line 3 "Use WINS forward lookup`t`t: " $WINSEnabled
						Line 3 "Do not replicate this record`t: " $WINS.RecordData.Replicate.ToString()
						Line 3 "Domain name to append`t`t: " $WINS.RecordData.ResultDomain
						Line 3 "Cache time-out`t`t`t: " $CacheTimeout
						Line 3 "Lookup time-out`t`t`t: " $LookupTimeout
						Line 3 "Submit DNS domain as NetBIOS scope: " $WINSReplicate
						Line 0 ""
					}
					If($HTML)
					{
						WriteHTMLLine 3 0 "WINS-R"
						$rowdata = @()
						$columnHeaders = @("Use WINS forward lookup",($htmlsilver -bor $htmlbold),$WINSEnabled,$htmlwhite)
						$rowdata += @(,('Do not replicate this record',($htmlsilver -bor $htmlbold),$WINS.RecordData.Replicate.ToString(),$htmlwhite))
						$rowdata += @(,('Domain name to append to returned name',($htmlsilver -bor $htmlbold),$WINS.RecordData.ResultDomain,$htmlwhite))
						$rowdata += @(,('Cache time-out',($htmlsilver -bor $htmlbold),$CacheTimeout,$htmlwhite))
						$rowdata += @(,('Lookup time-out',($htmlsilver -bor $htmlbold),$LookupTimeout,$htmlwhite))
						$rowdata += @(,('Submit DNS domain as NetBIOS scope',($htmlsilver -bor $htmlbold),$WINSReplicate,$htmlwhite))

						$msg = ""
						$columnWidths = @("200","200")
						FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "400"
						WriteHTMLLine 0 0 " "
					}
				}
				Else
				{
					$txt1 = "WINS"
					$txt2 = "Use WINS forward lookup: $WINSEnabled"
					$txt3 = "Unable to retrieve WINS details"
					If($MSWord -or $PDF)
					{
						WriteWordLine 3 0 $txt1
						WriteWordLine 0 0 $txt2
						WriteWordLine 0 0 $txt3
						WriteWordLine 0 0 ""
					}
					If($Text)
					{
						Line 1 $txt1
						Line 2 $txt2
						Line 0 $txt3
						Line 0 ""
					}
					If($HTML)
					{
						WriteHTMLLine 3 0 $txt1
						WriteHTMLLine 0 0 $txt2
						WriteHTMLLine 0 0 $txt3
						WriteHTMLLine 0 0 " "
					}
				}
			}
			Else
			{
				$WINSEnabled = "Not selected"
				If($MSWord -or $PDF)
				{
					WriteWordLine 3 0 "WINS-R"
					[System.Collections.Hashtable[]] $ScriptInformation = @()
					$ScriptInformation += @{ Data = "Use WINS-R lookup"; Value = $WINSEnabled; }
					$Table = AddWordTable -Hashtable $ScriptInformation `
					-Columns Data,Value `
					-List `
					-Format $wdTableGrid `
					-AutoFit $wdAutoFitFixed;

					SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

					$Table.Columns.Item(1).Width = 200;
					$Table.Columns.Item(2).Width = 200;

					$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

					FindWordDocumentEnd
					$Table = $Null
					WriteWordLine 0 0 ""
				}
				If($Text)
				{
					Line 2 "WINS-R"
					Line 3 "Use WINS-R lookup`t`t: " $WINSEnabled
					Line 0 ""
				}
				If($HTML)
				{
					WriteHTMLLine 3 0 "WINS-R"
					$rowdata = @()
					$columnHeaders = @("Use WINS-R lookup",($htmlsilver -bor $htmlbold),$WINSEnabled,$htmlwhite)

					$msg = ""
					$columnWidths = @("200","200")
					FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "400"
					WriteHTMLLine 0 0 " "
				}
			}
		}
	}
	
	#Zone Transfers tab
	Write-Verbose "$(Get-Date -Format G): `t`tZone Transfers"
	
	If( ( validObject $DNSZone SecureSecondaries ) )
	{
		If($DNSZone.SecureSecondaries -ne "NoTransfer")
		{
			If($DNSZone.SecureSecondaries -eq "TransferAnyServer")
			{
				$ZoneTransfer = "To any server"
			}
			ElseIf($DNSZone.SecureSecondaries -eq "TransferToZoneNameServer")
			{
				$ZoneTransfer = "Only to servers listed on the Name Servers tab"
			}
			ElseIf($DNSZone.SecureSecondaries -eq "TransferToSecureServers")
			{
				$ZoneTransfer = "Only to the following servers"
			}
			Else
			{
				$ZoneTransfer = "Unknown"
			}

			If($ZoneTransfer -eq "Only to the following servers")
			{
				$ipSecondaryServers = ""
				ForEach($ipAddress in $DNSZone.SecondaryServers)
				{
					$ipSecondaryServers += "$ipAddress"
				}
			}

			If($DNSZone.Notify -eq "NotifyServers")
			{
				$ipNotifyServers = ""
				ForEach($ipAddress in $DNSZone.NotifyServers)
				{
					$ipNotifyServers += "$ipAddress"
				}
			}
			
			If($MSWord -or $PDF)
			{
				WriteWordLine 3 0 "Zone Transfers"
				[System.Collections.Hashtable[]] $ScriptInformation = @()
				$ScriptInformation += @{ Data = "Allow zone transfers"; Value = $ZoneTransfer; }
				If($ZoneTransfer -eq "")
				{
					$First = $True
					ForEach($ipa in $ipSecondaryServers)
					{
						If($First)
						{
							$ScriptInformation += @{ Data = "Only to the following servers"; Value = $ipa; }
						}
						Else
						{
							$ScriptInformation += @{ Data = ""; Value = $ipa; }
						}
						$First = $False
					}
				}
				If($DNSZone.Notify -eq "NoNotify")
				{
					$ScriptInformation += @{ Data = "Automatically notify"; Value = "Not selected"; }
				}
				ElseIf($DNSZone.Notify -eq "Notify")
				{
					$ScriptInformation += @{ Data = "Automatically notify"; Value = "Servers listed on the Name Servers tab"; }
				}
				ElseIf($DNSZone.Notify -eq "NotifyServers")
				{
					$First = $True
					ForEach($ipa in $ipNotifyServers)
					{
						If($First)
						{
							$ScriptInformation += @{ Data = "Automatically notify the following servers"; Value = $ipa; }
						}
						Else
						{
							$ScriptInformation += @{ Data = ""; Value = $ipa; }
						}
						$First = $False
					}
				}
				$Table = AddWordTable -Hashtable $ScriptInformation `
				-Columns Data,Value `
				-List `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitFixed;

				SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

				$Table.Columns.Item(1).Width = 200;
				$Table.Columns.Item(2).Width = 200;

				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

				FindWordDocumentEnd
				$Table = $Null
				WriteWordLine 0 0 ""
			}
			If($Text)
			{
				Line 2 "Zone Transfers"
				Line 3 "Allow zone transfers`t`t: " $ZoneTransfer
				If($ZoneTransfer -eq "Only to the following servers")
				{
					ForEach($x in $ipSecondaryServers)
					{
						Line 7 "  " $x
					}
				}
				If($DNSZone.Notify -eq "NoNotify")
				{
					Line 3 "Automatically notify`t`t: Not selected"
				}
				ElseIf($DNSZone.Notify -eq "Notify")
				{
					Line 3 "Automatically notify`t`t: Servers listed on the Name Servers tab"
				}
				ElseIf($DNSZone.Notify -eq "NotifyServers")
				{
					Line 3 "Automatically notify`t`t: The following servers " 
					ForEach($x in $ipNotifyServers)
					{
						Line 7 "  " $x
					}
				}
				Line 0 ""
			}
			If($HTML)
			{
				WriteHTMLLine 3 0 "Zone Transfers"
				$rowdata = @()
				$columnHeaders = @("Allow zone transfers",($htmlsilver -bor $htmlbold),$ZoneTransfer,$htmlwhite)
				If($ZoneTransfer -eq "Only to the following servers")
				{
					$First = $True
					ForEach($ipa in $ipSecondaryServers)
					{
						If($First)
						{
							$rowdata += @(,('Only to the following servers',($htmlsilver -bor $htmlbold),$ipa,$htmlwhite))
						}
						Else
						{
							$rowdata += @(,('',($htmlsilver -bor $htmlbold),$ipa,$htmlwhite))
						}
						$First = $False
					}
				}
				If($DNSZone.Notify -eq "NoNotify")
				{
					$rowdata += @(,('Automatically notify',($htmlsilver -bor $htmlbold),"Not selected",$htmlwhite))
				}
				ElseIf($DNSZone.Notify -eq "Notify")
				{
					$rowdata += @(,('Automatically notify',($htmlsilver -bor $htmlbold),"Servers listed on the Name Servers tab",$htmlwhite))
				}
				ElseIf($DNSZone.Notify -eq "NotifyServers")
				{
					$First = $True
					ForEach($ipa in $ipNotifyServers)
					{
						If($First)
						{
							$rowdata += @(,('Automatically notify the following servers',($htmlsilver -bor $htmlbold),$ipa,$htmlwhite))
						}
						Else
						{
							$rowdata += @(,('',($htmlsilver -bor $htmlbold),$ipa,$htmlwhite))
						}
						$First = $False
					}
				}
				$msg = ""
				$columnWidths = @("200","200")
				FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "400"
				WriteHTMLLine 0 0 " "
			}
		}
		Else
		{
			$ZoneTransfer = "Not selected"
			If($MSWord -or $PDF)
			{
				WriteWordLine 3 0 "Zone Transfers"
				[System.Collections.Hashtable[]] $ScriptInformation = @()
				$ScriptInformation += @{ Data = "Allow zone transfers"; Value = $ZoneTransfer; }
				$Table = AddWordTable -Hashtable $ScriptInformation `
				-Columns Data,Value `
				-List `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitFixed;

				SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

				$Table.Columns.Item(1).Width = 200;
				$Table.Columns.Item(2).Width = 200;

				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

				FindWordDocumentEnd
				$Table = $Null
				WriteWordLine 0 0 ""
			}
			If($Text)
			{
				Line 2 "Zone Transfers"
				Line 3 "Allow zone transfers`t`t: " $ZoneTransfer
				Line 0 ""
			}
			If($HTML)
			{
				WriteHTMLLine 3 0 "Zone Transfers"
				$rowdata = @()
				$columnHeaders = @("Allow zone transfers",($htmlsilver -bor $htmlbold),$ZoneTransfer,$htmlwhite)

				$msg = ""
				$columnWidths = @("200","200")
				FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "400"
				WriteHTMLLine 0 0 " "
			}
		}
	}

	If($DNSZone.IsSigned -eq $True) #This is a DNSSEC Zone
	{
		Write-Verbose "$(Get-Date -Format G): `t`tDNSSEC Settings"
	
		$DNSSECSettings = Get-DnsServerDnsSecZoneSetting -ZoneName $DNSZone.ZoneName -ComputerName $DNSServerName -EA 0
		
		If($? -and $Null -ne $DNSSECSettings)
		{
			If($MSWord -or $PDF)
			{
				$Selection.InsertNewPage()
				WriteWordLine 3 0 "DNSSEC properties for $($DNSZone.ZoneName) zone"
			}
			If($Text)
			{
				Line 2 "DNSSEC properties for $($DNSZone.ZoneName) zone"
			}
			If($HTML)
			{
				WriteHTMLLine 3 0 "DNSSEC properties for $($DNSZone.ZoneName) zone"
			}

			If($MSWord -or $PDF)
			{
				WriteWordLine 4 0 "Key Master"
				[System.Collections.Hashtable[]] $ScriptInformation = @()
				$ScriptInformation += @{ Data = "This DNS server is the Key Master"; Value = $DNSSECSettings.KeyMasterServer; }
				$ScriptInformation += @{ Data = "Status"; Value = $DNSSECSettings.KeyMasterStatus; }
				$Table = AddWordTable -Hashtable $ScriptInformation `
				-Columns Data,Value `
				-List `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitFixed;

				SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

				$Table.Columns.Item(1).Width = 200;
				$Table.Columns.Item(2).Width = 200;

				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

				FindWordDocumentEnd
				$Table = $Null
				WriteWordLine 0 0 ""
			}
			If($Text)
			{
				Line 3 "Key Master"
				Line 4 "This DNS server is the Key Master`t: " $DNSSECSettings.KeyMasterServer
				Line 4 "Status`t`t`t`t`t: " $DNSSECSettings.KeyMasterStatus
				Line 0 ""
			}
			If($HTML)
			{
				WriteHTMLLine 4 0 "Key Master"
				$rowdata = @()
				$columnHeaders = @("This DNS server is the Key Master",($htmlsilver -bor $htmlbold),$DNSSECSettings.KeyMasterServer,$htmlwhite)
				$rowdata += @(,('Status',($htmlsilver -bor $htmlbold),$DNSSECSettings.KeyMasterStatus,$htmlwhite))
				$msg = ""
				$columnWidths = @("200","200")
				FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "400"
				WriteHTMLLine 0 0 " "
			}

			<# I can't find KSK or ZSK data
			If($MSWord -or $PDF)
			{
				WriteWordLine 4 0 "KSK"
			}
			If($Text)
			{
				Line 3 "KSK"
			}
			If($HTML)
			{
				WriteHTMLLine 4 0 "KSK"
			}

			If($MSWord -or $PDF)
			{
				WriteWordLine 4 0 "ZSK"
			}
			If($Text)
			{
				Line 3 "ZSK"
			}
			If($HTML)
			{
				WriteHTMLLine 4 0 "ZSK"
			}
			#>
			
			If($MSWord -or $PDF)
			{
				WriteWordLine 4 0 "Next Secure (NSEC)"
				[System.Collections.Hashtable[]] $ScriptInformation = @()
				If($DNSSECSettings.DenialOfExistence -eq "NSec3")
				{
					$ScriptInformation += @{ Data = "Use NSEC3"; Value = "True"; }
					$ScriptInformation += @{ Data = "Iterations"; Value = $DNSSECSettings.NSec3Iterations.ToString(); }
					If($DNSSECSettings.IsNSec3SaltConfigured -eq $True)
					{
						$ScriptInformation += @{ Data = "Use salt"; Value = "True"; }
						If($DNSSECSettings.NSec3UserSalt -eq "-")
						{
							$ScriptInformation += @{ Data = "Generate a random salt of length"; Value = $DNSSECSettings.NSec3RandomSaltLength.ToString(); }
						}
						Else
						{
							$ScriptInformation += @{ Data = "User given salt"; Value = $DNSSECSettings.NSec3UserSalt; }
						}
					}
					Else
					{
						$ScriptInformation += @{ Data = "Do not use salt"; Value = "True"; }
					}
				}
				Else
				{
					$ScriptInformation += @{ Data = "Use NSEC"; Value = "True"; }
				}
				$Table = AddWordTable -Hashtable $ScriptInformation `
				-Columns Data,Value `
				-List `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitFixed;

				SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

				$Table.Columns.Item(1).Width = 200;
				$Table.Columns.Item(2).Width = 200;

				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

				FindWordDocumentEnd
				$Table = $Null
				WriteWordLine 0 0 ""
			}
			If($Text)
			{
				Line 3 "Next Secure (NSEC)"
				If($DNSSECSettings.DenialOfExistence -eq "NSec3")
				{
					Line 4 "Use NSEC3`t`t`t`t: " "True"
					Line 4 "Iterations`t`t`t`t: " $DNSSECSettings.NSec3Iterations.ToString()
					If($DNSSECSettings.IsNSec3SaltConfigured -eq $True)
					{
						Line 4 "Use salt`t`t`t`t: " "True"
						If($DNSSECSettings.NSec3UserSalt -eq "-")
						{
							Line 4 "Generate a random salt of length`t: " $DNSSECSettings.NSec3RandomSaltLength.ToString()
						}
						Else
						{
							Line 4 "User given salt`t`t`t`t: " $DNSSECSettings.NSec3UserSalt
						}
					}
					Else
					{
						Line 4 "Do not use salt`t`t`t`t: " "True"
					}
				}
				Else
				{
					Line 4 "Use NSEC`t`t`t`t: " "True"
				}
				Line 0 ""
			}
			If($HTML)
			{
				WriteHTMLLine 4 0 "Next Secure (NSEC)"
				$rowdata = @()
				If($DNSSECSettings.DenialOfExistence -eq "NSec3")
				{
					$columnHeaders = @("Use NSEC3",($htmlsilver -bor $htmlbold),"True",$htmlwhite)
					$rowdata += @(,('Iterations',($htmlsilver -bor $htmlbold),$DNSSECSettings.NSec3Iterations.ToString(),$htmlwhite))
					If($DNSSECSettings.IsNSec3SaltConfigured -eq $True)
					{
						$rowdata += @(,('Use salt',($htmlsilver -bor $htmlbold),"True",$htmlwhite))
						If($DNSSECSettings.NSec3UserSalt -eq "-")
						{
							$rowdata += @(,('Generate a random salt of length',($htmlsilver -bor $htmlbold),$DNSSECSettings.NSec3RandomSaltLength.ToString(),$htmlwhite))
						}
						Else
						{
							$rowdata += @(,('User given salt',($htmlsilver -bor $htmlbold),$DNSSECSettings.NSec3UserSalt,$htmlwhite))
						}
					}
					Else
					{
						$rowdata += @(,('Do not use salt',($htmlsilver -bor $htmlbold),"True",$htmlwhite))
					}
				}
				Else
				{
					$columnHeaders = @("Use NSEC",($htmlsilver -bor $htmlbold),"True",$htmlwhite)
				}

				$msg = ""
				$columnWidths = @("200","200")
				FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "400"
				WriteHTMLLine 0 0 " "
			}
			
			If($MSWord -or $PDF)
			{
				WriteWordLine 4 0 "Trust Anchor"
				[System.Collections.Hashtable[]] $ScriptInformation = @()
				If($DNSSECSettings.DistributeTrustAnchor -contains "None")
				{
					$ScriptInformation += @{ Data = "Enable the distribution of trust anchors for this zone"; Value = "False"; }
				}
				Else
				{
					$ScriptInformation += @{ Data = "Enable the distribution of trust anchors for this zone"; Value = "True"; }
				}
				If($DNSSECSettings.EnableRfc5011KeyRollover -eq $True)
				{
					$ScriptInformation += @{ Data = "Enable automatic update of trust anchors on key rollover (RFC 5011)"; Value = "True"; }
				}
				Else
				{
					$ScriptInformation += @{ Data = "Enable automatic update of trust anchors on key rollover (RFC 5011)"; Value = "False"; }
				}
				$Table = AddWordTable -Hashtable $ScriptInformation `
				-Columns Data,Value `
				-List `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitFixed;

				SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

				$Table.Columns.Item(1).Width = 200;
				$Table.Columns.Item(2).Width = 200;

				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

				FindWordDocumentEnd
				$Table = $Null
				WriteWordLine 0 0 ""
			}
			If($Text)
			{
				Line 3 "Trust Anchor"
				If($DNSSECSettings.DistributeTrustAnchor -contains "None")
				{
					Line 4 "Enable the distribution of trust "
					Line 4 "anchors for this zone`t`t`t: " "False"
				}
				Else
				{
					Line 4 "Enable the distribution of trust "
					Line 4 "anchors for this zone`t`t`t: " "True"
				}
				If($DNSSECSettings.EnableRfc5011KeyRollover -eq $True)
				{
					Line 4 "Enable automatic update of trust "
					Line 4 "anchors on key rollover (RFC 5011)`t: " "True"
				}
				Else
				{
					Line 4 "Enable automatic update of trust "
					Line 4 "anchors on key rollover (RFC 5011)`t: " "False"
				}
				Line 0 ""
			}
			If($HTML)
			{
				WriteHTMLLine 4 0 "Trust Anchor"
				$rowdata = @()
				If($DNSSECSettings.DistributeTrustAnchor -contains "None")
				{
					$columnHeaders = @("Enable the distribution of trust anchors for this zone",($htmlsilver -bor $htmlbold),"False",$htmlwhite)
				}
				Else
				{
					$columnHeaders = @("Enable the distribution of trust anchors for this zone",($htmlsilver -bor $htmlbold),"True",$htmlwhite)
				}
				If($DNSSECSettings.EnableRfc5011KeyRollover -eq $True)
				{
					$rowdata += @(,('Enable automatic update of trust anchors on key rollover (RFC 5011)',($htmlsilver -bor $htmlbold),"True",$htmlwhite))
				}
				Else
				{
					$rowdata += @(,('Enable automatic update of trust anchors on key rollover (RFC 5011)',($htmlsilver -bor $htmlbold),"False",$htmlwhite))
				}

				$msg = ""
				$columnWidths = @("200","200")
				FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "400"
				WriteHTMLLine 0 0 " "
			}

			<#DSRecordGenerationAlgorithm   : 
			{Sha1, Sha256, Sha384} = All
			{Sha1, Sha256} = SHA-1 and SHA-256
			{Sha1, Sha384} = SHA-1 and SHA-384
			{None} = None
			{Sha1} = SHA-1
			{Sha256} = SHA-256
			{Sha256, Sha384} = SHA-256 and SHA-384
			{Sha384} = SHA-384#>

			If($DNSSECSettings.DSRecordGenerationAlgorithm -contains "Sha1" -and 
			$DNSSECSettings.DSRecordGenerationAlgorithm -contains "Sha256" -and 
			$DNSSECSettings.DSRecordGenerationAlgorithm -contains "Sha384")
			{
				 $DNSSECSettingslabel = "All"
			} 
			ElseIf($DNSSECSettings.DSRecordGenerationAlgorithm -contains "Sha1" -and 
			$DNSSECSettings.DSRecordGenerationAlgorithm -contains "Sha256")
			{
				 $DNSSECSettingslabel = " SHA-1 and SHA-256"
			} 
			ElseIf($DNSSECSettings.DSRecordGenerationAlgorithm -contains "Sha1" -and 
			$DNSSECSettings.DSRecordGenerationAlgorithm -contains "Sha384")
			{
				 $DNSSECSettingslabel = " SHA-1 and SHA-384"
			} 
			ElseIf($DNSSECSettings.DSRecordGenerationAlgorithm -contains "None")
			{
				 $DNSSECSettingslabel = "None"
			} 
			ElseIf($DNSSECSettings.DSRecordGenerationAlgorithm -contains "Sha256" -and 
			$DNSSECSettings.DSRecordGenerationAlgorithm -contains "Sha384")
			{
				 $DNSSECSettingslabel = "SHA-256 and SHA-384"
			} 
			ElseIf($DNSSECSettings.DSRecordGenerationAlgorithm -contains "Sha1")
			{
				 $DNSSECSettingslabel = "SHA-1"
			} 
			ElseIf($DNSSECSettings.DSRecordGenerationAlgorithm -contains "Sha256")
			{
				 $DNSSECSettingslabel = "SHA-256"
			} 
			ElseIf($DNSSECSettings.DSRecordGenerationAlgorithm -contains "Sha384")
			{
				 $DNSSECSettingslabel = "SHA-384"
			} 
			
			If($MSWord -or $PDF)
			{
				WriteWordLine 4 0 "Advanced"
				[System.Collections.Hashtable[]] $ScriptInformation = @()
				$ScriptInformation += @{ Data = "Signing and polling parameters"; Value = ""; }
				$ScriptInformation += @{ Data = "   DS record generation algorithm"; Value = $DNSSECSettingslabel; }
				$ScriptInformation += @{ Data = "   DS record TTL (seconds)"; Value = $DNSSECSettings.DSRecordSetTTL.TotalSeconds.ToString(); }
				$ScriptInformation += @{ Data = "   DNSKEY record TTL (seconds)"; Value = $DNSSECSettings.DnsKeyRecordSetTTL.TotalSeconds.ToString(); }
				$ScriptInformation += @{ Data = "   Signature inception (hours)"; Value = $DNSSECSettings.SignatureInceptionOffset.Hours.ToString(); }
				$ScriptInformation += @{ Data = "   Secure delegation polling pd (hours)"; Value = $DNSSECSettings.SecureDelegationPollingPeriod.Hours.ToString(); }

				$Table = AddWordTable -Hashtable $ScriptInformation `
				-Columns Data,Value `
				-List `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitFixed;

				SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

				$Table.Columns.Item(1).Width = 200;
				$Table.Columns.Item(2).Width = 200;

				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

				FindWordDocumentEnd
				$Table = $Null
				WriteWordLine 0 0 ""
			}
			If($Text)
			{
				Line 3 "Advanced"
				Line 4 "Signing and polling parameters" 
				Line 4 "   DS record generation algorithm`t: " $DNSSECSettingslabel
				Line 4 "   DS record TTL (seconds)`t`t: " $DNSSECSettings.DSRecordSetTTL.TotalSeconds.ToString()
				Line 4 "   DNSKEY record TTL (seconds)`t`t: " $DNSSECSettings.DnsKeyRecordSetTTL.TotalSeconds.ToString()
				Line 4 "   Signature inception (hours)`t`t: " $DNSSECSettings.SignatureInceptionOffset.Hours.ToString()
				Line 4 "   Secure delegation polling pd (hours)`t: " $DNSSECSettings.SecureDelegationPollingPeriod.Hours.ToString()
				Line 0 ""
			}
			If($HTML)
			{
				WriteHTMLLine 4 0 "Advanced"
				$rowdata = @()
				$columnHeaders = @("Signing and polling parameters",($htmlsilver -bor $htmlbold),"",$htmlwhite)
				$rowdata += @(,('   DS record generation algorithm',($htmlsilver -bor $htmlbold),$DNSSECSettingslabel,$htmlwhite))
				$rowdata += @(,('   DS record TTL (seconds)',($htmlsilver -bor $htmlbold),$DNSSECSettings.DSRecordSetTTL.TotalSeconds.ToString(),$htmlwhite))
				$rowdata += @(,('   DNSKEY record TTL (seconds)',($htmlsilver -bor $htmlbold),$DNSSECSettings.DnsKeyRecordSetTTL.TotalSeconds.ToString(),$htmlwhite))
				$rowdata += @(,('   Signature inception (hours)',($htmlsilver -bor $htmlbold),$DNSSECSettings.SignatureInceptionOffset.Hours.ToString(),$htmlwhite))
				$rowdata += @(,('   Secure delegation polling pd (hours)',($htmlsilver -bor $htmlbold),$DNSSECSettings.SecureDelegationPollingPeriod.Hours.ToString(),$htmlwhite))

				$msg = ""
				$columnWidths = @("200","200")
				FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "400"
				WriteHTMLLine 0 0 " "
			}
		}
		ElseIf($? -and $Null -eq $DNSSECSettings)
		{
			$txt = "There are no DNSSEC settings for zone $($DNSZone.ZoneName)"
			If($MSWord -or $PDF)
			{
				WriteWordLine 0 0 $txt
				WriteWordLine 0 0 ""
			}
			If($Text)
			{
				Line 0 $txt
				Line 0 ""
			}
			If($HTML)
			{
				WriteHTMLLine 0 0 $txt
				WriteHTMLLine 0 0 " "
			}
		}
		Else
		{
			$txt = "DNSSEC settings for zone $($DNSZone.ZoneName) could not be retrieved"
			If($MSWord -or $PDF)
			{
				WriteWordLine 0 0 $txt
				WriteWordLine 0 0 ""
			}
			If($Text)
			{
				Line 0 $txt
				Line 0 ""
			}
			If($HTML)
			{
				WriteHTMLLine 0 0 $txt
				WriteHTMLLine 0 0 " "
			}
		}
	}
}
#endregion

#region lookup zone details
Function ProcessLookupZoneDetails
{
	Param([string] $zType, [object] $DNSZone, [string] $DNSServerName)

	#V1.20, add support for the AllDNSServers parameter	
	
	Write-Verbose "$(Get-Date -Format G): `t`tProcessing details for zone $($DNSZone.ZoneName)"
	
	$ZoneDetails = Get-DNSServerResourceRecord -ZoneName $DNSZone.ZoneName -ComputerName $DNSServerName -EA 0

	If($? -and $Null -ne $ZoneDetails)
	{
		OutputLookupZoneDetails $ztype $ZoneDetails $DNSZone.ZoneName
	}
	ElseIf($? -and $Null -eq $ZoneDetails)
	{
		$txt = "There are no Resource Records for zone $($DNSZone.ZoneName)"
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 0 $txt
			WriteWordLine 0 0 ""
		}
		If($Text)
		{
			Line 0 $txt
			Line 0 ""
		}
		If($HTML)
		{
			WriteHTMLLine 0 0 $txt
			WriteHTMLLine 0 0 " "
		}
	}
	Else
	{
		$txt = "Resource Records for zone $($DNSZone.ZoneName) could not be retrieved"
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 0 $txt
			WriteWordLine 0 0 ""
		}
		If($Text)
		{
			Line 0 $txt
			Line 0 ""
		}
		If($HTML)
		{
			WriteHTMLLine 0 0 $txt
			WriteHTMLLine 0 0 " "
		}
	}
}

Function OutputLookupZoneDetails
{
	Param([string] $zType, [object] $ZoneDetails, [string] $ZoneName)
	
	If($MSWord -or $PDF)
	{
		$Selection.InsertNewPage()
		WriteWordLine 3 0 "Resource Records"
	}
	If($Text)
	{
		Line 2 "Resource Records"
	}
	If($HTML)
	{
		WriteHTMLLine 3 0 "Resource Records"
	}

	If($MSWord -or $PDF)
	{
		[System.Collections.Hashtable[]] $WordTable = @();
	}
	If($HTML)
	{
		$rowdata = @()
	}
	
	$ipprefix = ""
	If($zType -eq "Reverse")
	{
		$tmpArray = $ZoneName.Split(".")
		If($tmpArray[2] -eq "in-addr" -or $tmpArray[2] -eq "arpa")
		{
			$tmpArray[2] = ""
		}
		If($tmpArray[1] -eq "in-addr" -or $tmpArray[1] -eq "arpa")
		{
			$tmpArray[1] = ""
		}

   		$ipprefix = "$($tmparray[2]).$($tmparray[1]).$($tmparray[0])."
	}
	
	#https://technet.microsoft.com/en-us/library/cc958958.aspx
	
	<#
		-- A (GUI)
		-- AAAA (GUI)
		-- Afsdb (GUI)
		-- Atma (GUI)
		-- CName (GUI)
		-- DhcId (GUI)
		-- DName (GUI)
		-- DnsKey (GUI)
		-- DS (GUI)
		-- Gpos (???)
		-- HInfo (GUI)
		-- Isdn (GUI)
		-- Key (GUI)
		-- Loc (???)
		-- Mb (GUI)
		-- Md (???)
		-- Mf (???)
		-- Mg (GUI)
		-- MInfo (GUI)
		-- Mr (GUI)
		-- Mx (GUI)
		-- Naptr (GUI)
		-- NasP (???)
		-- NasPtr (???)
		-- Ns (GUI)
		-- NSec (Created by DNSSEC)
		-- NSec3 (Created by DNSSEC)
		-- NSec3Param (Created by DNSSEC)
		-- NsNxt (???)
		-- Ptr (GUI)
		-- Rp (GUI)
		-- RRSig (Created by DNSSEC)
		-- Rt (GUI)
		-- Soa (GUI)
		-- Srv (GUI)
		-- Txt (GUI)
		-- Wins (Cmdlet)
		-- WinsR (Cmdlet)
		-- Wks (GUI)
		-- X25 (GUI)
	#>	
	
	ForEach($Detail in $ZoneDetails)
	{
		$tmpType = ""
		Switch ($Detail.RecordType)
		{
			"A"				{$tmpType = "HOST (A)"; Break}
			"AAAA"			{$tmpType = "IPv6 HOST (AAAA)"; Break}
			"AFSDB"			{$tmpType = "AFS Database (AFSDB)"; Break}
			"ATMA"			{$tmpType = "ATM Address (ATMA)"; Break}
			"CNAME"			{$tmpType = "Alias (CNAME)"; Break}
			"DHCID"			{$tmpType = "DHCID"; Break}
			"DNAME"			{$tmpType = "Domain Alias (DNAME)"; Break}
			"DNSKEY"		{$tmpType = "DNS KEY (DNSKEY)"; Break}
			"DS"			{$tmpType = "Delegation Signer (DS)"; Break}
			"HINFO"			{$tmpType = "Host Information (HINFO)"; Break}
			"ISDN"			{$tmpType = "ISDN"; Break}
			"KEY"			{$tmpType = "Public Key (KEY)"; Break}
			"MB"			{$tmpType = "Mailbox (MB)"; Break}
			"MG"			{$tmpType = "Mail Group (MG)"; Break}
			"MINFO"			{$tmpType = "Mailbox Information (MINFO)"; Break}
			"MR"			{$tmpType = "Renamed Mailbox (MR)"; Break}
			"MX"			{$tmpType = "Mail Exchanger (MX)"; Break}
			"NAPTR"			{$tmpType = "Naming Authority Pointer (NAPTR)"; Break}
			"NS"			{$tmpType = "Name Server (NS)"; Break}
			"NSEC"			{$tmpType = "Next Secure (NSEC)"; Break}
			"NSEC3"			{$tmpType = "Next Secure 3 (NSEC3)"; Break}
			"NSEC3PARAM"	{$tmpType = "Next Secure 3 Parameters (NSEC3PARAM)"; Break}
			"NXT"			{$tmpType = "Next Domain (NXT)"; Break}
			"PTR"			{$tmpType = "Pointer (PTR)"; Break}
			"RP"			{$tmpType = "Responsible Person (RP)"; Break}
			"RRSIG"			{$tmpType = "RR Signature (RRSIG)"; Break}
			"RT"			{$tmpType = "Route Through (RT)"; Break}
			"SIG"			{$tmpType = "Signature (SIG)"; Break}
			"SOA"			{$tmpType = "Start of Authority (SOA)"; Break}
			"SRV"			{$tmpType = "Service Location (SRV)"; Break}
			"TXT"			{$tmpType = "Text (TXT)"; Break}
			"WINS"			{$tmpType = "WINS Lookup"; Break}
			"WINSR"			{$tmpType = "WINS Reverse Lookup (WINS-R_"; Break}
			"WKS"			{$tmpType = "Well Known Services (WKS)"; Break}
			"X25"			{$tmpType = "X.25"; Break}
			Default 		{$tmpType = "Unable to determine Record Type: $($Detail.RecordType)"; Break}
		}
			
		If($zType -eq "Reverse")	#V1.09 fixed from = to -eq
		{
			If($Detail.HostName -eq "@")
			{
				$xHostName = "(same as parent folder)"
			}
			Else
			{
                If($Detail.RecordData.PtrDomainName -eq "localhost.")
                {
    				$xHostName = "127.0.0.1"
                }
                Else
                {
    				$xHostName = "$($ipprefix)$($Detail.HostName)"
                }
			}
		}
		Else
		{
			$xHostName = $Detail.HostName #V1.09 change from "" 
		}

		#The follow resource record types are obsolete and do not return any RecordData value
		# KEY, MB, MG, MINFO, MR, NXT, SIG
		# NAPTR is not obsolete but returns no data in RecordData
		
		$DetailData = ""
		If($Detail.HostName -eq "@" -and $Detail.RecordType -eq "NS")
		{
			$DetailData = $Detail.RecordData.NameServer
		}
		ElseIf($Detail.HostName -eq "@" -and $Detail.RecordType -eq "SOA")
		{
			$DetailData = "[$($Detail.RecordData.SerialNumber)], $($Detail.RecordData.PrimaryServer), $($Detail.RecordData.ResponsiblePerson)"
		}
		ElseIf($Detail.HostName -eq "@" -and $Detail.RecordType -eq "A")
		{
			$DetailData = $Detail.RecordData.IPv4Address
		}
		ElseIf($Detail.RecordType -eq "NS")
		{
			$DetailData = $Detail.RecordData.NameServer
		}
		ElseIf($Detail.RecordType -eq "A")
		{
			$DetailData = $Detail.RecordData.IPv4Address
		}
		ElseIf($Detail.RecordType -eq "AAAA")
		{
			$DetailData = $Detail.RecordData.IPv6Address
		}
		ElseIf($Detail.RecordType -eq "AFSDB")
		{
			$tmp = ""
			If($Detail.RecordData.SubType -eq 1)
			{
				$tmp = "AFS"
			}
			ElseIf($Detail.RecordData.SubType -eq 2)
			{
				$tmp = "DCE"
			}
			Else
			{
				$tmp = $Detail.RecordData.SubType
			}
			$DetailData = "[$($tmp)] $($Detail.RecordData.ServerName)"
		}
		ElseIf($Detail.RecordType -eq "ATMA")
		{
			$tmp = ""
			If($Detail.RecordData.AddressType -eq "E164")
			{
				$tmp = "E164"
			}
			ElseIf($Detail.RecordData.AddressType -eq "AESA")
			{
				$tmp = "NSAP"
			}
			$DetailData = "($($tmp)) $($Detail.RecordData.Address)"
		}
		ElseIf($Detail.RecordType -eq "CNAME")
		{
			$DetailData = $Detail.RecordData.HostNameAlias
		}
		ElseIf($Detail.RecordType -eq "DHCID")
		{
			$DetailData = $Detail.RecordData.DHCID
		}
		ElseIf($Detail.RecordType -eq "DNAME")
		{
			$DetailData = $Detail.RecordData.DomainNameAlias
		}
		ElseIf($Detail.RecordType -eq "DNSKEY")
		{
			$Crypto = ""
			Switch ($Detail.RecordData.CryptoAlgorithm) 
			{
				"ECDsaP256Sha256"	{$Crypto = "ECDSAP256/SHA-256"; Break}
				"ECDsaP384Sha384"	{$Crypto = "ECDSAP384/SHA-384"; Break}
				"RsaSha1"			{$Crypto = "RSA/SHA-1"; Break}
				"RsaSha1NSec3"		{$Crypto = "RSA/SHA-1 (NSEC)"; Break}
				"RsaSha256"			{$Crypto = "RSA/SHA-256"; Break}
				"RsaSha512"			{$Crypto = "RSA/SHA-512"; Break}
				Default 			{$Crypto = "Unknown CryptoAlgorithm: $($Detail.RecordData.CryptoAlgorithm)"; Break}
			}
			
			$DetailData = "[$($Detail.RecordData.KeyFlags)][DNSSEC][$($Crypto)][$($Detail.RecordData.KeyTag)]"
		}
		ElseIf($Detail.RecordType -eq "DS")
		{
			$Crypto = ""
			Switch ($Detail.RecordData.CryptoAlgorithm) 
			{
				"ECDsaP256Sha256"	{$Crypto = "ECDSAP256/SHA-256"; Break}
				"ECDsaP384Sha384"	{$Crypto = "ECDSAP384/SHA-384"; Break}
				"RsaSha1"			{$Crypto = "RSA/SHA-1"; Break}
				"RsaSha1NSec3"		{$Crypto = "RSA/SHA-1 (NSEC)"; Break}
				"RsaSha256"			{$Crypto = "RSA/SHA-256"; Break}
				"RsaSha512"			{$Crypto = "RSA/SHA-512"; Break}
				Default 			{$Crypto = "Unknown CryptoAlgorithm: $($Detail.RecordData.CryptoAlgorithm)"; Break}
			}
			
			$DigestType = ""
			Switch ($Detail.RecordData.DigestType)
			{
				"Sha1"		{$DigestType = "SHA-1"; Break}
				"Sha256"	{$DigestType = "SHA-256"; Break}
				"Sha384"	{$DigestType = "SHA-384"; Break}
				Default		{$DigestType = "Unknown DigestType: $($Detail.RecordData.DigestType)"; Break}
			}
			$DetailData = "[$($Detail.RecordData.KeyTag)][$($DigestType)][$($Crypto)][$($Detail.RecordData.Digest)]"
		}
		ElseIf($Detail.RecordType -eq "HINFO")
		{
			$DetailData = "$($Detail.RecordData.CPU), $($Detail.RecordData.OperatingSystem)"
		}
		ElseIf($Detail.RecordType -eq "ISDN")
		{
			$DetailData = "$($Detail.RecordData.IsdnNumber), $($Detail.RecordData.IsdnSubAddress)"
		}
		ElseIf($Detail.RecordType -eq "MB")
		{
			$DetailData = $Detail.RecordData
		}
		ElseIf($Detail.RecordType -eq "KEY")
		{
			$DetailData = $Detail.RecordData
		}
		ElseIf($Detail.RecordType -eq "MG")
		{
			$DetailData = $Detail.RecordData
		}
		ElseIf($Detail.RecordType -eq "MINFO")
		{
			$DetailData = $Detail.RecordData
		}
		ElseIf($Detail.RecordType -eq "MR")
		{
			$DetailData = $Detail.RecordData
		}
		ElseIf($Detail.RecordType -eq "MX")
		{
			$DetailData = "[$($Detail.RecordData.Preference)] $($Detail.RecordData.MailExchange)"
		}
		ElseIf($Detail.RecordType -eq "NAPTR")
		{
			$DetailData = $Detail.RecordData
		}
		ElseIf($Detail.RecordType -eq "NSEC")
		{
			$CoveredRecordTypes = ""
			
			If($Null -ne $Detail.RecordData.CoveredRecordTypes)
			{
				ForEach($Item in $Detail.RecordData.CoveredRecordTypes)
				{
					$CoveredRecordTypes += "$($Item) "
				}
			}
			
			$DetailData = "[$($Detail.RecordData.Name)][$($CoveredRecordTypes)]"
		}
		ElseIf($Detail.RecordType -eq "NSEC3")
		{
			$Crypto = ""
			Switch ($Detail.RecordData.HashAlgorithm) 
			{
				"ECDsaP256Sha256"	{$Crypto = "ECDSAP256/SHA-256"; Break}
				"ECDsaP384Sha384"	{$Crypto = "ECDSAP384/SHA-384"; Break}
				"RsaSha1"			{$Crypto = "RSA/SHA-1"; Break}
				"RsaSha1NSec3"		{$Crypto = "RSA/SHA-1 (NSEC)"; Break}
				"RsaSha256"			{$Crypto = "RSA/SHA-256"; Break}
				"RsaSha512"			{$Crypto = "RSA/SHA-512"; Break}
				Default 			{$Crypto = "Unknown CryptoAlgorithm: $($Detail.RecordData.HashAlgorithm)"; Break}
			}

			$OptOut = "NO Opt-Out"
			If($Detail.RecordData.OptOut -eq $True)
			{
				$OptOut = "YES Opt-Out"
			}

			$CoveredRecordTypes = ""
			
			If($Null -ne $Detail.RecordData.CoveredRecordTypes)
			{
				ForEach($Item in $Detail.RecordData.CoveredRecordTypes)
				{
					$CoveredRecordTypes += "$($Item) "
				}
			}
			
			$DetailData = "[$($Crypto)][$($OptOut)][$($Detail.RecordData.Iterations)][$($Detail.RecordData.Salt)][$($Detail.RecordData.NextHashedOwnerName)][$($CoveredRecordTypes)]"
		}
		ElseIf($Detail.RecordType -eq "NSEC3PARAM")
		{
			$Crypto = ""
			Switch ($Detail.RecordData.HashAlgorithm) 
			{
				"ECDsaP256Sha256"	{$Crypto = "ECDSAP256/SHA-256"; Break}
				"ECDsaP384Sha384"	{$Crypto = "ECDSAP384/SHA-384"; Break}
				"RsaSha1"			{$Crypto = "RSA/SHA-1"; Break}
				"RsaSha1NSec3"		{$Crypto = "RSA/SHA-1 (NSEC)"; Break}
				"RsaSha256"			{$Crypto = "RSA/SHA-256"; Break}
				"RsaSha512"			{$Crypto = "RSA/SHA-512"; Break}
				Default 			{$Crypto = "Unknown CryptoAlgorithm: $($Detail.RecordData.HashAlgorithm)"; Break}
			}
			
			$Timestamp = ""
			
			If($Null -eq $Detail.Timestamp )
			{
				$Timestamp = "0"
			}
			Else
			{
				$Timestamp = $Detail.Timestamp
			}
			
			$DetailData = "[$($Crypto)][$($Timestamp)][$($Detail.RecordData.Iterations)][$($Detail.RecordData.Salt)]"
		}
		ElseIf($Detail.RecordType -eq "NXT")
		{
			$DetailData = $Detail.RecordData
		}
		ElseIf($Detail.RecordType -eq "PTR")
		{
			$DetailData = $Detail.RecordData.PtrDomainName
		}
		ElseIf($Detail.RecordType -eq "RP")
		{
			$DetailData = "$($Detail.RecordData.ResponsiblePerson), $($Detail.RecordData.Description)"
		}
		ElseIf($Detail.RecordType -eq "RRSIG")
		{
			$Crypto = ""
			Switch ($Detail.RecordData.CryptoAlgorithm) 
			{
				"ECDsaP256Sha256"	{$Crypto = "ECDSAP256/SHA-256"; Break}
				"ECDsaP384Sha384"	{$Crypto = "ECDSAP384/SHA-384"; Break}
				"RsaSha1"			{$Crypto = "RSA/SHA-1"; Break}
				"RsaSha1NSec3"		{$Crypto = "RSA/SHA-1 (NSEC)"; Break}
				"RsaSha256"			{$Crypto = "RSA/SHA-256"; Break}
				"RsaSha512"			{$Crypto = "RSA/SHA-512"; Break}
				Default 			{$Crypto = "Unknown CryptoAlgorithm: $($Detail.RecordData.CryptoAlgorithm)"; Break}
			}
			
			$InceptionDate = $Detail.RecordData.SignatureInception.ToUniversalTime().ToShortDateString()
			$InceptionTime = $Detail.RecordData.SignatureInception.ToUniversalTime().ToLongTimeString()
			$ExpirationDate = $Detail.RecordData.SignatureExpiration.ToUniversalTime().ToShortDateString()
			$ExpirationTime = $Detail.RecordData.SignatureExpiration.ToUniversalTime().ToLongTimeString()
			
			$DetailData = "[$($Detail.RecordData.TypeCovered)][Inception(UTC): $($InceptionDate) $($InceptionTime)][Expiration(UTC): $($ExpirationDate) $($ExpirationTime)][$($Detail.RecordData.NameSigner)][$($Crypto)][$($Detail.RecordData.LabelCount)][$($Detail.RecordData.KeyTag)]"
		}
		ElseIf($Detail.RecordType -eq "RT")
		{
			$DetailData = "[$($Detail.RecordData.Preference)] $($Detail.RecordData.IntermediateHost)"
		}
		ElseIf($Detail.RecordType -eq "SIG")
		{
			$DetailData = $Detail.RecordData
		}
		ElseIf($Detail.RecordType -eq "SRV")
		{
			$DetailData = "[$($Detail.RecordData.Priority)][$($Detail.RecordData.Weight))][$($Detail.RecordData.Port))][$($Detail.RecordData.DomainName)]"
		}
		ElseIf($Detail.RecordType -eq "TXT")
		{
			$DetailData = "$($Detail.RecordData.DescriptiveText)"
		}
		ElseIf($Detail.RecordType -eq "WINS")
		{
			$xServer = ""
			ForEach($xData in $Detail.RecordData.WinsServers)
			{
				$xServer += "$($xData) "
			}
			$DetailData = "[$($xServer)]"
		}
		ElseIf($Detail.RecordType -eq "WINSR")
		{
			$DetailData = $Detail.RecordData.ResultDomain
		}
		ElseIf($Detail.RecordType -eq "WKS")
		{
			$xService = ""
			ForEach($xData in $Detail.RecordData.Service)
			{
				$xService += "$($xData) "
			}
			$DetailData = "[$($Detail.RecordData.InternetProtocol)] $xService"
		}
		ElseIf($Detail.RecordType -eq "X25")
		{
			$DetailData = $Detail.RecordData.PSDNAddress
		}
		Else
		{
			$DetailData = "Unknown: RR=$($Detail.RecordType), RecordData=$($Detail.RecordData)"
		}

		If($Null -eq $Detail.TimeStamp)
		{
			$TimeStamp = "Static"
		}
		Else
		{
			$TimeStamp = $Detail.TimeStamp
		}
		
		If($MSWord -or $PDF)
		{
			$WordTableRowHash = @{
			DetailHostName = $xHostName; 
			DetailType = $tmpType; 
			DetailData = $DetailData; 
			DetailTimeStamp = $TimeStamp; 
			}
			$WordTable += $WordTableRowHash;
		}
		If($Text)
		{
			Line 3 "Name`t`t: " $xHostName
			Line 3 "Type`t`t: " $tmpType
			Line 3 "Data`t`t: " $DetailData
			Line 3 "Timestamp`t: " $TimeStamp
			Line 0 ""
		}
		If($HTML)
		{
			$rowdata += @(,(
			$xHostName,$htmlwhite,
			$tmpType,$htmlwhite,
			$DetailData,$htmlwhite,
			$TimeStamp,$htmlwhite))
		}
	}
	
	If($MSWord -or $PDF)
	{
		$Table = AddWordTable -Hashtable $WordTable `
		-Columns  DetailHostName, DetailType, DetailData, DetailTimeStamp `
		-Headers  "Name", "Type", "Data", "Timestamp" `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		SetWordCellFormat -Collection $Table -Size 9 -BackgroundColor $wdColorWhite
		SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 105;
		$Table.Columns.Item(2).Width = 105;
		$Table.Columns.Item(3).Width = 155;
		$Table.Columns.Item(4).Width = 110;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
	}
	If($HTML)
	{
		$columnHeaders = @(
		'Name',($htmlsilver -bor $htmlbold),
		'Type',($htmlsilver -bor $htmlbold),
		'Data',($htmlsilver -bor $htmlbold),
		'Timestamp',($htmlsilver -bor $htmlbold)
		)

		$columnWidths = @("150","150","100","150")
		$msg = ""
		FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "550"
		WriteHTMLLine 0 0 " "
	}
}
#endregion

#region ProcessReverseLookupZones
Function ProcessReverseLookupZones
{
	Param([string] $DNSServerName)
	
	#V1.20, add support for the AllDNSServers parameter	
	
	Write-Verbose "$(Get-Date -Format G): Processing Reverse Lookup Zones"
	
	$txt = "Reverse Lookup Zones"
	If($MSWord -or $PDF)
	{
		$Selection.InsertNewPage()
		WriteWordLine 1 0 $txt
	}
	If($Text)
	{
		Line 0 $txt
	}
	If($HTML)
	{
		WriteHTMLLine 1 0 $txt
	}

	$First = $True
	$DNSZones = $Script:DNSServerData.ServerZone | Where-Object {$_.IsReverseLookupZone -eq $True}
	
	ForEach($DNSZone in $DNSZones)
	{
		If(!$First)
		{
			If($MSWord -or $PDF)
			{
				$Selection.InsertNewPage()
			}
		}
		OutputLookupZone "Reverse" $DNSZone $DNSServerName
		If($Details)
		{
			ProcessLookupZoneDetails "Reverse" $DNSZone $DNSServerName
		}
		$First = $False
	}
}
#endregion

#region ProcessTrustPoints
Function ProcessTrustPoints
{
	Param([string] $DNSServerName)
	
	#V1.20, add support for the AllDNSServers parameter	
	
	Write-Verbose "$(Get-Date -Format G): Processing Trust Points"
	
	$txt = "Trust Points"
	If($MSWord -or $PDF)
	{
		$Selection.InsertNewPage()
		WriteWordLine 1 0 $txt
	}
	If($Text)
	{
		Line 0 $txt
	}
	If($HTML)
	{
		WriteHTMLLine 1 0 $txt
	}

	$TrustPoints = Get-DNSServerTrustPoint -ComputerName $DNSServerName -EA 0
	
	If($? -and $Null -ne $TrustPoints)
	{
		ForEach($Trust in $TrustPoints)
		{
		
			$Anchors = Get-DnsServerTrustAnchor -name $Trust.TrustPointName -ComputerName $DNSServerName -EA 0
			
			If($? -and $Null -ne $Anchors)
			{
				$First = $True
				ForEach($Anchor in $Anchors)
				{
					If(!$First)
					{
						If($MSWord -or $PDF)
						{
							$Selection.InsertNewPage()
						}
					}
					OutputTrustPoint $Trust $Anchor
				}
			}
			$First = $False
		}
	}
	ElseIf($? -and $Null -eq $TrustPoints)
	{
		$txt1 = "Trust Zones"
		$txt2 = "There is no Trust Zones data"
		If($MSWord -or $PDF)
		{
			WriteWordLine 3 0 $txt1
			WriteWordLine 0 0 $txt2
			WriteWordLine 0 0 ""
		}
		If($Text)
		{
			Line 1 $txt1
			Line 0 $txt2
			Line 0 ""
		}
		If($HTML)
		{
			WriteHTMLLine 3 0 $txt1
			WriteHTMLLine 0 0 $txt2
			WriteHTMLLine 0 0 " "
		}
	}
	Else
	{
		$txt1 = "Trust Zones"
		$txt2 = "Trust Zones data could not be retrieved"
		If($MSWord -or $PDF)
		{
			WriteWordLine 3 0 $txt1
			WriteWordLine 0 0 $txt2
			WriteWordLine 0 0 ""
		}
		If($Text)
		{
			Line 1 $txt1
			Line 0 $txt2
			Line 0 ""
		}
		If($HTML)
		{
			WriteHTMLLine 3 0 $txt1
			WriteHTMLLine 0 0 $txt2
			WriteHTMLLine 0 0 " "
		}
	}
}

Function OutputTrustPoint
{
	Param([object] $Trust, [object] $Anchor)

	Write-Verbose "$(Get-Date -Format G): `tProcessing $($Trust.TrustPointName)"
	
	If($Anchor.TrustAnchorData.ZoneKey)
	{
		$ZoneKey = "Selected"
		Switch ($Anchor.TrustAnchorData.KeyProtocol)
		{
			"DnsSec" {$KeyProtocol = "DNSSEC"}
			Default {$KeyProtocol = "Unknown: Zone Key Protocol = $($Anchor.TrustAnchorData.KeyProtocol)"}
		}
	}
	Else
	{
		$ZoneKey = "Not Selected"
		$KeyProtocol = "N/A"
	}

	If($Anchor.TrustAnchorData.SecureEntryPoint)
	{
		$SEP = "Selected"
		Switch ($Anchor.TrustAnchorData.CryptoAlgorithm)
		{	
			"RsaSha1"			{$SEPAlgorithm = "RSA/SHA-1"; Break}
			"RsaSha1NSec3"		{$SEPAlgorithm = "RSA/SHA-1 (NSEC3)"; Break}
			"RsaSha256"			{$SEPAlgorithm = "RSA/SHA-256"; Break}
			"RsaSha512"			{$SEPAlgorithm = "RSA/SHA-512"; Break}
			"ECDsaP256Sha256"	{$SEPAlgorithm = "ECDSA256/SHA-256"; Break} #added in V2.00
			"ECDsaP384Sha384"	{$SEPAlgorithm = "ECDSAP384/SHA-384"; Break} #added in V2.00
			Default 			{$SEPAlgorithm = "Unknown: Algorithm = $($Anchor.TrustAnchorData.CryptoAlgorithm)"; Break}
		}
	}
	Else
	{
		$SEP = "Not Selected"
		$SEPAlgorithm = "N/A"
	}
	
	If($MSWord -or $PDF)
	{
		If($Trust.TrustPointName -eq ".")
		{
			WriteWordLine 3 0 "$($Trust.TrustPointName)(root) Properties"
		}
		Else
		{
			WriteWordLine 3 0 "$($Trust.TrustPointName) Properties"
		}
		[System.Collections.Hashtable[]] $ScriptInformation = @()
		$ScriptInformation += @{ Data = "Name"; Value = "(Same as parent folder)"; }
		$ScriptInformation += @{ Data = "Status"; Value = $Trust.TrustPointState; } # added in V2.00
		If( $Trust.PSObject.Properties[ 'TrustPointType' ] )
		{
			$ScriptInformation += @{ Data = "Type"; Value = $Trust.TrustPointType; }
		}
		$ScriptInformation += @{ Data = "Valid From"; Value = $Trust.LastActiveRefreshTime; }
		$ScriptInformation += @{ Data = "Valid To"; Value = $Trust.NextActiveRefreshTime; }
		$ScriptInformation += @{ Data = "Fully qualified domain name (FQDN)"; Value = $Trust.TrustPointName; }
		$ScriptInformation += @{ Data = "Key Tag"; Value = $Anchor.TrustAnchorData.KeyTag; }
		$ScriptInformation += @{ Data = "Zone Key"; Value = $ZoneKey; }
		$ScriptInformation += @{ Data = "Protocol"; Value = $KeyProtocol; }
		$ScriptInformation += @{ Data = "Secure Entry Point"; Value = $SEP; }
		$ScriptInformation += @{ Data = "Algorithm"; Value = $SEPAlgorithm; }
		#$ScriptInformation += @{ Data = "Delete this record when it becomes stale"; Value = "Can't find"; }
		#$ScriptInformation += @{ Data = "Record Timestamp"; Value = "Can't find"; }
		$Table = AddWordTable -Hashtable $ScriptInformation `
		-Columns Data,Value `
		-List `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 200;
		$Table.Columns.Item(2).Width = 200;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
	}
	If($Text)
	{
		If($Trust.TrustPointName -eq ".")
		{
			Line 1 "$($Trust.TrustPointName)(root) Properties"
		}
		Else
		{
			Line 1 "$($Trust.TrustPointName) Properties"
		}
		Line 2 "Name`t`t`t`t`t`t: (Same as parent folder)"
		Line 2 "Status`t`t`t`t`t`t: " $Trust.TrustPointState
		If( $Trust.PSObject.Properties[ 'TrustPointType' ] )
		{
			Line 2 "Type`t`t`t`t`t`t: " $Trust.TrustPointType # added in V2.00
		}
		Line 2 "Valid From`t`t`t`t`t: " $Trust.LastActiveRefreshTime
		Line 2 "Valid To`t`t`t`t`t: " $Trust.NextActiveRefreshTime
		Line 2 "Fully qualified domain name (FQDN)`t`t: " $Trust.TrustPointName
		Line 2 "Key Tag`t`t`t`t`t`t: " $Anchor.TrustAnchorData.KeyTag
		Line 2 "Zone Key`t`t`t`t`t: " $ZoneKey
		Line 2 "Protocol`t`t`t`t`t: " $KeyProtocol
		Line 2 "Secure Entry Point`t`t`t`t: " $SEP
		Line 2 "Algorithm`t`t`t`t`t: " $SEPAlgorithm
		#Line 2 "Delete this record when it becomes stale`t: " "Can't find"
		#Line 2 "Record Timestamp`t`t`t`t: " "Can't find"
		Line 0 ""
	}
	If($HTML)
	{
		If($Trust.TrustPointName -eq ".")
		{
			WriteHTMLLine 3 0 "$($Trust.TrustPointName)(root) Properties"
		}
		Else
		{
			WriteHTMLLine 3 0 "$($Trust.TrustPointName) Properties"
		}
		$rowdata = @()
		$columnHeaders = @("Name",($htmlsilver -bor $htmlbold),"(Same as parent folder)",$htmlwhite)
		$rowdata += @(,('Status',($htmlsilver -bor $htmlbold),$Trust.TrustPointState,$htmlwhite))
		If( $Trust.PSObject.Properties[ 'TrustPointType' ] )
		{
			$rowdata += @(,('Type',($htmlsilver -bor $htmlbold),$Trust.TrustPointType,$htmlwhite)) # added in V2.00
		}
		$rowdata += @(,('Valid From',($htmlsilver -bor $htmlbold),$Trust.LastActiveRefreshTime,$htmlwhite))
		$rowdata += @(,('Valid To',($htmlsilver -bor $htmlbold),$Trust.NextActiveRefreshTime,$htmlwhite))
		$rowdata += @(,('Fully qualified domain name (FQDN)',($htmlsilver -bor $htmlbold),$Trust.TrustPointName,$htmlwhite))
		$rowdata += @(,('Key Tag',($htmlsilver -bor $htmlbold),$Anchor.TrustAnchorData.KeyTag,$htmlwhite))
		$rowdata += @(,('Zone Key',($htmlsilver -bor $htmlbold),$ZoneKey,$htmlwhite))
		$rowdata += @(,('Protocol',($htmlsilver -bor $htmlbold),$KeyProtocol,$htmlwhite))
		$rowdata += @(,('Secure Entry Point',($htmlsilver -bor $htmlbold),$SEP,$htmlwhite))
		$rowdata += @(,('Algorithm',($htmlsilver -bor $htmlbold),$SEPAlgorithm,$htmlwhite))
		#$rowdata += @(,('Delete this record when it becomes stale',($htmlsilver -bor $htmlbold),"Can't find",$htmlwhite))
		#$rowdata += @(,('Record Timestamp',($htmlsilver -bor $htmlbold),"Can't find",$htmlwhite))

		$msg = ""
		$columnWidths = @("200","200")
		FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "400"
		WriteHTMLLine 0 0 " "
	}
}
#endregion

#region ProcessConditionalForwarders
Function ProcessConditionalForwarders
{
	Write-Verbose "$(Get-Date -Format G): Processing Conditional Forwarders"
	
	$txt = "Conditional Forwarders"
	If($MSWord -or $PDF)
	{
		$Selection.InsertNewPage()
		WriteWordLine 1 0 $txt
	}
	If($Text)
	{
		Line 0 $txt
	}
	If($HTML)
	{
		WriteHTMLLine 1 0 $txt
	}

	$First = $True
	$DNSZones = $Script:DNSServerData.ServerZone | Where-Object {$_.ZoneType -eq "Forwarder"}
	
	If($? -and $Null -ne $DNSZones)
	{
		ForEach($DNSZone in $DNSZones)
		{
			If(!$First)
			{
				If($MSWord -or $PDF)
				{
					$Selection.InsertNewPage()
				}
			}
			OutputConditionalForwarder $DNSZone
			$First = $False
		}
	}
	ElseIf($? -and $Null -eq $DNSZones)
	{
		$txt2 = "There is no Conditional Forwarders data"
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 0 $txt2
			WriteWordLine 0 0 ""
		}
		If($Text)
		{
			Line 0 $txt2
			Line 0 ""
		}
		If($HTML)
		{
			WriteHTMLLine 0 0 $txt2
			WriteHTMLLine 0 0 " "
		}
	}
	Else
	{
		$txt2 = "Conditional Forwarders data could not be retrieved"
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 0 $txt2
			WriteWordLine 0 0 ""
		}
		If($Text)
		{
			Line 0 $txt2
			Line 0 ""
		}
		If($HTML)
		{
			WriteHTMLLine 0 0 $txt2
			WriteHTMLLine 0 0 " "
		}
	}
}

Function OutputConditionalForwarder
{
	Param([object] $DNSZone)

	Write-Verbose "$(Get-Date -Format G): `tProcessing $($DNSZone.ZoneName)"
	
	#General tab
	Write-Verbose "$(Get-Date -Format G): `t`tGeneral"
	Switch ($DNSZone.ReplicationScope)
	{
		"Forest" {$Replication = "All DNS servers in this forest"; Break}
		"Domain" {$Replication = "All DNS servers in this domain"; Break}
		"Legacy" {$Replication = "All domain controllers in this domain (for Windows 2000 compatibility"; Break}
		"None" {$Replication = "Not an Active-Directory-Integrated zone"; Break}
		Default {$Replication = "Unknown: $($DNSZone.ReplicationScope)"; Break}
	}
	
	$IPAddresses = $DNSZone.MasterServers
	
	If($MSWord -or $PDF)
	{
		WriteWordLine 2 0 "$($DNSZone.ZoneName) Properties"
		WriteWordLine 3 0 "General"

		[System.Collections.Hashtable[]] $ScriptInformation = @()
		$ScriptInformation += @{ Data = "Type"; Value = "Conditional Forwarder"; }
		$ScriptInformation += @{ Data = "Replication"; Value = $Replication; }
		$ScriptInformation += @{ Data = "Number of seconds before forward queries time out"; Value = $DNSZone.ForwarderTimeout; }
		$Table = AddWordTable -Hashtable $ScriptInformation `
		-Columns Data,Value `
		-List `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 200;
		$Table.Columns.Item(2).Width = 200;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
	}
	If($Text)
	{
		Line 1 "$($DNSZone.ZoneName) Properties"
		Line 2 "General"
		Line 3 "Type`t`t`t`t`t`t: " "Conditional Forwarder"
		Line 3 "Replication`t`t`t`t`t: " $Replication
		Line 3 "# of seconds before forward queries time out`t: " $DNSZone.ForwarderTimeout
		Line 0 ""
	}
	If($HTML)
	{
		WriteHTMLLine 2 0 "$($DNSZone.ZoneName) Properties"
		WriteHTMLLine 3 0 "General"
		$rowdata = @()
		$columnheaders = @('Type',($htmlsilver -bor $htmlbold),"Conditional Forwarder",$htmlwhite)
		$rowdata += @(,('Replication',($htmlsilver -bor $htmlbold),$Replication,$htmlwhite))
		$rowdata += @(,('Number of seconds before forward queries time out',($htmlsilver -bor $htmlbold),$DNSZone.ForwarderTimeout,$htmlwhite))

		$msg = ""
		$columnWidths = @("200","200")
		FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "400"
		WriteHTMLLine 0 0 " "
	}

	If($MSWord -or $PDF)
	{
		WriteWordLine 3 0 "Master Servers"
		[System.Collections.Hashtable[]] $NSWordTable = @();
	}
	If($Text)
	{
		Line 2 "Master Servers:"
	}
	If($HTML)
	{
		WriteHTMLLine 3 0 "Master Servers"
		$rowdata = @()
	}

	ForEach($ip in $IPAddresses)
	{
		$Resolved = ResolveIPtoFQDN $IP
		
		If($MSWord -or $PDF)
		{
			$WordTableRowHash = @{ 
			IPAddress = $IP;
			ServerFQDN = $Resolved;
			}

			$NSWordTable += $WordTableRowHash;
		}
		If($Text)
		{
			Line 3 "Server FQDN`t`t`t`t`t: " $Resolved
			Line 3 "IP Address`t`t`t`t`t: " $ip.IPAddressToString
			Line 0 ""
		}
		If($HTML)
		{
			$rowdata += @(,(
			$Resolved,$htmlwhite,
			$IP,$htmlwhite))
		}
	}

	If($MSWord -or $PDF)
	{
		$Table = AddWordTable -Hashtable $NSWordTable `
		-Columns ServerFQDN, IPAddress `
		-Headers "Server FQDN", "IP Address" `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 200;
		$Table.Columns.Item(2).Width = 200;
		
		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
	}
	If($HTML)
	{
		$columnHeaders = @(
		'Server FQDN',($htmlsilver -bor $htmlbold),
		'IP Address',($htmlsilver -bor $htmlbold))

		$msg = ""
		$columnWidths = @("200","150")
		FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "350"
		WriteHTMLLine 0 0 " "
	}
}
#endregion

#region AppendixA
Function ProcessAppendixA
{
	Write-Verbose "$(Get-Date -Format G): `tProcessing DNS Server Forwarders"
	ForEach($DNSServer in $Script:DNSServerNames)
	{
		$Forwarders = Get-DNSServerForwarder -ComputerName $DNSServer -EA 0

		If($? -and $null -ne $Forwarders)
		{
			ForEach($Forwarder in $Forwarders)
			{
				ForEach($Item in $Forwarder.IPAddress)
				{
					$obj1 = [PSCustomObject] @{
						ComputerName = $DNSServer
						DNSForwarder = $Item.IPAddressToString
					}
					$null = $Script:DNSForwarders.Add($obj1)
				}
			}
		}
		ElseIf($? -and $null -eq $Forwarders)
		{
			$txt2 = "There is no DNS Forwarders data for Server $DNSServer"
			If($MSWord -or $PDF)
			{
				WriteWordLine 0 0 $txt2
				WriteWordLine 0 0 ""
			}
			If($Text)
			{
				Line 0 $txt2
				Line 0 ""
			}
			If($HTML)
			{
				WriteHTMLLine 0 0 $txt2
				WriteHTMLLine 0 0 " "
			}
		}
		Else
		{
			$txt2 = "Unable to retrieve DNS Forwarders data for Server $DNSServer"
			If($MSWord -or $PDF)
			{
				WriteWordLine 0 0 $txt2
				WriteWordLine 0 0 ""
			}
			If($Text)
			{
				Line 0 $txt2
				Line 0 ""
			}
			If($HTML)
			{
				WriteHTMLLine 0 0 $txt2
				WriteHTMLLine 0 0 " "
			}
		}
	}

	Write-Verbose "$(Get-Date -Format G): `tProcessing Zone Configuration"
	ForEach($DNSServer in $Script:DNSServerNames)
	{
		$Zones = Get-DNSServerZone -ComputerName $DNSServer -EA 0

		If($? -and $null -ne $Zones)
		{
			ForEach($Zone in $Zones)
			{
				$results = Get-DnsServerZoneAging -Name $Zone.ZoneName -ComputerName $DNSServer -EA 0
				
				If($? -and $null -ne $results)
				{
					$AgingEnabled      = $results.AgingEnabled.ToString()
					$RefreshInterval   = $results.RefreshInterval.ToString()
					$NoRefreshInterval = $results.NoRefreshInterval.ToString()
					If($Null -ne $results.ScavengeServers)
					{
						$ScavengeServers = $results.ScavengeServers.ToString()
					}
					ElseIf($Zone.IsAutoCreated -eq $True)
					{
						$ScavengeServers = "N/A"
					}
					Else
					{
						$ScavengeServers = "Not Configured"
					}
				}
				Else
				{
					$AgingEnabled      = "N/A"
					$RefreshInterval   = "N/A"
					$NoRefreshInterval = "N/A"
					$ScavengeServers   = "N/A"
				}
				
				$obj1 = [PSCustomObject] @{
					ComputerName      = $DNSServer
					ZoneName          = $Zone.ZoneName
					ZoneType          = $Zone.ZoneType
					IsDsIntegrated    = $Zone.IsDsIntegrated.ToString()
					IsSigned          = $Zone.IsSigned.ToString()
					DynamicUpdate     = $Zone.DynamicUpdate
					ReplicationScope  = $Zone.ReplicationScope
					AgingEnabled      = $AgingEnabled     
					RefreshInterval   = $RefreshInterval  
					NoRefreshInterval = $NoRefreshInterval
					ScavengeServers   = $ScavengeServers
				}
				$null = $Script:DNSZones.Add($obj1)
			}
		}
		ElseIf($? -and $null -eq $Zones)
		{
			#we shouldn't be here
			$txt2 = "There is no DNS Zones data for Server $DNSServer"
			If($MSWord -or $PDF)
			{
				WriteWordLine 0 0 $txt2
				WriteWordLine 0 0 ""
			}
			If($Text)
			{
				Line 0 $txt2
				Line 0 ""
			}
			If($HTML)
			{
				WriteHTMLLine 0 0 $txt2
				WriteHTMLLine 0 0 " "
			}
		}
		Else
		{
			$txt2 = "Unable to retrieve DNS Zone data for Server $DNSServer"
			If($MSWord -or $PDF)
			{
				WriteWordLine 0 0 $txt2
				WriteWordLine 0 0 ""
			}
			If($Text)
			{
				Line 0 $txt2
				Line 0 ""
			}
			If($HTML)
			{
				WriteHTMLLine 0 0 $txt2
				WriteHTMLLine 0 0 " "
			}
		}
	}

}

Function OutputAppendixA
{
	Write-Verbose "$(Get-Date -Format G): `tCreating Appendix A"
	
	$Script:DNSZones = $Script:DNSZones | Sort-Object ZoneName,ComputerName
	
	If($MSWord -or $PDF)
	{
		$selection.InsertNewPage()
		WriteWordLine 1 0 "Appendix A - DNS Server Configuration Items"
		WriteWordLine 0 0 ""
	}
	If($Text)
	{
		Line 0 "Appendix A - DNS Server Configuration Items"
		Line 0 ""
	}
	If($HTML)
	{
		WriteHTMLLine 1 0 "Appendix A - DNS Server Configuration Items"
	}
	
	If($Script:DNSForwarders.Count -gt 0)
	{
		Write-Verbose "$(Get-Date -Format G): `t`tOutput Server DNS Forwarders"
		If($MSWord -or $PDF)
		{
			WriteWordLine 2 0 "Forwarders"
			$Save = ""
			$First = $True
			$AppendixWordTable = @()
			ForEach($Item in $Script:DNSForwarders)
			{
				If(!$First -and $Save -ne "$($Item.ComputerName)")
				{
					$AppendixWordTable += @{ 
					ComputerName = "";
					Forwarder = "";
					}
				}

				$AppendixWordTable += @{ 
				ComputerName = $Item.ComputerName;
				Forwarder = $Item.DNSForwarder
				}
				
				$Save = "$($Item.ComputerName)"
				If($First)
				{
					$First = $False
				}
			}
			
			$Table = AddWordTable -Hashtable $AppendixWordTable `
			-Columns ComputerName, Forwarder `
			-Headers "Computer Name", "Forwarders" `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitContent;

			SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
			WriteWordLine 0 0 ""
		}
		If($Text)
		{
			Line 1 "Forwarders"
			Line 2 "Computer Name                  Forwarders     " 
			Line 2 "=============================================="
			#       123456789012345678901234567890S123456789012345
			#                                      255.255.255.255
			$Save = ""
			$First = $True
			ForEach($Item in $Script:DNSForwarders)
			{
				If(!$First -and $Save -ne "$($Item.ComputerName)")
				{
					Line 0 ""
				}

				Line 2 ( "{0,-30} {1,-15}" -f `
				$Item.ComputerName, $Item.DNSForwarder )
				
				$Save = "$($Item.ComputerName)"
				If($First)
				{
					$First = $False
				}
			}
			Line 0 ""
		}
		If($HTML)
		{
			WriteHTMLLine 2 0 "Forwarders"
			$Save = ""
			$First = $True
			$rowdata = @()
			ForEach($Item in $Script:DNSForwarders)
			{
				If(!$First -and $Save -ne "$($Item.ComputerName)")
				{
					$rowdata += @(,(
					"",$htmlwhite))
				}

				$rowdata += @(,(
				$Item.ComputerName,$htmlwhite,
				$Item.DNSForwarder,$htmlwhite))
				
				$Save = "$($Item.ComputerName)"
				If($First)
				{
					$First = $False
				}
			}
			$columnHeaders = @(
			'Computer Name',($global:htmlsb),
			'Forwarders',($global:htmlsb))

			$msg = ""
			FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders
			WriteHTMLLine 0 0 ""
		}
	}
	Else
	{
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 0 "No DNS Forwarders found"
		}
		If($Text)
		{
			Line 0 "No DNS Forwarders found"
			Line 0 ""
		}
		If($HTML)
		{
			WriteHTMLLine 0 0 "No DNS Forwarders found"
		}
	}

	If($Script:DNSZones.Count -gt 0)
	{
		Write-Verbose "$(Get-Date -Format G): `t`tOutput Zone Configuration"
		If($MSWord -or $PDF)
		{
			WriteWordLine 2 0 "Zone Configuration Part 1"
			$Save = ""
			$First = $True
			$AppendixWordTable = @()
			ForEach($Item in $Script:DNSZones)
			{
				If(!$First -and $Save -ne "$($Item.ZoneName)")
				{
					$AppendixWordTable += @{ 
						ComputerName     = "";
						ZoneName         = "";
						ZoneType         = "";
						IsDsIntegrated   = "";
						IsSigned         = "";
						DynamicUpdate    = "";
						ReplicationScope = "";
					}
				}

				$AppendixWordTable += @{ 
					ComputerName     = $Item.ComputerName;
					ZoneName         = $Item.ZoneName;
					ZoneType         = $Item.ZoneType;
					IsDsIntegrated   = $Item.IsDsIntegrated;
					IsSigned         = $Item.IsSigned;
					DynamicUpdate    = $Item.DynamicUpdate;
					ReplicationScope = $Item.ReplicationScope
				}
				
				$Save = "$($Item.ZoneName)"
				If($First)
				{
					$First = $False
				}
			}
			
			$Table = AddWordTable -Hashtable $AppendixWordTable `
			-Columns ZoneName, ComputerName, ZoneType, IsDsIntegrated, IsSigned, DynamicUpdate, ReplicationScope `
			-Headers "Zone Name", "Computer Name", "Zone Type", "AD Integrated", "Signed", "Dynamic Update", "Replication Scope" `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitContent;

			SetWordCellFormat -Collection $Table -Size 9 -BackgroundColor $wdColorWhite
			SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
			WriteWordLine 0 0 ""

			WriteWordLine 2 0 "Zone Configuration Part 2"
			$Save = ""
			$First = $True
			$AppendixWordTable = @()
			ForEach($Item in $Script:DNSZones)
			{
				If(!$First -and $Save -ne "$($Item.ZoneName)")
				{
					$AppendixWordTable += @{ 
						ComputerName      = "";
						ZoneName          = "";
						AgingEnabled      = "";
						RefreshInterval   = "";
						NoRefreshInterval = "";
						ScavengeServers   = "";
					}
				}

				$AppendixWordTable += @{ 
					ComputerName      = $Item.ComputerName;
					ZoneName          = $Item.ZoneName;
					AgingEnabled      = $Item.AgingEnabled;
					RefreshInterval   = $Item.RefreshInterval;
					NoRefreshInterval = $Item.NoRefreshInterval;
					ScavengeServers   = $Item.ScavengeServers;
				}
				
				$Save = "$($Item.ZoneName)"
				If($First)
				{
					$First = $False
				}
			}
			
			$Table = AddWordTable -Hashtable $AppendixWordTable `
			-Columns ZoneName, ComputerName, AgingEnabled, RefreshInterval, NoRefreshInterval, ScavengeServers `
			-Headers "Zone Name", "Computer Name", "Aging Enabled", "Refresh Interval", "NoRefresh Interval", "Scavenge Servers" `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitContent;

			SetWordCellFormat -Collection $Table -Size 9 -BackgroundColor $wdColorWhite
			SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
		}
		If($Text)
		{
			Line 1 "Zone Configuration"
			Line 0 ""
			Line 2 "Zone Name                      Computer Name                  Zone Type  AD         Signed Dynamic            Replication Aging    Refresh     NoRefresh   Scavenge       " 
			Line 2 "                                                                         Integrated        Update             Scope       Enabled  Interval    Interval    Servers        "
			Line 2 "=========================================================================================================================================================================="
			#       123456789012345678901234567890S123456789012345678901234567890S1234567890S1234567890S123456S123456789012345678S12345678901S1234567SS12345678901S12345678901S123456789012345
			#                                                                     Secondary  False      False  NonsecureAndSecure Forest      False    99:99:99:99 99:99:99:99 255.255.255.255
			#       0                              1                              2          3          4      5                  6           7           8           9        10
			$Save = ""
			$First = $True
			ForEach($Item in $Script:DNSZones)
			{
				If(!$First -and $Save -ne "$($Item.ZoneName)")
				{
					Line 0 ""
				}

				Line 2 ( "{0,-30} {1,-30} {2,-10} {3,-10} {4,-6} {5,-18} {6,-11} {7,-7}  {8,-11} {9,-11} {10, -15}" -f `
				$Item.ZoneName,
				$Item.ComputerName, 
				$Item.ZoneType,
				$Item.IsDsIntegrated,
				$Item.IsSigned,
				$Item.DynamicUpdate,
				$Item.ReplicationScope,
				$Item.AgingEnabled,
				$Item.RefreshInterval,
				$Item.NoRefreshInterval,
				$Item.ScavengeServers
				)
				
				$Save = "$($Item.ZoneName)"
				If($First)
				{
					$First = $False
				}
			}
		}
		If($HTML)
		{
			WriteHTMLLine 2 0 "Zone Configuration"
			$Save = ""
			$First = $True
			$rowdata = @()
			ForEach($Item in $Script:DNSZones)
			{
				If(!$First -and $Save -ne "$($Item.ZoneName)")
				{
					$rowdata += @(,(
					"",$htmlwhite))
				}

				$rowdata += @(,(
				$Item.ZoneName,$htmlwhite,
				$Item.ComputerName,$htmlwhite,
				$Item.ZoneType,$htmlwhite,
				$Item.IsDsIntegrated,$htmlwhite,
				$Item.IsSigned,$htmlwhite,
				$Item.DynamicUpdate,$htmlwhite,
				$Item.ReplicationScope,$htmlwhite,
				$Item.AgingEnabled,$htmlwhite,
				$Item.RefreshInterval,$htmlwhite,
				$Item.NoRefreshInterval,$htmlwhite,
				$Item.ScavengeServers))
				
				$Save = "$($Item.ZoneName)"
				If($First)
				{
					$First = $False
				}
			}
			$columnHeaders = @(
			'Zone Name',($global:htmlsb),
			'Computer Name',($global:htmlsb),
			'Zone Type',($global:htmlsb),
			'AD Integrated',($global:htmlsb),
			'Signed',($global:htmlsb),
			'Dynamic Update',($global:htmlsb),
			'Replication Scope',($global:htmlsb),
			'Aging Enabled',($global:htmlsb),
			'Refresh Interval',($global:htmlsb),
			'NoRefresh Interval',($global:htmlsb),
			'Scavenge Servers',($global:htmlsb))

			$msg = ""
			FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders
		}
	}
	Else
	{
		#we should never get here
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 0 "No DNS Zones found"
		}
		If($Text)
		{
			Line 0 "No DNS Zones found"
			Line 0 ""
		}
		If($HTML)
		{
			WriteHTMLLine 0 0 "No DNS Zones found"
		}
	}
}

#endregion

#region script core
#Script begins

ProcessScriptStart

#V1.20, add support for the AllDNSServers parameter
If($AllDNSServers -eq $False)
{
	[string]$Script:Title = "DNS Inventory Report for Server $ComputerName for the Domain $Script:RptDomain"
	SetFileNames "DNS Inventory Report for Server $ComputerName for the Domain $Script:RptDomain"
}
Else
{
	[string]$Script:Title = "DNS Inventory Report for All DNS Servers for the Domain $Script:RptDomain"
	SetFileNames "DNS Inventory for All DNS Servers for the Domain $Script:RptDomain"
}

ForEach($DNSServer in $Script:DNSServerNames)
{
	ProcessDNSServer $DNSServer

	ProcessForwardLookupZones $DNSServer

	ProcessReverseLookupZones $DNSServer

	ProcessTrustPoints $DNSServer

	ProcessConditionalForwarders $DNSServer
}

If($AllDNSServers -eq $True)
{
	#aded in V2.00
	Write-Verbose "$(Get-Date -Format G): Processing Appendix A"
	
	$Script:DNSForwarders = New-Object System.Collections.ArrayList 
	$Script:DNSZones      = New-Object System.Collections.ArrayList 
	$Script:DNSScavening  = New-Object System.Collections.ArrayList 

	ProcessAppendixA
	OutputAppendixA
	Write-Verbose "$(Get-Date -Format G): Finished Appendix A"
	Write-Verbose "$(Get-Date -Format G): "
}
#endregion

#region finish script
Write-Verbose "$(Get-Date -Format G): Finishing up document"
#end of document processing

$AbstractTitle = "DNS Inventory Report"
$SubjectTitle = "DNS Inventory Report"

UpdateDocumentProperties $AbstractTitle $SubjectTitle

ProcessDocumentOutput

ProcessScriptEnd

#endregion