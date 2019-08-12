#Script Created by: Chris Bates 
#Version 1.0
#Date Last Updated: 8-12-18
#8-12-19: Added ability to exclude using a sec group.

#Setup O365 Connection
$SetPath = "E:\Scripts" #Path Used to Store Files, etc.
$MSOLCred = IMPORT-CLIXML "$SetPath\Creds\MSOL@tenant.onmicrosoft.com_cred.xml"
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://outlook.office365.com/powershell-liveid/" -Credential $MSOLCred -Authentication Basic -AllowRedirection
Import-PSSession $Session -AllowClobber
Connect-MsolService –Credential $MSOLCred

Import-Module activedirectory

$inactiveUsers = Get-ADUser -SearchBase "OU=Users,OU=INACTIVE,DC=cmh,DC=domain,DC=com" -Filter * | Select Name, UserPrincipalName, memberof | where {$_.UserPrincipalName -ne $null} #Grabs all users in the INACTIVE/Users OU to check
$exclusionGroup = Get-ADGroup -Identity InactiveOUScriptExclusion | Select-Object -ExpandProperty DistinguishedName #Grabs Exclusion Group for use later
$fulldateandtime = get-date -Format "MM-dd-yyyy  hh-mm tt dddd"
$logFilePath = "C:\Documents\InactiveLicensedUserLogFile$fulldateandtime.txt"
$errorLog = "C:\Documents\InactiveUserSharedCheckErrorLog$fulldateandtime.txt"

foreach($user in $inactiveUsers) {
    #Compiles Users Groups
    $userGroups = $user | Select-Object -ExpandProperty memberof
    #Compares Users Groups to see if its in the exclusion group
    if ($userGroups -match $exclusionGroupNJA){
        Add-Content -Path $logFilePath -Value "$user is in $exclusionGroupNJA. Skipping this User....."
    } else{	
    $isLicensed = Get-MsolUser -UserPrincipalName $user.UserPrincipalName | Select -ExpandProperty isLicensed #Pulls if they are licensed or not
    $UPN = $user | Select -ExpandProperty UserPrincipalName

    if($isLicensed -notlike "false" ){ #Checks if they are licensed.
        
        $assignedLicenses = (Get-MsolUser -UserPrincipalName $UPN).licenses.AccountSkuId
        Try{ #Confirm if we can even find the mailbox, this is a potential point we need to observe for UPN issues, etc.
        $mailboxType = Get-Mailbox $user.UserPrincipalName | Select UserPrincipalName, RecipientTypeDetails
        $RecipientType = $mailboxType | Select -ExpandProperty RecipientTypeDetails}
        Catch{"ERROR Finding Mailbox for $UPN" | Add-Content $errorLog}

        if($RecipientType -ne "SharedMailbox"){ #If they are licensed and are not a SharedMailbox then convert them and remove licenses after.
            Try{ #Confirm if we can convert the  mailbox, need to observe this in case there are issue. if it does error we should not remove the licenses.
            Set-Mailbox $UPN -Type "Shared"
            Add-Content -Path $logFilePath -Value "$UPN converted from $RecipientType to SharedMailbox"
            foreach($license in $assignedLicenses){
                Try{ #Confirm if we can remove licenses.
                Set-MsolUserLicense -UserPrincipalName $UPN -RemoveLicenses $license
                Add-Content -Path $logFilePath -Value "$UPN removed license $license"}
                Catch{"ERROR Removing $license from $UPN" | Add-Content $errorLog}
                } 
            }
            Catch{"ERROR converting $UPN to Shared" | Add-Content $errorLog}           

        }
        else{ #If they are a shared mailbox but still have licenses, remove the licenses.
            Add-Content -Path $logFilePath -Value "$UPN is already $RecipientType, proceeding to remove licenses"
            foreach($license in $assignedLicenses){
                Try{#Confirm if we can remove licenses.
                Set-MsolUserLicense -UserPrincipalName $UPN -RemoveLicenses $license
                Add-Content -Path $logFilePath -Value "$UPN removed license $license"}
                Catch{"ERROR Removing $license from $UPN" | Add-Content $errorLog}
            }
            
        }
    }
    else{ #Not licensed, no action needed.
        Add-Content -Path $logFilePath -Value "$UPN is not licensed. No action taken."
    }  
}
}

if(Test-Path $errorLog){ #Confirms if Error Log file was created, if it was, then send an email with it attached for review.
#generates email to user using .net smtpclient to notify them who has client Mailbox Rule forwards.
		         $emailFrom = "no-reply@domain.com"
		         $emailTo = "user@domain.com"
		         $subject = "Inactive User to Shared Mailbox Check"
				 
##########################################################
##########################################################
##### Start of Email #####################################
##########################################################
##########################################################
				 $body = @"
				 <html>

<head>
<meta http-equiv=Content-Type content="text/html; charset=windows-1252">
<meta name=Generator content="Microsoft Word 14 (filtered)">
<style>
<!--
 /* Font Definitions */
 @font-face
	{font-family:Calibri;
	panose-1:2 15 5 2 2 2 4 3 2 4;}
 /* Style Definitions */
 p.MsoNormal, li.MsoNormal, div.MsoNormal
	{margin:0in;
	margin-bottom:.0001pt;
	font-size:11.0pt;
	font-family:"Calibri","sans-serif";}
a:link, span.MsoHyperlink
	{color:blue;
	text-decoration:underline;}
a:visited, span.MsoHyperlinkFollowed
	{color:purple;
	text-decoration:underline;}
.MsoChpDefault
	{font-family:"Calibri","sans-serif";}
.MsoPapDefault
	{margin-bottom:10.0pt;
	line-height:115%;}
@page WordSection1
	{size:8.5in 11.0in;
	margin:1.0in 1.0in 1.0in 1.0in;}
div.WordSection1
	{page:WordSection1;}
-->
</style>

</head>

<body lang=EN-US link=blue vlink=purple>

<div class=WordSection1>

<p class=MsoNormal style='margin-left:.5in'>Hello,<br>
<br>
Attached you will find the error log for the Inactive User Shared Mailbox Script. Please review and process as needed.</p>
<br>

<p class=MsoNormal style='margin-left:.5in'>&nbsp;</p>


<p class=MsoNormal><span style='color:#1F497D'>&nbsp;</span></p>

<p class=MsoNormal>&nbsp;</p>

</div>

</body>

</html>
"@
##########################################################
##########################################################
##### End of Email #######################################
##########################################################
##########################################################
				 $smtpServer = "smtp.netjets.com"
		         $smtp = new-object Net.Mail.SmtpClient($smtpServer)
		         Send-MailMessage -SmtpServer $smtpServer -To $emailTo -From $emailFrom -Attachments $errorLog -Subject $subject -Body $body -BodyAsHtml
				 Add-Content $logFilePath -Value "Error Log Detected. Email Sent."
}
else{ #If no error log created, simply update normal log to show this.
    Add-Content $logFilePath -Value "No Error Log Detected. YAY!"
}
