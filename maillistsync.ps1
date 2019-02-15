#10/10/2018
#Script for Syncing Emails from SQL to Distro list in Exchange
#Only Exchange Admin's should attempt to run this script

$SYNCNAME = "All Mailing Lists"
#Name of Sync

#The following sets up the connection to Exchange, it will ask you for Exchange Admin Credentials to start the remote session
$Username = "DOMAIN\USER"
$Password = 'PASSWORD' | ConvertTo-SecureString -AsPlainText -Force
Import-PSSession $Session -DisableNameChecking
$UserCredential = New-Object System.Management.Automation.PSCredential -ArgumentList $Username,$Password
#Or you can have it prompt you for the credentials by using the following line instead. 
#The following sets up the connection to Exchange, it will ask you for Exchange Admin Credentials to start the remote session
#$UserCredential = Get-Credential

$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://EXCHANGE_FQDN/PowerShell/ -Authentication Kerberos -Credential $UserCredential
Import-PSSession $Session -DisableNameChecking

#Declare group list that the Users will be put into. 
#There is an assumption that this group exists already in exchange, if it does not you must create it first!!
$group1="DISTRO1"



#Get List of emails from SQL that need to be in this sync, plus any extras that need manually added.
#DISTRO1
$emaillist1 = (Invoke-SQLcmd -server 'SQL_SERVER' -database 'DATABASE' -Query "select email from database.dbo.maillist where listname = 'DISTRO1' order by lastname, firstname")



#Remove users from list that are currently in the list. This is being done with the update-distibutiongroupmember command.
Update-DistributionGroupMember -Identity $group1 -Members user@domain.com -Confirm:$false




#Get results of SQL and aux list and add email to specified group. 
$emaillist1 | ForEach {Add-DistributionGroupMember -Identity $group1 -Member $_.email}
Add-DistributionGroupMember -Identity $group1 -Member "specialadd@domain.com"


#required to exit SQLServer
cd c:


#Save results to \\FILESERVER\DESTINATION and email a copy to the admin
Get-DistributionGroupmember -identity $group1 | select DisplayName,PrimarySMTPAddress | Export-CSV "C:\MailLists\Export\$group1 $(get-date -f yyyy-MM-dd).log"
Send-MailMessage -SMTPServer SMTP_SERVER -To "emailadmin@domain.com" -From "EMAILSYNC@domain.com" -Subject "Distro List SYNC Complete $SYNCNAME" -Body "Sync results can be found in \\FILESERVER\DESTINATION" -Attachment "C:\MailLists\Export\$group1 $(get-date -f yyyy-MM-dd).log"

exit
