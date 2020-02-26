################################# Exchange 2016    ####################################
####                                                                               ####
#### Script version 0.1                                                            ####
#### 26/02/2020 - Alex ter Neuzen - Buzz ICT                                       ####
#### Contact me @ info@buzz-ict.nl                                                 ####
#### See https://www.powershellworld.com for more scripts and WIKI                 ####
#######################################################################################
####                                                                               ####
#### Script to install on each Exchange server with Mailbox role a SSL Certificate ####
####                                                                               ####
#######################################################################################


#### Changelog ####
#### Atn - v0.1 25/2/2020 - Initial setup without Forms





#### Setting variables 

$CertPath = 
$SubjectName = 
$DomainName = 
$EXserver = Get-ExchangeServer | Where {($_.AdminDisplayVersion -Like "Version 15.1*") -And ($_.ServerRole -Like "*Mailbox*")} 



#### Running

Foreach ($server in $EXserver)
 {
New-ExchangeCertificate -GenerateRequest -RequestFile $CertPath -FriendlyName “$server” -SubjectName $SubjectName -DomainName $DomainName
}