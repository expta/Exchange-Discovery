<#

Features:
    1) This script Creates a TXT and CSV file with the following information:
        a) TXT file: Recipient Address Statistics
        b) CSV file: Output of everyone's SMTP proxy addresses.

Instructions:
    1) Run this from "regular" PowerShell.  Exchange Management Shell may cause problems, especially in Exchange 2010, due to PSv2.
    2) Usage: RecipientReportv5.ps1 server5.domain.local

Requirements:
    1) Exchange 2010 or 2013
    2) PowerShell 4.0

 
April 4 2015
Mike Crowley
 
http://BaselineTechnologies.com
 
#>

param(
    [parameter(Position=0,Mandatory=$true,ValueFromPipeline=$false,HelpMessage='Type the name of a Client Access Server')][string]$ExchangeFQDN
    )

if ($host.version.major -le 2) {
    Write-Host ""
    Write-Host "This script requires PowerShell 3.0 or later." -ForegroundColor Red
    Write-Host "Note: Exchange 2010's EMC always runs as version 2.  Perhaps try launching PowerShell normally." -ForegroundColor Red
    Write-Host ""
    Write-Host "Exiting..." -ForegroundColor Red
    Sleep 3
    Exit
    }


if ((Test-Connection $ExchangeFQDN -Count 1 -Quiet) -ne $true) {
    Write-Host ""
    Write-Host ("Cannot connect to: " + $ExchangeFQDN) -ForegroundColor Red
    Write-Host ""
    Write-Host "Exiting..." -ForegroundColor Red
    Sleep 3
    Exit
    }

#Misc variables
#$ExchangeFQDN = "exchserv1.domain1.local"
$ReportTimeStamp = (Get-Date -Format s) -replace ":", "."
$TxtFile = "RecipientAddressReport_Part_1of2.txt"
$CsvFile = "RecipientAddressReport_Part_2of2.csv"

#Connect to Exchange
Write-Host ("Connecting to " + $ExchangeFQDN + "...") -ForegroundColor Cyan
Get-PSSession | Where-Object {$_.ConfigurationName -eq 'Microsoft.Exchange'} | Remove-PSSession
$Session = @{
    ConfigurationName = 'Microsoft.Exchange'
    ConnectionUri = 'http://' + $ExchangeFQDN + '/PowerShell/?SerializationLevel=Full' 
    Authentication = 'Kerberos'
    }
Import-PSSession (New-PSSession @Session) -AllowClobber

#Get Data
Write-Host "Getting data from Exchange..." -ForegroundColor Cyan
$AcceptedDomains = Get-AcceptedDomain
$InScopeRecipients = @(
    'DynamicDistributionGroup'
    'UserMailbox'
    'MailUniversalDistributionGroup'
    'MailUniversalSecurityGroup'
    'MailNonUniversalGroup'
    'PublicFolder'
    )
$AllRecipients = Get-Recipient -recipienttype $InScopeRecipients -ResultSize unlimited | select name, emailaddresses, RecipientType
$UniqueRecipientDomains = ($AllRecipients.emailaddresses | Where {$_ -like 'smtp*'}) -split '@' | where {$_ -NotLike 'smtp:*'} | select -Unique

Write-Host "Preparing Output 1 of 2..." -ForegroundColor Cyan
#Output address stats
$TextBlock = @(
    "Total Number of Recipients: " + $AllRecipients.Count
    "Number of Dynamic Distribution Groups: " +         ($AllRecipients | Where {$_.RecipientType -eq 'DynamicDistributionGroup'}).Count
    "Number of User Mailboxes: " + 	                    ($AllRecipients | Where {$_.RecipientType -eq 'UserMailbox'}).Count
    "Number of Mail-Universal Distribution Groups: " + 	($AllRecipients | Where {$_.RecipientType -eq 'MailUniversalDistributionGroup'}).Count
    "Number of Mail-UniversalSecurity Groups: " + 	    ($AllRecipients | Where {$_.RecipientType -eq 'MailUniversalSecurityGroup'}).Count
    "Number of Mail-NonUniversal Groups: " + 	        ($AllRecipients | Where {$_.RecipientType -eq 'MailNonUniversalGroup'}).Count
    "Number of Public Folders: " + 	                    ($AllRecipients | Where {$_.RecipientType -eq 'PublicFolder'}).Count
    ""
    "Number of Accepted Domains: " + $AcceptedDomains.count 
    ""
    "Number of domains found on recipients: " + $UniqueRecipientDomains.count 
    ""
    $DomainComparrison = Compare-Object $AcceptedDomains.DomainName $UniqueRecipientDomains
    "These domains have been assigned to recipients, but are not Accepted Domains in the Exchange Organization:"
    ($DomainComparrison | Where {$_.SideIndicator -eq '=>'}).InputObject 
    ""
    "These Accepted Domains are not assigned to any recipients:" 
    ($DomainComparrison | Where {$_.SideIndicator -eq '<='}).InputObject
    ""
    "See this CSV for a complete listing of all addresses: " + $CsvFile
    )

Write-Host "Preparing Output 2 of 2..." -ForegroundColor Cyan

$RecipientsAndSMTPProxies = @()
$CounterWatermark = 1
 
$AllRecipients | ForEach-Object {
    
    #Create a new placeholder object
    $RecipientOutputObject = New-Object PSObject -Property @{
        Name = $_.Name
        RecipientType = $_.RecipientType
        SMTPAddress0 =  ($_.emailaddresses | Where {$_ -clike 'SMTP:*'} ) -replace "SMTP:"
        }    
    
    #If applicable, get a list of other addresses for the recipient
    if (($_.emailaddresses).count -gt '1') {       
        $OtherAddresses = @()
        $OtherAddresses = ($_.emailaddresses | Where {$_ -clike 'smtp:*'} ) -replace "smtp:"
        
        $Counter = $OtherAddresses.count
        if ($Counter -gt $CounterWatermark) {$CounterWatermark = $Counter}
        $OtherAddresses | ForEach-Object {
            $RecipientOutputObject | Add-Member -MemberType NoteProperty -Name (“SmtpAddress” + $Counter) -Value ($_ -replace "smtp:")
            $Counter--
            }
        }
        $RecipientsAndSMTPProxies += $RecipientOutputObject
    }
  
 
$AttributeList = @(
    'Name'
    'RecipientType'
    )
$AttributeList += 0..$CounterWatermark | ForEach-Object {"SMTPAddress" + $_}


Write-Host "Saving report files to your desktop:" -ForegroundColor Green
Write-Host ""
Write-Host $TxtFile -ForegroundColor Green
Write-Host $CsvFile -ForegroundColor Green

$TextBlock | Out-File $TxtFile
$RecipientsAndSMTPProxies | Select $AttributeList | sort RecipientType, Name | Export-CSV $CsvFile -NoTypeInformation

Write-Host ""
Write-Host ""
Write-Host "Report Complete!" -ForegroundColor Green

# SIG # Begin signature block
# MIINHAYJKoZIhvcNAQcCoIINDTCCDQkCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUhL/zE0CHcg7njfTE1owXwCIa
# HnmgggpeMIIFJjCCBA6gAwIBAgIQDabkR8675p80ZdtFokcNRTANBgkqhkiG9w0B
# AQsFADByMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYD
# VQQLExB3d3cuZGlnaWNlcnQuY29tMTEwLwYDVQQDEyhEaWdpQ2VydCBTSEEyIEFz
# c3VyZWQgSUQgQ29kZSBTaWduaW5nIENBMB4XDTE2MTAxNzAwMDAwMFoXDTE3MTAy
# NTEyMDAwMFowYzELMAkGA1UEBhMCVVMxEzARBgNVBAgTCkNhbGlmb3JuaWExETAP
# BgNVBAcTCFBhY2lmaWNhMRUwEwYDVQQKEwxKZWZmIEd1aWxsZXQxFTATBgNVBAMT
# DEplZmYgR3VpbGxldDCCASIwDQYJKoZIhvcNAQEBBQADggEPADCCAQoCggEBAOxa
# 8mnNJehWpp0i/MMapjX2T5XXhZ+IdiW263HRXdtnjYTIXfWURyn+BjEb4VrnxHYC
# rXF9aktE9uzRSyHVt6gfz/Pt1slIT86umGW8zQBQR5f4etwfbBx3jPErKs8Qa6v4
# 0e8Cihpcv6Q3vVfOOzQgoGCsT+p7UBL5eRDfIa3KPcuD30DOcwSivwUOgKA18+ju
# yj0GjZdazLY0WKNVnDYpj1Aimjf44Ey1U0nWUocQj59AW27qRShf2z+bhY1EsY+y
# gxoKW30OP9kZg9gGSesArWRyoaxFQLmRX9T34/yVymr+70jGBQ9PlGun2Mu77Bdz
# i4KmiP3U30UYg8MVx7MCAwEAAaOCAcUwggHBMB8GA1UdIwQYMBaAFFrEuXsqCqOl
# 6nEDwGD5LfZldQ5YMB0GA1UdDgQWBBQqfYM6cJlPtyB43KxSogkKnl3yyjAOBgNV
# HQ8BAf8EBAMCB4AwEwYDVR0lBAwwCgYIKwYBBQUHAwMwdwYDVR0fBHAwbjA1oDOg
# MYYvaHR0cDovL2NybDMuZGlnaWNlcnQuY29tL3NoYTItYXNzdXJlZC1jcy1nMS5j
# cmwwNaAzoDGGL2h0dHA6Ly9jcmw0LmRpZ2ljZXJ0LmNvbS9zaGEyLWFzc3VyZWQt
# Y3MtZzEuY3JsMEwGA1UdIARFMEMwNwYJYIZIAYb9bAMBMCowKAYIKwYBBQUHAgEW
# HGh0dHBzOi8vd3d3LmRpZ2ljZXJ0LmNvbS9DUFMwCAYGZ4EMAQQBMIGEBggrBgEF
# BQcBAQR4MHYwJAYIKwYBBQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBO
# BggrBgEFBQcwAoZCaHR0cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0
# U0hBMkFzc3VyZWRJRENvZGVTaWduaW5nQ0EuY3J0MAwGA1UdEwEB/wQCMAAwDQYJ
# KoZIhvcNAQELBQADggEBAIzUknh+MUZLMkro4Kwez8KUdbEdwO7+dDCenjm4Ga7m
# VkiH2LrgPaowjDcuzU4EacAH9KHCG79k2+XEmHFWXA94EPP1LEx/Wuy7UoSy/6A/
# wFxnrHozOhRGzHsAwQpeeYWS2VpMH9/ZWDcMcLjCiU3W8Dc75PeXiAI7W9qdDcm9
# 1JUqiAcZ9IEvhtJEC/B4Aa9y8haXAbqIyxeConBsCOk3dtg4OKcinMGhSbxlordW
# byeAdKB46nso2+n12dUiWOKBRlhJLUduIqgH+tOuOEPZ72gAp7l2aF5dWA9TH/H2
# qSw2gN7CIN/SWxc18xqJzMxnEcXbZQoT3EJ1Ve3mR1gwggUwMIIEGKADAgECAhAE
# CRgbX9W7ZnVTQ7VvlVAIMA0GCSqGSIb3DQEBCwUAMGUxCzAJBgNVBAYTAlVTMRUw
# EwYDVQQKEwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20x
# JDAiBgNVBAMTG0RpZ2lDZXJ0IEFzc3VyZWQgSUQgUm9vdCBDQTAeFw0xMzEwMjIx
# MjAwMDBaFw0yODEwMjIxMjAwMDBaMHIxCzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxE
# aWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xMTAvBgNVBAMT
# KERpZ2lDZXJ0IFNIQTIgQXNzdXJlZCBJRCBDb2RlIFNpZ25pbmcgQ0EwggEiMA0G
# CSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQD407Mcfw4Rr2d3B9MLMUkZz9D7RZmx
# OttE9X/lqJ3bMtdx6nadBS63j/qSQ8Cl+YnUNxnXtqrwnIal2CWsDnkoOn7p0WfT
# xvspJ8fTeyOU5JEjlpB3gvmhhCNmElQzUHSxKCa7JGnCwlLyFGeKiUXULaGj6Ygs
# IJWuHEqHCN8M9eJNYBi+qsSyrnAxZjNxPqxwoqvOf+l8y5Kh5TsxHM/q8grkV7tK
# tel05iv+bMt+dDk2DZDv5LVOpKnqagqrhPOsZ061xPeM0SAlI+sIZD5SlsHyDxL0
# xY4PwaLoLFH3c7y9hbFig3NBggfkOItqcyDQD2RzPJ6fpjOp/RnfJZPRAgMBAAGj
# ggHNMIIByTASBgNVHRMBAf8ECDAGAQH/AgEAMA4GA1UdDwEB/wQEAwIBhjATBgNV
# HSUEDDAKBggrBgEFBQcDAzB5BggrBgEFBQcBAQRtMGswJAYIKwYBBQUHMAGGGGh0
# dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBDBggrBgEFBQcwAoY3aHR0cDovL2NhY2Vy
# dHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9vdENBLmNydDCBgQYD
# VR0fBHoweDA6oDigNoY0aHR0cDovL2NybDQuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0
# QXNzdXJlZElEUm9vdENBLmNybDA6oDigNoY0aHR0cDovL2NybDMuZGlnaWNlcnQu
# Y29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9vdENBLmNybDBPBgNVHSAESDBGMDgGCmCG
# SAGG/WwAAgQwKjAoBggrBgEFBQcCARYcaHR0cHM6Ly93d3cuZGlnaWNlcnQuY29t
# L0NQUzAKBghghkgBhv1sAzAdBgNVHQ4EFgQUWsS5eyoKo6XqcQPAYPkt9mV1Dlgw
# HwYDVR0jBBgwFoAUReuir/SSy4IxLVGLp6chnfNtyA8wDQYJKoZIhvcNAQELBQAD
# ggEBAD7sDVoks/Mi0RXILHwlKXaoHV0cLToaxO8wYdd+C2D9wz0PxK+L/e8q3yBV
# N7Dh9tGSdQ9RtG6ljlriXiSBThCk7j9xjmMOE0ut119EefM2FAaK95xGTlz/kLEb
# Bw6RFfu6r7VRwo0kriTGxycqoSkoGjpxKAI8LpGjwCUR4pwUR6F6aGivm6dcIFzZ
# cbEMj7uo+MUSaJ/PQMtARKUT8OZkDCUIQjKyNookAv4vcn4c10lFluhZHen6dGRr
# sutmQ9qzsIzV6Q3d9gEgzpkxYz0IGhizgZtPxpMQBvwHgfqL2vmCSfdibqFT+hKU
# GIUukpHqaGxEMrJmoecYpJpkUe8xggIoMIICJAIBATCBhjByMQswCQYDVQQGEwJV
# UzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYDVQQLExB3d3cuZGlnaWNlcnQu
# Y29tMTEwLwYDVQQDEyhEaWdpQ2VydCBTSEEyIEFzc3VyZWQgSUQgQ29kZSBTaWdu
# aW5nIENBAhANpuRHzrvmnzRl20WiRw1FMAkGBSsOAwIaBQCgeDAYBgorBgEEAYI3
# AgEMMQowCKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisG
# AQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMCMGCSqGSIb3DQEJBDEWBBTgVyiVMLHt
# Vvyj8OCklWAMgxqkazANBgkqhkiG9w0BAQEFAASCAQCnahVd3DPS8FQt/h6Ip3a+
# uIJgfjwpmymXZNWecdLq8Cx6ZO1x/VN13u6I8coookLMSIHWiDdnHQn49MrPXMMI
# K5rI02Wx98F6L7v355dyIo/M5XkxMrZge24uI7pMI2hRSfs08V0jgu9qGB46bPMq
# qm3tJkF4H5BKcDCPTJyuRtHa5ukJzIY8zHWZ3nknCl904H+uVh8IPB7fjAhHKFXO
# nlM/gn2ItnYPYb43nPsKLrhKp82pgH/IniZy1BON0EfPyhAnTls6+FQze4G61VvY
# MEkvXfRrJRaiaqQON8/dVDy7WQpi0D5CxHXxaomuL83avsPS1tk9502Vz8voOdic
# SIG # End signature block
