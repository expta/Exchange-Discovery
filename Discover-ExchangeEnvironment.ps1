#Discover-ExchangeEnvironment.ps1
#Jeff Guillet, MCSM | MVP
#Gathers Exchange environment information
#10/03/2016

if (!(Get-PSSession).ConfigurationName -eq "Microsoft.Exchange") {
	Write-Host "Discovery must be run from the Exchange Management Shell on the highest version Exchange Server in the organization." -ForegroundColor Red
	Exit
}

$Now = Get-Date
Write-Host "Started discovery: $Now from server $env:computername.$env:userdnsdomain" -ForegroundColor Green
Write-Host

$Org = (Get-OrganizationConfig).Name

Write-Host "Gathering yesterday's message stats..." -ForegroundColor Cyan
# Note: Change 1 to the number of days to offset. For example, if today is Monday use "3" to use the logs stats from three days ago (Friday) instead of Sunday.
# See http://bit.ly/1WFf3yC for how to use this output.
.\MessageStats.ps1 1

Write-Host "Running virtual directory report..." -ForegroundColor Cyan
. .\Get-VirDirInfo.ps1
Start-Sleep -s 2
Get-VirDirInfo -ADPropertiesOnly -Filepath .

Write-Host "Creating Exchange environment report..." -ForegroundColor Cyan
.\Get-ExchangeEnvironmentReport.ps1 -HtmlReport $Org-Environment.htm

Write-Host "Creating Exchange backup report..." -ForegroundColor Cyan
.\Get-DailyBackupAlerts.ps1

Write-Host "Getting Active Directory info..." -ForegroundColor Cyan
.\Get-ADInfo.ps1

Write-Host "Getting quota info..." -ForegroundColor Cyan
.\Get-Quotas.ps1

Write-Host "Getting Mailbox info..." -ForegroundColor Cyan
Get-Mailbox -Resultsize Unlimited | ft UserPrincipalName,Name,WindowsEmailAddress,emailaddresses -Wrap | Out-File -FilePath $Org-EmailAddresses.txt

Write-Host "Getting Group info..." -ForegroundColor Cyan
Get-DistributionGroup -Resultsize Unlimited | ft Name,DisplayName,WindowsEmailAddress,emailaddresses -Wrap | Out-File -FilePath $Org-GroupEmailAddresses.txt

Write-Host "Getting Contact info..." -ForegroundColor Cyan
Get-Contact -Resultsize Unlimited | ft name,windowsemailaddress -Wrap | Out-File -FilePath $Org-ContactEmailAddresses.txt

Write-Host "Calculating average mailbox size...." -ForegroundColor Cyan
Get-Mailbox -Resultsize Unlimited | Get-MailboxStatistics | %{$_.TotalItemSize.Value.ToMB()} | Measure-Object -Average | Out-File -FilePath $Org-AvgMailboxSizeInMB.txt

Write-Host "Gathering log file usage percentages..." -ForegroundColor Cyan
.\GetLogFileUsage.ps1 -Server all | Out-File -FilePath $Org-LogFilePercentages.txt

Write-Host "Running Exchange Server health report..." -ForegroundColor Cyan
.\Test-ExchangeServerHealth.ps1 -ReportMode -ReportFile $org-ExchangeServerHealth.htm

Write-Host "Running Exchange client reports..." -ForegroundColor Cyan
Get-ClientAccessServer -WA SilentlyContinue | % {.\Get-OutlookClients.ps1 $_.Name}
Get-Content *Clients.csv | Out-File "All Exchange Clients.csv"

#Write-Host "Gathering mailbox report. This can take quite a while to run..." -ForegroundColor Cyan
#.\Get-MailboxReport.ps1 -all

$Now = Get-Date
Write-Host "Ended discovery: $Now from server $env:computername.$env:userdnsdomain" -ForegroundColor Green
Write-Host

Write-Host "Compressing discovery files to $Org-DiscoveryFiles.zip..." -Foregroundcolor Cyan

.\Get-BufferHtml.ps1 > ScreenOutput.htm

function Add-Zip
{
	param([string]$zipfilename)

	if(-not (Test-Path($zipfilename)))
	{
		Set-Content $zipfilename ("PK" + [char]5 + [char]6 + ("$([char]0)" * 18))
		(dir $zipfilename).IsReadOnly = $false  
		Start-Sleep -Seconds 2
	}

	$zipfilename = Resolve-Path $zipfilename
	$shellApplication = New-Object -ComObject shell.application
	$zipPackage = $shellApplication.NameSpace($zipfilename)

	foreach($file in $input) 
	{ 
		$zipPackage.MoveHere($file.FullName)
		Start-Sleep -Seconds 1
	}
}

dir *.csv | Add-Zip $Org-DiscoveryFiles.zip
dir *.txt | Add-Zip $Org-DiscoveryFiles.zip
dir *.ht* | Add-Zip $Org-DiscoveryFiles.zip

Write-Host
Write-Host "Done! Please send the ""$Org-DiscoveryFiles.zip"" file to your SPS engineer. Thank you!" -ForegroundColor White

# SIG # Begin signature block
# MIINHAYJKoZIhvcNAQcCoIINDTCCDQkCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUUgOj4zRLqB4r/joN26MZyAYE
# 88qgggpeMIIFJjCCBA6gAwIBAgIQDabkR8675p80ZdtFokcNRTANBgkqhkiG9w0B
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
# AQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMCMGCSqGSIb3DQEJBDEWBBRbXkuS0SM1
# RI9M/A5kt5bnuE6wRzANBgkqhkiG9w0BAQEFAASCAQAkZpLoUdM9RS9Kk1/MGRhc
# EzvTFNXsdIAYhxZSNMd/A1+BFnm/yRTfHmIiCXFJtgf5qgu2JpfZe+A5d/25Jcqm
# Q5vN4sJwdJAJ3pXUZbztuiIt33REUmGXZRIduuxmKRBMLxLJqZaJrAtNVZke6lfa
# msg63WpoyiYckBKuvpMIVVa7LxuncqHSNRrHKRxKz1x16Yc72Dkv5O7TAKgwh+VI
# oUv6amN+FWVvDSkQ1HRdo7w+jbTz8U+EdrTWkjtvmioccUAxZFYGnH/K8DXE6Rta
# neA42ag7e6y3VrAg6OWRkD9L1ZWJXyMtUqkw1T36PgfxFRrSUMY3vuTFzRBfxae1
# SIG # End signature block
