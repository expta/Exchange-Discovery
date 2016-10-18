<#
.SYNOPSIS
Get-ADInfo.ps1 - PowerShell script to collect Active Directory information

.DESCRIPTION 
This PowerShell Script collects some basic information about an Active Directory
environment that is useful for verifying the pre-requisites for an Exchange
Server deployment or upgrade.

.OUTPUTS
Results are output to the console and to a HTML file.

.EXAMPLE
.\Get-ADInfo.ps1
Runs the script and generates the output.

.NOTES
Written by: Paul Cunningham

Find me on:

* My Blog:	http://paulcunningham.me
* Twitter:	https://twitter.com/paulcunningham
* LinkedIn:	http://au.linkedin.com/in/cunninghamp/
* Github:	https://github.com/cunninghamp

For more Exchange Server tips, tricks and news
check out Exchange Server Pro.

* Website:	http://exchangeserverpro.com
* Twitter:	http://twitter.com/exchservpro

Change Log
V1.00, 26/06/2015 - Initial version
#>

#requires -version 2

$htmlreport = $null
$htmlbody = $null
$htmlfile = "ADInfo.html"
$spacer = "<br />"

Import-Module ActiveDirectory

#---------------------------------------------------------------------
# Collect AD Forest information and convert to HTML fragment
#---------------------------------------------------------------------

$forest = Get-ADForest

$htmlbody += "<h3>Forest Details</h3>"

$forestinfo = New-Object PSObject

Write-Host -ForeGroundColor Yellow "*** Forest: $($forest.RootDomain) ***"
Write-Host ""
$forestinfo  | Add-Member NoteProperty -Name "Forest" -Value $($forest.RootDomain)

Write-Host "Forest Mode: $($forest.ForestMode)"
$forestinfo  | Add-Member NoteProperty -Name "Forest Mode" -Value $($forest.ForestMode)

Write-Host "Schema Master: $($forest.SchemaMaster)"
$forestinfo  | Add-Member NoteProperty -Name "Schema Master" -Value $($forest.SchemaMaster)

Write-Host "Domain Naming Master: $($forest.DomainNamingMaster)"
$forestinfo  | Add-Member NoteProperty -Name "Domain Naming Master" -Value $($forest.DomainNamingMaster)

Write-Host "Additional UPN Suffixes: $($forest.UPNSuffixes)"
$forestinfo  | Add-Member NoteProperty -Name "Additional UPN Suffixes" -Value $($forest.UPNSuffixes)

$htmlbody += $forestinfo | ConvertTo-Html -Fragment
$htmlbody += $spacer

#---------------------------------------------------------------------
# Collect AD Domain information and convert to HTML fragment
#---------------------------------------------------------------------

$htmlbody += "<h3>Domain Details</h3>"

$domains = @($forest | Select -ExpandProperty:Domains)

Foreach ($domain in $domains)
{
    Write-Host ""
    Write-Host -ForeGroundColor Yellow "*** Domain: $domain ***"
    Write-Host ""

    $domaindetails = Get-ADDomain $domain

    $domaininfo = New-Object PSObject
    
    $domaininfo | Add-Member NoteProperty -Name "Name" -Value $domaindetails.Name
    
    Write-Host "NetBIOS Name: $($domaindetails.NetBIOSName)"
    $domaininfo | Add-Member NoteProperty -Name "NetBIOS Name" -Value $domaindetails.NetBIOSName

    Write-Host "Domain Mode: $($domaindetails.DomainMode)"
    $domaininfo | Add-Member NoteProperty -Name "Mode" -Value $($domaindetails.DomainMode)
    
    Write-Host "PDC Emulator: $($domaindetails.PDCEmulator)"
    $domaininfo | Add-Member NoteProperty -Name "PDC Emulator" -Value $($domaindetails.PDCEmulator)
    
    Write-Host "Infrastructure Master: $($domaindetails.InfrastructureMaster)"
    $domaininfo | Add-Member NoteProperty -Name "Infrastructure Master" -Value $($domaindetails.InfrastructureMaster)   
    
    Write-Host "RID Master: $($domaindetails.RIDMaster)"
    $domaininfo | Add-Member NoteProperty -Name "RID Master" -Value $($domaindetails.RIDMaster)

    $htmlbody += $domaininfo | ConvertTo-Html -Fragment
    $htmlbody += $spacer

}

$htmlbody += "<h3>Global Catalog Servers by Site/OS</h3>"

$domaincontrollers = @(Get-ADDomainController -Filter {IsGlobalCatalog -eq $true})

$gcs = @($domaincontrollers | Group-Object -Property:Site,OperatingSystem | Select @{Expression="Name";Label="Site, OS"},Count)

Write-Host ""
Write-Host -ForeGroundColor Yellow "*** Global Catalogs by Site/OS ***"

$gcs = $gcs | Sort "Site, OS"
$gcs | ft -auto

$htmlbody += $gcs | ConvertTo-Html -Fragment


Write-Verbose "Producing HTML report"
    
$reportime = Get-Date

#Common HTML head and styles
$htmlhead="<html>
			<style>
			BODY{font-family: Arial; font-size: 8pt;}
			H1{font-size: 20px;}
			H2{font-size: 18px;}
			H3{font-size: 16px;}
			TABLE{border: 1px solid black; border-collapse: collapse; font-size: 8pt;}
			TH{border: 1px solid black; background: #dddddd; padding: 5px; color: #000000;}
			TD{border: 1px solid black; padding: 5px; }
			td.pass{background: #7FFF00;}
			td.warn{background: #FFE600;}
			td.fail{background: #FF0000; color: #ffffff;}
			td.info{background: #85D4FF;}
			</style>
			<body>
			<h1 align=""center"">Active Directory Information</h1>
			<h3 align=""center"">Generated: $reportime</h3>"

$htmltail = "</body>
		</html>"

$htmlreport = $htmlhead + $htmlbody + $htmltail

$htmlreport | Out-File $htmlfile -Encoding Utf8
# SIG # Begin signature block
# MIINHAYJKoZIhvcNAQcCoIINDTCCDQkCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUk/dM3A9ouLtV/J85kmY8xM4T
# wkWgggpeMIIFJjCCBA6gAwIBAgIQDabkR8675p80ZdtFokcNRTANBgkqhkiG9w0B
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
# AQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMCMGCSqGSIb3DQEJBDEWBBTYfPb0CtXD
# lO/9j0ejthFx0gxzgDANBgkqhkiG9w0BAQEFAASCAQASUnUkJh4AvF1LbQ1xODmG
# KQUn0fcv92vN59JuG3iCT2UvIAtQwPjlSdnRDsN/EW/BgD+5fqcBdwI1gOj3/K+k
# eHoVd5F2x7/yMou8eeOgthxzqD1Ok5qj5aM1urTzJ6DRYrK7AVcumHr//WNvQfEB
# vPkUCAGi8cizxc3wvt0tg7TjU48kxgnfCAjGfJCzMoj2DFaQmaJEGckrOXiOHLEa
# PubXhXCRMAReylVIxoyv6G85PQ9MYnCkPHGlLfxBwajl5SaVFJxO6RRi/sPEB8mM
# Wjw7s9SD4DuCAGWPjtAlCXM2kYDaMyPHLUwtLxhcxno8fIFxJjVVAOU3ejO8uI4A
# SIG # End signature block
