<#
.SYNOPSIS
CertificateReport.ps1 - Exchange Server SSL Certificate Report Script

.DESCRIPTION 
Generates a report of the SSL certificates installed on Exchange Server 2010 servers

.OUTPUTS
Outputs to a HTML file.

.EXAMPLE
.\CertificateReport.ps1
Reports SSL certificates for Exchange Server servers and outputs to a HTML file.

.LINK
http://exchangeserverpro.com/powershell-script-ssl-certificate-report

.NOTES
Written By: Paul Cunningham
Website:	http://exchangeserverpro.com
Twitter:	http://twitter.com/exchservpro

Change Log
V1.00, 13/03/2014 - Initial Version
V1.01, 13/03/2014 - Minor bug fix

#>

$myDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$reportfile = "$myDir\CertificateReport.html"

$htmlreport = @()

$exchangeservers = @(Get-ExchangeServer | where {$_.ServerRole -notlike "Edge"})

foreach ($server in $exchangeservers)
{
    $htmlsegment = @()
    
    $serverdetails = "Server: $($server.Name) ($($server.ServerRole)) [$($server.AdminDisplayVersion)]"
    Write-Host $serverdetails
    
    $certificates = @(Get-ExchangeCertificate -Server $server -ErrorAction SilentlyContinue)

    $certtable = @()

    foreach ($cert in $certificates)
    {
        
        $iis = $null
        $smtp = $null
        $pop = $null
        $imap = $null
        $um = $null       
        
        $subject = ((($cert.Subject -split ",")[0]) -split "=")[1]
                
        if ($($cert.IsSelfSigned))
        {
            $selfsigned = "Yes"
        }
        else
        {
            $selfsigned = "No"
        }

        $issuer = ((($cert.Issuer -split ",")[0]) -split "=")[1]

        #$domains = @($cert | Select -ExpandProperty:CertificateDomains)
        $certdomains = @($cert | Select -ExpandProperty:CertificateDomains)
        if ($($certdomains.Count) -gt 1)
        {
            $domains = $null
            $domains = $certdomains -join ", "
        }
        else
        {
            $domains = $certdomains[0]
        }

        #$services = @($cert | Select -ExpandProperty:Services)
        $services = $cert.ServicesStringForm.ToCharArray()

        if ($services -icontains "W") {$iis = "Yes"}
        if ($services -icontains "S") {$smtp = "Yes"}
        if ($services -icontains "P") {$pop = "Yes"}
        if ($services -icontains "I") {$imap = "Yes"}
        if ($services -icontains "U") {$um = "Yes"}

        $certObj = New-Object PSObject
        $certObj | Add-Member NoteProperty -Name "Subject" -Value $subject
        $certObj | Add-Member NoteProperty -Name "Status" -Value $cert.Status
        $certObj | Add-Member NoteProperty -Name "Expires" -Value $cert.NotAfter.ToShortDateString()
        $certObj | Add-Member NoteProperty -Name "Self Signed" -Value $selfsigned
        $certObj | Add-Member NoteProperty -Name "Issuer" -Value $issuer
        $certObj | Add-Member NoteProperty -Name "SMTP" -Value $smtp
        $certObj | Add-Member NoteProperty -Name "IIS" -Value $iis
        $certObj | Add-Member NoteProperty -Name "POP" -Value $pop
        $certObj | Add-Member NoteProperty -Name "IMAP" -Value $imap
        $certObj | Add-Member NoteProperty -Name "UM" -Value $um
        $certObj | Add-Member NoteProperty -Name "Thumbprint" -Value $cert.Thumbprint
        $certObj | Add-Member NoteProperty -Name "Domains" -Value $domains
        
        $certtable += $certObj
    }

    $htmlcerttable = $certtable | ConvertTo-Html -Fragment

    $htmlserver = "<p>$serverdetails</p>" + $htmlcerttable

    $htmlreport += $htmlserver
}


$htmlhead="<html>
			<style>
			BODY{font-family: Arial; font-size: 10pt;}
			H1{font-size: 16px;}
			H2{font-size: 14px;}
			H3{font-size: 12px;}
			TABLE{border: 1px solid black; border-collapse: collapse; font-size: 10pt;}
			TH{border: 1px solid black; background: #dddddd; padding: 5px; color: #000000;}
			TD{border: 1px solid black; padding: 5px; }
			td.pass{background: #7FFF00;}
			td.warn{background: #FFE600;}
			td.fail{background: #FF0000; color: #ffffff;}
			td.info{background: #85D4FF;}
			</style>
			<body>
			<h3 align=""center"">Exchange Server Certificate Report</h3>"

$htmltail = "</body>
			</html>"

$htmlreport = $htmlhead + $htmlreport + $htmltail

$htmlreport | Out-File -Encoding utf8 -FilePath $reportfile

# SIG # Begin signature block
# MIINHAYJKoZIhvcNAQcCoIINDTCCDQkCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQU6TdBVOURF2ha0V5mchm+SRvV
# ZTqgggpeMIIFJjCCBA6gAwIBAgIQDabkR8675p80ZdtFokcNRTANBgkqhkiG9w0B
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
# AQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMCMGCSqGSIb3DQEJBDEWBBQ/m9Gq6pmn
# hNRg5/PVvJh2ann9xDANBgkqhkiG9w0BAQEFAASCAQCOi2RjwBZ5F8kEg8xEy98o
# q5HnwZoKfXPP4UaPbEj2QsDPefOSPWvUa6bLAhJM3S2b7vp5Sqo9cUNrlCfrq4JJ
# isg2GSuna5/oi1EDZAAtjEBH1aAJamjTcvEmE6Z8lookkYT4FguGBa/6osHpRXyP
# n+sO1r4SYynujAVr6N7lZHp7DMJmoeQtE5TTHDMA0izNfWluF5EpuzI3p3ERicV+
# LJKWLyduUYitzxzNiIesbxNq8aEuNbTBFkTeQxzexpwnKDT+9tAIGLs3sJ+WKFKA
# UvbsRhO9V9OQDpiuoVHU/vfqJUafjK3bcomSeaU+bOJkum5tK1cl1Y35SDUi9anR
# SIG # End signature block
