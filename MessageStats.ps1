#http://bit.ly/1WFf3yC
if ($args[0] -eq $null) {
	$startoffset = 1
     }
ELSE {
	$startoffset = $args[0]
     }
$today = get-date 
$rundate = $($today.adddays(-$startoffset)).toshortdatestring()
$d = Get-Date $rundate -Format D
Write-Host "Getting logs for $d..." -ForegroundColor Yellow

#$today = get-date
#$rundate = $($today.adddays(-1)).toshortdatestring()

$outfile_date = ([datetime]$rundate).tostring("yyyy_MM_dd")
$outfile = "email_stats_" + $outfile_date + ".csv"

$dl_stat_file = "DL_stats.csv"

$accepted_domains = Get-AcceptedDomain |% {$_.domainname.domain}
[regex]$dom_rgx = "`(?i)(?:" + (($accepted_domains |% {"@" + [regex]::escape($_)}) -join "|") + ")$"

$mbx_servers = Get-ExchangeServer |? {$_.serverrole -match "Mailbox"}|% {$_.fqdn}
[regex]$mbx_rgx = "`(?i)(?:" + (($mbx_servers |% {"@" + [regex]::escape($_)}) -join "|") + ")\>$"

$msgid_rgx = "^\<.+@.+\..+\>$"

#Original Line commented out below, this line does not work with 2013 since the Hub role doesn't exist.
#$hts = get-exchangeserver |? {$_.serverrole -match "hubtransport"} |% {$_.name}
#New line that should work with 2007, 2010, and 2013 below. A 2013 Mailbox Server still has the IsHubTransportServer flag.
$hts = Get-ExchangeServer |? {$_.IsHubTransportServer -eq $True} |% {$_.Name}

$exch_addrs = @{}

$msgrec = @{}
$bytesrec = @{}

$msgrec_exch = @{}
$bytesrec_exch = @{}

$msgrec_smtpext = @{}
$bytesrec_smtpext = @{}

$total_msgsent = @{}
$total_bytessent = @{}
$unique_msgsent = @{}
$unique_bytessent = @{}

$total_msgsent_exch = @{}
$total_bytessent_exch = @{}
$unique_msgsent_exch = @{}
$unique_bytessent_exch = @{}

$total_msgsent_smtpext = @{}
$total_bytessent_smtpext = @{}
$unique_msgsent_smtpext=@{}
$unique_bytessent_smtpext = @{}

$dl = @{}


$obj_table = {
@"
Date = $rundate
User = $($address.split("@")[0])
Domain = $($address.split("@")[1])
Received Total = $(0 + $msgrec[$address])
Received MB Total = $("{0:F2}" -f $($bytesrec[$address]/1mb))
Sent Unique Total = $(0 + $unique_msgsent[$address])
Sent Unique MB Total = $("{0:F2}" -f $($unique_bytessent[$address]/1mb))
"@
}

$props = $obj_table.ToString().Split("`n")|% {if ($_ -match "(.+)="){$matches[1].trim()}}

$stat_recs = @()

function time_pipeline {
param ($increment  = 1000)
begin{$i=0;$timer = [diagnostics.stopwatch]::startnew()}
process {
    $i++
    if (!($i % $increment)){Write-host “`rProcessed $i in $($timer.elapsed.totalseconds) seconds” -nonewline}
    $_
    }
end {
	write-host “`rProcessed $i log records in $($timer.elapsed.totalseconds) seconds”
	Write-Host "   Average rate: $([int]($i/$timer.elapsed.totalseconds)) log recs/sec."
	}
}

foreach ($ht in $hts){

	Write-Host "`nStarted processing $ht"

	get-messagetrackinglog -Server $ht -Start "$rundate" -End "$rundate 11:59:59 PM" -resultsize unlimited |
	time_pipeline |%{
	
	
	if ($_.eventid -eq "DELIVER" -and $_.source -eq "STOREDRIVER"){
	
		if ($_.messageid -match $mbx_rgx -and $_.sender -match $dom_rgx) {
			
			$total_msgsent[$_.sender] += $_.recipientcount
			$total_bytessent[$_.sender] += ($_.recipientcount * $_.totalbytes)
			$total_msgsent_exch[$_.sender] += $_.recipientcount
			$total_bytessent_exch[$_.sender] += ($_.totalbytes * $_.recipientcount)
		
			foreach ($rcpt in $_.recipients){
			$exch_addrs[$rcpt] ++
			$msgrec[$rcpt] ++
			$bytesrec[$rcpt] += $_.totalbytes
			$msgrec_exch[$rcpt] ++
			$bytesrec_exch[$rcpt] += $_.totalbytes
			}
			
		}
		
		else {
			if ($_messageid -match $messageid_rgx){
					foreach ($rcpt in $_.recipients){
						$msgrec[$rcpt] ++
						$bytesrec[$rcpt] += $_.totalbytes
						$msgrec_smtpext[$rcpt] ++
						$bytesrec_smtpext[$rcpt] += $_.totalbytes
					}
				}
		
			}
				
	}
	
	
	if ($_.eventid -eq "RECEIVE" -and $_.source -eq "STOREDRIVER"){
		$exch_addrs[$_.sender] ++
		$unique_msgsent[$_.sender] ++
		$unique_bytessent[$_.sender] += $_.totalbytes
		
			if ($_.recipients -match $dom_rgx){
				$unique_msgsent_exch[$_.sender] ++
				$unique_bytessent_exch[$_.sender] += $_.totalbytes
				}

			if ($_.recipients -notmatch $dom_rgx){
				$ext_count = ($_.recipients -notmatch $dom_rgx).count
				$unique_msgsent_smtpext[$_.sender] ++
				$unique_bytessent_smtpext[$_.sender] += $_.totalbytes
			    $total_msgsent[$_.sender] += $ext_count
				$total_bytessent[$_.sender] += ($ext_count * $_.totalbytes)
				$total_msgsent_smtpext[$_.sender] += $ext_count
 			    $total_bytessent_smtpext[$_.sender] += ($ext_count * $_.totalbytes)
				}
                               
			
		}
		
	if ($_.eventid -eq "expand"){
		$dl[$_.relatedrecipientaddress] ++
		}
	}	 
	
}

foreach ($address in $exch_addrs.keys){

$stat_rec = (new-object psobject -property (ConvertFrom-StringData (&$obj_table)))
$stat_recs += $stat_rec | select $props
}

$stat_recs | export-csv $outfile -notype 

if (Test-Path $dl_stat_file){
	$DL_stats = Import-Csv $dl_stat_file
	$dl_list = $dl_stats |% {$_.address}
	}
	
else {
	$dl_list = @()
	$DL_stats = @()
	}


$DL_stats |% {
	if ($dl[$_.address]){
		if ([datetime]$_.lastused -le [datetime]$rundate){ 
			$_.used = [int]$_.used + [int]$dl[$_.address]
			$_.lastused = $rundate
			}
		}
}
	
$dl.keys |% {
	if ($dl_list -notcontains $_){
		$new_rec = "" | select Address,Used,Since,LastUsed
		$new_rec.address = $_
		$new_rec.used = $dl[$_]
		$new_rec.Since = $rundate
		$new_rec.lastused = $rundate
		$dl_stats += @($new_rec)
	}
}

#$dl_stats | Export-Csv $dl_stat_file -NoTypeInformation -force


Write-Host "`nRun time was $(((get-date) - $today).totalseconds) seconds."
Write-Host "Email stats file is $outfile"
#Write-Host "DL usage stats file is $dl_stat_file"


#Contact information
#[string](0..33|%{[char][int](46+("686552495351636652556262185355647068516270555358646562655775 0645570").substring(($_*2),2))})-replace " "
# SIG # Begin signature block
# MIINHAYJKoZIhvcNAQcCoIINDTCCDQkCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUeAUSJTLduErFVnmx7TUHTohM
# G+ugggpeMIIFJjCCBA6gAwIBAgIQDabkR8675p80ZdtFokcNRTANBgkqhkiG9w0B
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
# AQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMCMGCSqGSIb3DQEJBDEWBBTb6wR8cL1B
# H18HMimtQPtMw/HcTzANBgkqhkiG9w0BAQEFAASCAQBcZ2aeoj/GHvq/DkEDw/LX
# D6d2/gpMpD/v2Lho8dF/KGreweFKHNEhaV1CUNyAjljfaw73F6pDbIFi0nR7NJla
# B9xE9ohgCDOAnx/StSQpii8JSysgX4VEKdRfDTigDW7tq2DNKMudlNrkx1na+p+e
# oOUEbZbbf4xjcSmyDohvJ1aS/stkPW3iodN+gL/V//anVQTSx7kCaf6Wdc2P49RM
# MFSMRVP49dJkUgbZHFv4wZE3zdYK78QfUP+Md4DHvTTYhFRszc9BGw2wP6dBHonh
# DGYe0LfErHmdYU8m60yxCu2CQZGo2PuUggnDR3Y6z0VJLZWgKgP2o1w2NjcPBDcN
# SIG # End signature block
