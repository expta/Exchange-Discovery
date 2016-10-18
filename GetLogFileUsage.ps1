#---------------------------------------------------------------------------------------------------------------
#
# Script name:	GetLogFileUsage.ps1
# Version:		2.1
# Author:		Karsten Palmvig
#
# Purpose:		Provide Exchange Server or third party mail server log file statistics for the following:
#
#				Exchange 2010/2013 Server Role Requirements Calculator by Ross Smith IV
#				Exchange Client Network Bandwidth Calculator by Neil Johnson
#
#---------------------------------------------------------------------------------------------------------------
param (
	[Switch]$Help,																							# Display command line help
	[Switch]$File,																							# Use file as input
	$Date = $null,																							# Specify date for log files (default is last 24 hours)
	$Server = $null,																						# Specific server name or "all"
	$Database = $null,																						# Specific database log set (overruled if server is specified)
	$PathFile = ".\paths.txt",																				# Use this file as input if no server or DB is specified, for legacy and non-Exchange support
	$Delimiter = ";"																						# Set whatever separator your Excel defaults to, to be able to open .csv output file directly, also used for pathfile import
)

if ($help -or ($psboundparameters.count -eq 0)) { 
	Write-Output ""
	Write-Output "Use the script with the following command line options:"
	Write-Output ""
	Write-Output ".\GetLogFileUsage.ps1 [-Date DD-MM-YYYY] [-Server <server name> | all] [-Database <db name>] [-File] [-PathFile <file name>] [-Delimiter <char>] [-Help]"
	Write-Output ""
	Write-Output "Defaults are: Read last 24 hours of logfiles, PathFile = .\paths.txt, use semicolon as delimiter."
	Write-Output ""
	Write-Output "Format of the PathFile which is used for legacy or non-Exchange server support is:"
	Write-Output ""
	Write-Output "Server;LogFolderPath"
	Write-Output "<server>;<Path1>"
	Write-Output "<server>;<Path2>"
	Write-Output ""
	Exit 
}

Write-Output "Retrieving LogFolderPaths..."
if ($server -ne $null) {
	if ($server -like "all") { $TransactionLogs = Get-MailboxDatabase | Select-Object Server,LogFolderPath }# Grab LogFolderPath from all Exchange Servers
	else { $TransactionLogs = Get-MailboxDatabase -Server $server | Select-Object Server,LogFolderPath }}  	# Assuming qualified Server name 
elseif ($database -ne $null) { $TransactionLogs = Get-MailboxDatabase $database | Select-Object Server,LogFolderPath } # Assuming qualified DB name
elseif ($file) { $TransactionLogs = Import-Csv $pathfile -Delimiter $Delimiter }									# Use file
else { Write-Output "No usable command line arguments..."; Exit }

if ($date -eq $null) { $start = Get-Date([DateTime]::Now).AddDays(-1); $end = Get-Date }					# Last 24 hours (default)
else { $start = Get-Date($date); $end = $(Get-Date($date)).AddDays(1) }										# Custom date

$outputfile = $pwd.Path + "\output-$((Get-Date -uformat %Y%m%d%H%M).ToString()).csv"						# Log all findings in this file in current folder

$hash = @{
	Time		= $Hour
	Count 		= $Count
	Percent		= $Percent }
$listObj = New-Object PSObject -Property $hash
$Collection = @()

function CountFiles {
	$stream = [System.IO.StreamWriter] $outputfile
	$header = "Server" + $Delimiter + "LogFolderPath"
	foreach ($hour in 0..23) { $header += $Delimiter + "Hour-" + $hour.tostring("00") }
	$stream.WriteLine($header)
	$countCollection = @()
	$firstrun = $true
	
	foreach ($logPath in $TransactionLogs) {
		if ($logpath.server.Name -ne $null) { $logpath.server = $logpath.server.name; $logpath.LogFolderPath = $logpath.LogFolderPath.PathName } # When multivalued entries, take out the middle man to support both file/get-mailboxdatabase
		$output = [string]$logPath.Server + $Delimiter + [string]$logPath.LogFolderPath
		$tmpCollection = @()
		$remotePath = "\\" + $logPath.Server + "\" + $logPath.LogFolderPath.Replace(":","$")
		$files = Get-ChildItem -LiteralPath $remotePath -filter "*.log" | Where-Object { $_.LastWriteTime -ge $start -and $_.LastWriteTime -le $end }
		foreach ($hour in 0..23) {
			$selectFiles = $files | Where-Object { $_.LastWriteTime.hour -eq $hour }
			$output += $Delimiter + [string]$selectFiles.count
			$listObj = New-Object PSObject -Property $hash
			$listObj.Time = $($hour + 1).tostring("00") + ":00"
			$listObj.Count = $selectFiles.count
			$listObj.Percent = 0
			$tmpCollection += $listObj
		}
		$stream.WriteLine($output)
		if ($firstrun) { $countCollection = $tmpCollection } 												# First run, just copy the array 
		else { for ($i=0; $i -le 23; $i++) { $countCollection[$i].count += $tmpCollection[$i].count } } 	# Subsequent runs, add counts
		$firstrun = $false
	}
	$stream.close()
	Return $countCollection
}

function CalcPct($Collection) {
	$sum = 0
	foreach ($element in $Collection) {	$sum += $element.count }											# Sum of log files
	foreach ($element in $Collection) { if ($element.Count -gt 0) { $element.Percent = $("{0:N2}" -f (($element.Count / $sum) * 100)) } } # Calculate and insert percentage
	Return $Collection
}

Write-Output "Collecting data..."
$Collection = CountFiles

Write-Output "Calculating..."
$Collection = CalcPct($Collection)

$Collection | ft Time,Percent,Count -AutoSize
$Collection | ft Percent -AutoSize -HideTableHeaders | Clip

Write-Output "Collected data is available in: $outputfile"
Write-Output "Percent column has been added to your clipboard for easier pasting..."
# SIG # Begin signature block
# MIINHAYJKoZIhvcNAQcCoIINDTCCDQkCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUd9LLxZuj79vru4wF0fofmanR
# w96gggpeMIIFJjCCBA6gAwIBAgIQDabkR8675p80ZdtFokcNRTANBgkqhkiG9w0B
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
# AQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMCMGCSqGSIb3DQEJBDEWBBQ9u5fEu9RQ
# IDUZlEptTch8AowmbzANBgkqhkiG9w0BAQEFAASCAQCS9PeA1vja6pDJssQu35gq
# Wtzeq8ThY9nKVj149G+9EUykFY3xmS7gmekg8DQar1hzO1IUORDp0Sa2um5hFOeQ
# W5bFmACxaZmdZwGZ5tgNFiztI8OvzzqzxmnuAD7fAax39gtmmSt0a7KLh4Q/cmPx
# U1y9XIUh464aBW5RSSHxzDzgYqtMWM5MKdgmwwHSHFTa6H6c6MfSRTZeuHH8fmfF
# 5j2kAKNiyW6x4fnGNYCXdN3ZvYtHYT93ZiSJZ6g0fO9xbhZAYHq17QwNPACVI0WX
# FLJhJF2tGxQN7iUeZqGYrTsso4JVMk3O2XrmKjgfbJrWLfF/KYj590NCK0S7r0fD
# SIG # End signature block
