#------------------------------------------------------------------------------
# Copyright 2006-2007 Adrian Milliner (ps1 at soapyfrog dot com)
# http://ps1.soapyfrog.com
#
# This work is licenced under the Creative Commons 
# Attribution-NonCommercial-ShareAlike 2.5 License. 
# To view a copy of this licence, visit 
# http://creativecommons.org/licenses/by-nc-sa/2.5/ 
# or send a letter to 
# Creative Commons, 559 Nathan Abbott Way, Stanford, California 94305, USA.
#------------------------------------------------------------------------------

# $Id: get-bufferhtml.ps1 162 2007-01-26 16:30:12Z adrian $

#------------------------------------------------------------------------------
# This script grabs text from the console buffer and outputs to the pipeline
# lines of HTML that represent it.
#
# Usage: get-bufferhtml [args]
#
# Where args are:
#
# -last n       - how many lines back from current line to grab
#                 default is (effectively) everything
# -all          - grab all lines in console, overrides -last
# -trim         - trims blank space from the right of each line
#                 this is ok unless you have lots of text with
#                 varying background colours
# -font s       - optional css font name. default is nothing which
#                 means the browser will use whatever is default for a
#                 <pre> tag. "Courier New" is quite a good alternative
# -fontsize s   - optional css font size, eg "9pt" or "80%"
# -style s      - optional addition css, eg "overflow:hidden"
# -palette p    - choose a colour palette, one of:
#                 "powershell" normal for a PowerShell window (ie with
#                              strange colours for darkmagenta and darkyellow
#                 "standard"   normal ansi colours as used by a standard
#                              cmd.exe session
#                 "print"      like powershell, but with colours handy
#                              for printing where you want to save ink.
#
# The output is one large wrapped <pre> tag to keep whitespace intact.
#

param(
  [int]$last = 50000,             
  [switch]$all,                   
  [switch]$trim,                  
  [string]$font=$null,            
  [string]$fontsize=$null,        
  [string]$style="",              
  [string]$palette="powershell"   
  )
$ui = $host.ui.rawui
[int]$start = 0
if ($all) { 
  [int]$end = $ui.BufferSize.Height  
  [int]$start = 0
}
else { 
  [int]$end = ($ui.CursorPosition.Y - 1)
  [int]$start = $end - $last
  if ($start -le 0) { $start = 0 }
}
$height = $end - $start
if ($height -le 0) {
  write-warning "There must be one or more lines to get"
  return
}
$width = $ui.BufferSize.Width
$dims = 0,$start,($width-1),($end-1)
$rect = new-object Management.Automation.Host.Rectangle -argumentList $dims
$cells = $ui.GetBufferContents($rect)

# set default colours
$fg = $ui.ForegroundColor; $bg = $ui.BackgroundColor
$defaultfg = $fg; $defaultbg = $bg

# character translations
# wordpress weirdness means I do special stuff for < and \
$cmap = @{
    [char]"<" = "<span>&lt;</span>"
    [char]"\" = "&#x5c;"
    [char]">" = "&gt;"
    [char]"'" = "&#39;"
    [char]"`"" = "&#34;"
    [char]"&" = "&amp;"
}

# console colour mapping
# the powershell console has some odd colour choices, 
# marked with a 6-char hex codes below
$palettes = @{}
$palettes.powershell = @{
    "Black"       ="#000"
    "DarkBlue"    ="#008"
    "DarkGreen"   ="#080"
    "DarkCyan"    ="#088"
    "DarkRed"     ="#800"
    "DarkMagenta" ="#012456"
    "DarkYellow"  ="#eeedf0"
    "Gray"        ="#ccc"
    "DarkGray"    ="#888"
    "Blue"        ="#00f"
    "Green"       ="#0f0"
    "Cyan"        ="#0ff"
    "Red"         ="#f00"
    "Magenta"     ="#f0f"
    "Yellow"      ="#ff0"
    "White"       ="#fff"
  }
# now a variation for the standard console (used by cmd.exe) based
# on ansi colours
$palettes.standard = ($palettes.powershell).Clone()
$palettes.standard.DarkMagenta = "#808"
$palettes.standard.DarkYellow = "#880"

# this is a weird one... takes the normal powershell one and
# inverts a few colours so normal ps1 output would save ink when
# printed (eg from a web page).
$palettes.print = ($palettes.powershell).Clone()
$palettes.print.DarkMagenta = "#eee"
$palettes.print.DarkYellow = "#000"
$palettes.print.Yellow = "#440"
$palettes.print.Black = "#fff"
$palettes.print.White = "#000"

$comap = $palettes[$palette]

# inner function to translate a console colour to an html/css one
function c2h{return $comap[[string]$args[0]]}
$f=""
if ($font) { $f += " font-family: `"$font`";" }
if ($fontsize) { $f += " font-size: $fontsize;" }
$line  = "<pre style='color: $(c2h $fg); background-color: $(c2h $bg);$f $style'>" 
for ([int]$row=0; $row -lt $height; $row++ ) {
  for ([int]$col=0; $col -lt $width; $col++ ) {
    $cell = $cells[$row,$col]
    # do we need to change colours?
    $cfg = [string]$cell.ForegroundColor
    $cbg = [string]$cell.BackgroundColor
    if ($fg -ne $cfg -or $bg -ne $cbg) {
      if ($fg -ne $defaultfg -or $bg -ne $defaultbg) { 
        $line += "</span>" # remove any specialisation
        $fg = $defaultfg; $bg = $defaultbg;
      }
      if ($cfg -ne $defaultfg -or $cbg -ne $defaultbg) { 
        # start a new colour span
        $line += "<span style='color: $(c2h $cfg); background-color: $(c2h $cbg)'>" 
      }
      $fg = $cfg
      $bg = $cbg
    }
    $ch = $cell.Character
    $ch2 = $cmap[$ch]; if ($ch2) { $ch = $ch2 }
    $line += $ch
  }
  if ($trim) { $line = $Line.TrimEnd() }
  $line
  $line=""
}
if ($fg -ne $defaultfg -or $bg -ne $defaultbg) { "</span>" } # close off any specialization of colour
"</pre>"


# SIG # Begin signature block
# MIINHAYJKoZIhvcNAQcCoIINDTCCDQkCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUfHVwpH0gt40fbHal93GVoqhC
# F5ygggpeMIIFJjCCBA6gAwIBAgIQDabkR8675p80ZdtFokcNRTANBgkqhkiG9w0B
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
# AQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMCMGCSqGSIb3DQEJBDEWBBTzvVnpROmx
# 3rGltrdn0bQtzPqWiDANBgkqhkiG9w0BAQEFAASCAQAy+BywrJhKwjPl/H5i5GvA
# ObslY/pQLtACL+xJNNuPYOGk1f5ku1CttARcJiVpTYEP/YqT2wv3SRq2To6UTVWh
# ZWgf86ukdCKNEQ5N6iphVc2cY8Y98hmNjFcsSCIlFeL6ocROY5wjsLR+JfAr+y2d
# 3j9ZBohKHvfvImvgyjGSyX763jsTGDcBCJq9MNeLEu/dF9+ZfMkftAE6mowb7MeP
# yABVsa44uUkmwbeXDTBkZEOOqcsL4bOLZqxKYjyfbgiDWnQy14HNZ++h+6cu3qOm
# Yl83CFq4WPUQ8b5Rs0ofJfC0GyacvHqQ2o6bSEBBNniyMWqu34AWq5RS8bZqbfvY
# SIG # End signature block
