set-executionpolicy RemoteSigned

$WsusServer = ([system.net.dns]::GetHostByName('localhost')).hostname
$UseSSL = $false
$WsusName = localhost
$PortNumber = 8530
$WSUS = Get-WsusServer -Name $WsusName -PortNumber $PortNumber
$TrialRun = 0
# 1 = Yes
# 0 = No#

"WSUS Reinigung gestartet"

# Connect to the WSUS 3.0 interface.
[reflection.assembly]::LoadWithPartialName("Microsoft.UpdateServices.Administration") | out-null
$WsusServerAdminProxy = [Microsoft.UpdateServices.Administration.AdminProxy]::GetUpdateServer($WsusServer,$UseSSL,$PortNumber);

# Searching in just the title of the update
# Itanium/IA64
$itanium = $WsusServerAdminProxy.GetUpdates() | ?{-not $_.IsDeclined -and $_.Title -match “ia64|itanium”}
$IA64_counted = $itanium.count
    If ($itanium.count -lt 1)
        {
            $IA64_counted = 0
        }
    If ($TrialRun -eq 0 -and $itanium.count -gt 0)
        {
            $itanium  | %{$_.Decline()}
        }

"$IA64_counted Itanium updates wurden abgelehnt"

# MS Office 64-Bit
$Office64 = $WsusServerAdminProxy.GetUpdates() | ?{-not $_.IsDeclined -and $_.Title -match “Excel|Lync|Office|Outlook|Powerpoint|Visio|word” -and $_.Title -match "64-bit"}
$Office64_count = $Office64.count
    If ($Office64.count -lt 1)
        {
            $Office64_count = 0
        }
    If ($TrialRun -eq 0 -and $Office64.count -gt 0)
        {
            $Office64 | %{$_.Decline()}
        }

"$Office64_count MS Office 64-Bit wurden abgelehnt"

# Searching in just the title of the update
# Language Pack
$LanguagePack = $WsusServerAdminProxy.GetUpdates() | ?{-not $_.IsDeclined -and $_.Title -match “Language Interface Pack”}
$LanguagePack_counted = $LanguagePack.count
    If ($LanguagePack.count -lt 1)
        {
            $LanguagePack_counted = 0
        }
    If ($TrialRun -eq 0 -and $LanguagePack.count -gt 0)
        {
            $LanguagePack  | %{$_.Decline()}
        }

"$LanguagePack_counted Language Interface Packs wurden abgelehnt"

# Searching in just the title of the update
# Language Pack - Windows 10 -
$LanguagePackW10 = $WsusServerAdminProxy.GetUpdates() | ?{-not $_.IsDeclined -and $_.Title -match “Language Pack - Windows 10 -”}
$LanguagePackW10_counted = $LanguagePackW10.count
    If ($LanguagePackW10.count -lt 1)
        {
            $LanguagePackW10_counted = 0
        }
    If ($TrialRun -eq 0 -and $LanguagePackW10.count -gt 0)
        {
            $LanguagePackW10  | %{$_.Decline()}
        }

"$LanguagePackW10_counted Language Pack - Windows 10 wurden abgelehnt"

# Searching in just the title of the update
# LanguagePack - Windows 10 Insider Preview
$LPW10InsiderPreview = $WsusServerAdminProxy.GetUpdates() | ?{-not $_.IsDeclined -and $_.Title -match “LanguagePack - Windows 10 Insider Preview”}
$LPW10InsiderPreview_counted = $LPW10InsiderPreview.count
    If ($LPW10InsiderPreview.count -lt 1)
        {
            $LPW10InsiderPreview_counted = 0
        }
    If ($TrialRun -eq 0 -and $LPW10InsiderPreview.count -gt 0)
        {
            $LPW10InsiderPreview  | %{$_.Decline()}
        }
"$LPW10InsiderPreview_counted LanguagePack - Windows 10 Insider Preview wurden abgelehnt"

# Searching in just the title of the update
# LanguageInterfacePack - Windows 10 Insider Preview
$LiPW10InsiderPreview = $WsusServerAdminProxy.GetUpdates() | ?{-not $_.IsDeclined -and $_.Title -match “LanguageInterfacePack - Windows 10 Insider Preview”}
$LiPW10InsiderPreview_counted = $LiPW10InsiderPreview.count
    If ($LiPW10InsiderPreview.count -lt 1)
        {
            $LiPW10InsiderPreview_counted = 0
        }
    If ($TrialRun -eq 0 -and $LiPW10InsiderPreview.count -gt 0)
        {
            $LiPW10InsiderPreview  | %{$_.Decline()}
        }

"$LiPW10InsiderPreview_counted LanguageInterfacePack - Windows 10 Insider Preview wurden abgelehnt"

# Searching in just the title of the update
# LanguageInterfacePack - Windows 10 Insider Preview
$LiPW10InsiderPreview = $WsusServerAdminProxy.GetUpdates() | ?{-not $_.IsDeclined -and $_.Title -match “LanguageInterfacePack - Windows 10 Insider Preview”}
$LiPW10InsiderPreview_counted = $LiPW10InsiderPreview.count
    If ($LiPW10InsiderPreview.count -lt 1)
        {
            $LiPW10InsiderPreview_counted = 0
        }
    If ($TrialRun -eq 0 -and $LiPW10InsiderPreview.count -gt 0)
        {
            $LiPW10InsiderPreview  | %{$_.Decline()}
        }


# Searching in just the title of the update
# “ARM64
$arm64 = $WsusServerAdminProxy.GetUpdates() | ?{-not $_.IsDeclined -and $_.Title -match “ARM64”}
$ARM64_counted = $arm64.count
    If ($ARM64.count -lt 1)
        {
            $ARM64_counted = 0
        }
    If ($TrialRun -eq 0 -and $arm64.count -gt 0)
        {
            $ARM64  | %{$_.Decline()}
        }

"$ARM64_counted ARM64 Updates wurden abgelehnt"
"

Invoke-WsusServerCleanup -UpdateServer $WSUS -DeclineExpiredUpdates -DeclineSupersededUpdates -CleanupObsoleteUpdates -CleanupUnneededContentFiles -CompressUpdates
