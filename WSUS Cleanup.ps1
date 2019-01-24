set-executionpolicy RemoteSigned

$WsusServer = ([system.net.dns]::GetHostByName('localhost')).hostname
$UseSSL = $false
$WsusName = "localhost"
$PortNumber = 8530
$WSUS = Get-WsusServer -Name $WsusName -PortNumber $PortNumber
$TrialRun = 0
# 1 = Yes
# 0 = No#

#Ermittle Sprache
[String]$Sprache = [CultureInfo]::InstalledUICulture | select TwoLetter* 
$Sprache = $Sprache.Remove(0,27)
$Sprache= $Sprache.Replace("}",$null)


####Sprachtexte anpassen
if($Sprache -eq "de")
{
    echo "Deutsche Text ausgabe " "`0"
    $IA64_text =" Itanium updates wurden abgelehnt"
    $ARM64_text =" ARM64 Updates wurden abgelehnt"
    $embedded_text =" Windows Embedded Updates wurden abgelehnt"
    $Office64_text =" MS Office 64-Bit wurden abgelehnt"
    $LanguagePack_text =" Language Interface Packs / Sprachpakete wurden abgelehnt"
    $Start_Name =" WSUS Reinigung gestartet"
}

if($Sprache -eq "en")
{
    echo "Englische Text ausgabe " "`0"
    $IA64_text =" Itanium updates were declined"
    $ARM64_text =" ARM64 Updates were declined"
    $embedded_text =" Windows Embedded Updates were declined"
    $Office64_text =" MS Office 64-Bit were declined"
    $LanguagePack_text =" Language Interface Packs were declined"
    $Start_Name =" WSUS cleaning started"
}


echo "$Start_Name"

# Connect to the WSUS 3.0 interface.
[reflection.assembly]::LoadWithPartialName("Microsoft.UpdateServices.Administration") | out-null
$WsusServerAdminProxy = [Microsoft.UpdateServices.Administration.AdminProxy]::GetUpdateServer($WsusServer,$UseSSL,$PortNumber);




# Suche Updates nach Namen und lehne diese ab

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

echo "$IA64_counted $IA64_text"


# ARM64
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

echo "$ARM64_counted $ARM64_text"


# Windows Embedded
$embedded = $WsusServerAdminProxy.GetUpdates() | ?{-not $_.IsDeclined -and $_.Title -match “Windows Embedded Standard”}
$embedded_counted = $embedded.count
    If ($embedded.count -lt 1)
        {
            $embedded_counted = 0
        }
    If ($TrialRun -eq 0 -and $embedded.count -gt 0)
        {
            $embedded  | %{$_.Decline()}
        }

echo "$embedded_counted $embedded_text"


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

echo "$Office64_count $Office64_text"


# Language Pack
$LanguagePack = $WsusServerAdminProxy.GetUpdates() | ?{-not $_.IsDeclined -and $_.Title -match “Language Interface Pack|LanguageInterfacePack|LanguageFeatureOnDemand|Sprachpaket für|Language Pack - Windows 10 -|LanguagePack - Windows 10 Insider Preview|Lan Pack (Language Features)”}
$LanguagePack_counted = $LanguagePack.count
    If ($LanguagePack.count -lt 1)
        {
            $LanguagePack_counted = 0
        }
    If ($TrialRun -eq 0 -and $LanguagePack.count -gt 0)
        {
            $LanguagePack  | %{$_.Decline()}
        }

echo "$LanguagePack_counted $LanguagePack_text"


Invoke-WsusServerCleanup -UpdateServer $WSUS -DeclineExpiredUpdates -DeclineSupersededUpdates -CleanupObsoleteUpdates -CleanupUnneededContentFiles -CompressUpdates
