####################################
#            WSUS-Cleanup          #
####################################


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
    #echo "Deutsche Text ausgabe " "`0"

    $IA64_text =" Itanium updates wurden abgelehnt"
    $ARM64_text =" ARM64 Updates wurden abgelehnt"
    $embedded_text =" Windows Embedded Updates wurden abgelehnt"
    $Office64_text =" MS Office 64-Bit wurden abgelehnt"
    $LanguagePack_text =" Language Interface Packs / Sprachpakete wurden abgelehnt"
    $Start_Name ="WSUS Reinigung gestartet"
    
    $IA64_titel ="Itanium updates werden abgelehnt"
    $ARM64_titel ="ARM64 Updates werden abgelehnt"
    $embedded_titel ="Windows Embedded Updates werden abgelehnt"
    $Office64_titel ="MS Office 64-Bit werden abgelehnt"
    $LanguagePack_titel ="Language Interface Packs / Sprachpakete werden abgelehnt"
    $working_declining = "Updates gefunden, Ablehnen..."
    $working_noup = "Es wurden keine Updates gefunden, die abgelehnt werden mussten."
}
else
{ 
    #echo "Englische Text ausgabe " "`0"

    $IA64_text =" Itanium updates were declined"
    $ARM64_text =" ARM64 Updates were declined"
    $embedded_text =" Windows Embedded Updates were declined"
    $Office64_text =" MS Office 64-Bit were declined"
    $LanguagePack_text =" Language Interface Packs were declined"
    $Start_Name ="WSUS cleaning started"

    $IA64_titel ="Checking for Updates: Itanium, IA64"
    $ARM64_titel ="Checking for Updates: ARM64"
    $embedded_titel ="Checking for Updates: Windwos Embedded"
    $Office64_titel ="Checking for Updates: Office x64"
    $LanguagePack_titel ="Checking for Updates: Language Packs"
    $working_declining = "Updates found. Declining..."
    $working_noup = "No updates found that needed declining."
}
echo "################## $Start_Name ##################"

# Connect to the WSUS 3.0 interface.
[reflection.assembly]::LoadWithPartialName("Microsoft.UpdateServices.Administration") | out-null
$WsusServerAdminProxy = [Microsoft.UpdateServices.Administration.AdminProxy]::GetUpdateServer($WsusServer,$UseSSL,$PortNumber);


# Suche Updates nach Namen und lehne diese ab

# Itanium/IA64

Write-Host -ForegroundColor Yellow "$IA64_titel"
    $Updates = $null;    # Reset variable
    $Updates = $WsusServerAdminProxy.GetUpdates() | ?{-not $_.IsDeclined -and $_.Title -match “ia64|itanium”}
    If($Updates) {
        Write-Host -ForegroundColor Green "$working_declining"
	    $Updates | %{$_.Decline()}
	    $Table = @{Name="Title";Expression={[string]$_.Title}},`
		    @{Name="KB";Expression={[string]$_.KnowledgebaseArticles}},`
		    @{Name="Classification";Expression={[string]$_.UpdateClassificationTitle}},`
		    @{Name="Product";Expression={[string]$_.ProductTitles}},`
		    @{Name="Family";Expression={[string]$_.ProductFamilyTitles}}
	    # Display to Console
            $Updates | Select $Table | sort Classification | ft Title,Classification,KB -AutoSize
    }
    Else { Write-Host -ForegroundColor DarkGreen "$working_noup" }


# ARM64

Write-Host -ForegroundColor Yellow "$ARM64_titel"
    $Updates = $null;    # Reset variable
    $Updates = $WsusServerAdminProxy.GetUpdates() | ?{-not $_.IsDeclined -and $_.Title -match “ARM64”}
    If($Updates) {
        Write-Host -ForegroundColor Green "$working_declining"
	    $Updates | %{$_.Decline()}
	    $Table = @{Name="Title";Expression={[string]$_.Title}},`
		    @{Name="KB";Expression={[string]$_.KnowledgebaseArticles}},`
		    @{Name="Classification";Expression={[string]$_.UpdateClassificationTitle}},`
		    @{Name="Product";Expression={[string]$_.ProductTitles}},`
		    @{Name="Family";Expression={[string]$_.ProductFamilyTitles}}
	    # Display to Console
            $Updates | Select $Table | sort Classification | ft Title,Classification,KB -AutoSize
    }
    Else { Write-Host -ForegroundColor DarkGreen "$working_noup" }


# Windows Embedded

Write-Host -ForegroundColor Yellow "$embedded_titel"
    $Updates = $null;    # Reset variable
    $Updates = $WsusServerAdminProxy.GetUpdates() | ?{-not $_.IsDeclined -and $_.Title -match “Windows Embedded Standard”}
    If($Updates) {
        Write-Host -ForegroundColor Green "$working_declining"
	    $Updates | %{$_.Decline()}
	    $Table = @{Name="Title";Expression={[string]$_.Title}},`
		    @{Name="KB";Expression={[string]$_.KnowledgebaseArticles}},`
		    @{Name="Classification";Expression={[string]$_.UpdateClassificationTitle}},`
		    @{Name="Product";Expression={[string]$_.ProductTitles}},`
		    @{Name="Family";Expression={[string]$_.ProductFamilyTitles}}
	    # Display to Console
            $Updates | Select $Table | sort Classification | ft Title,Classification,KB -AutoSize
    }
    Else { Write-Host -ForegroundColor DarkGreen "$working_noup" }


# MS Office 64-Bit

Write-Host -ForegroundColor Yellow "$Office64_titel"
    $Updates = $null;    # Reset variable
    $Updates = $WsusServerAdminProxy.GetUpdates() | ?{-not $_.IsDeclined -and $_.Title -match "Excel|Lync|Office|Outlook|Powerpoint|Visio|word" -and $_.Title -match "64-bit"}
    If($Updates) {
        Write-Host -ForegroundColor Green "$working_declining"
	    $Updates | %{$_.Decline()}
	    $Table = @{Name="Title";Expression={[string]$_.Title}},`
		    @{Name="KB";Expression={[string]$_.KnowledgebaseArticles}},`
		    @{Name="Classification";Expression={[string]$_.UpdateClassificationTitle}},`
		    @{Name="Product";Expression={[string]$_.ProductTitles}},`
		    @{Name="Family";Expression={[string]$_.ProductFamilyTitles}}
	    # Display to Console
            $Updates | Select $Table | sort Classification | ft Title,Classification,KB -AutoSize
    }
    Else { Write-Host -ForegroundColor DarkGreen "$working_noup" }


# Language Pack

Write-Host -ForegroundColor Yellow "$LanguagePack_titel"
    $Updates = $null;    # Reset variable
    $Updates = $WsusServerAdminProxy.GetUpdates() | ?{-not $_.IsDeclined -and $_.Title -match "Language Interface Pack|LanguageInterfacePack|LanguageFeatureOnDemand|Sprachpaket f|Language Pack - Windows 10 -|LanguagePack - Windows 10 Insider Preview|Lan Pack (Language Features)"}
    If($Updates) {
        Write-Host -ForegroundColor Green "$working_declining"
	    $Updates | %{$_.Decline()}
	    $Table = @{Name="Title";Expression={[string]$_.Title}},`
		    @{Name="KB";Expression={[string]$_.KnowledgebaseArticles}},`
		    @{Name="Classification";Expression={[string]$_.UpdateClassificationTitle}},`
		    @{Name="Product";Expression={[string]$_.ProductTitles}},`
		    @{Name="Family";Expression={[string]$_.ProductFamilyTitles}}
	    # Display to Console
            $Updates | Select $Table | sort Classification | ft Title,Classification,KB -AutoSize
    }
    Else { Write-Host -ForegroundColor DarkGreen "$working_noup" }


Invoke-WsusServerCleanup -UpdateServer $WSUS -DeclineExpiredUpdates -DeclineSupersededUpdates -CleanupObsoleteUpdates -CleanupUnneededContentFiles -CompressUpdates


