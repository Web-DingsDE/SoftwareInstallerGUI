##### Individualanpassungen


$InstallPath = "\\Filesrv01\install"
#$workdir = "\\DOMSRV01\install\$env:UserName"
$InstallPathBRZclient = "\\Filesrv01\brz32"

$LogWindowTitel = "ClientInstaller"


Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()


Import-Module BITSTransfer

#
#
#
#
#
#
######################
##### Sources

function Sources{
    Set-Variable -Name path -Value $TextBoxPermInstallPath.text -Scope global

    Set-Variable -Name sourceTeamViewerQS -Value "https://download.teamviewer.com/download/TeamViewerQS.exe" -Scope global
    Set-Variable -Name destinationTeamViewerQS -Value "$path\TeamViewerQS.exe" -Scope global

    Set-Variable -Name sourceFirefox -Value "https://download.mozilla.org/?product=firefox-latest-ssl&os=win64&lang=de" -Scope global
    Set-Variable -Name destinationFirefox -Value "$path\Firefox.exe" -Scope global

    Set-Variable -Name source7z -Value "http://www.7-zip.org/a/7z1604-x64.msi" -Scope global
    Set-Variable -Name destination7z -Value "$path\7z.msi" -Scope global

    Set-Variable -Name sourceGreenshot -Value "https://dl.dropboxusercontent.com/s/pd7fu4tcg8krvec/Greenshot.exe?dl=0" -Scope global
    #Set-Variable -Name ZIPdestinationGreenshot -Value "$path\O365ProPlusRetail.zip" -Scope global
    #Set-Variable -Name ZIPContentdestinationGreenshot -Value "$path\O365ProPlusRetail" -Scope global
    Set-Variable -Name destinationGreenshot -Value "$path\Greenshot.exe" -Scope global
    
    Set-Variable -Name sourceCDburnerXP -Value "https://cdburnerxp.se/downloadsetup.exe" -Scope global
    Set-Variable -Name destinationCDburnerXP -Value "$path\CDburnerXP.exe" -Scope global

    Set-Variable -Name sourceO365ProPlusRetail -Value "https://dl.dropboxusercontent.com/s/y9rg75yi41c9082/O365ProPlusRetail.exe?dl=0" -Scope global
    #$ZIPdestinationO365ProPlusRetail -Value "$path\O365ProPlusRetail.zip"
    #$ZIPContentdestinationO365ProPlusRetail -Value "$path\O365ProPlusRetail"
    Set-Variable -Name destinationO365ProPlusRetail -Value "$path\O365ProPlusRetail.exe" -Scope global

    Set-Variable -Name destinationAcrobatreaderDC -Value "$path\AdobeDC.exe" -Scope global
    Set-Variable -Name sourceAcrobatreaderDC -Value "http://ardownload.adobe.com/pub/adobe/reader/win/AcrobatDC/1502320053/AcroRdrDC1502320053_de_DE.exe" -Scope global

    Set-Variable -Name pathBRZ -Value $TextBoxPermInstallPathBRZclient.Text  -Scope global  
    Set-Variable -Name destinationBRZclient -Value "$pathBRZ\Client\Disk1\setup.exe" -Scope global

    Set-Variable -Name sourceGoogleEarth -Value "https://dl.google.com/dl/earth/client/advanced/current/googleearthprowin-7.3.2-x64.exe" -Scope global    
    Set-Variable -Name destinationGoogleEarth -Value "$path\GoogleEarth.exe" -Scope global

    Set-Variable -Name sourceGoogleChrome -Value "https://dl.google.com/tag/s/appguid%3D%7B8A69D345-D564-463C-AFF1-A69D9E530F96%7D%26iid%3D%7B03FE9563-80F9-119F-DA3D-72FBBB94BC26%7D%26lang%3Den%26browser%3D4%26usagestats%3D0%26appname%3DGoogle%2520Chrome%26needsadmin%3Dprefers%26ap%3Dx64-stable/dl/chrome/install/googlechromestandaloneenterprise64.msi" -Scope global
    Set-Variable -Name destinationGoogleChrome -Value "$path\GoogleChrome.msi" -Scope global
}
#
#
#
#
#
##################
#
#
#

#####GUI01#####

$GlobalFont = "Arial,8"


$CheckBoxHeight = "25"
$CheckBoxWidth = "165"
$CheckBoxVertical = "30"

$TextBoxHeight = "25"
$TextBoxWidth = "165"
$TextBoxVertical = "30"

##### Fenster

$form1                           = New-Object system.Windows.Forms.Form
$form1.ClientSize                = '800,600'
$form1.text                      = "Software Setup"
$form1.TopMost                   = $false


##### fenster Inhalt

# Inhaltsfenster
$LogWindow                          = New-Object system.Windows.Forms.TextBox
$LogWindow.multiline                = $true
$LogWindow.Scrollbars               = "Vertical"
$LogWindow.width                    = 580
$LogWindow.height                   = 450
$LogWindow.location                 = New-Object System.Drawing.Point(200,25)
$LogWindow.Font                     = $GlobalFont
$LogWindow.ForeColor                = "Black"
$LogWindow.BackColor                = "White"

##### CheckBox


###TeamViewerQS
$CheckBoxTeamViewerQS             = New-Object system.Windows.Forms.CheckBox
##edit
$CheckBoxTeamViewerQS.text        = "TeamViewerQS"
$CheckBoxTeamViewerQS.location    = New-Object System.Drawing.Point($CheckBoxVertical,75)
##notedit
$CheckBoxTeamViewerQS.AutoSize    = $true
$CheckBoxTeamViewerQS.Checked     = $true
$CheckBoxTeamViewerQS.width       = $CheckBoxWidth
$CheckBoxTeamViewerQS.height      = $CheckBoxHeight
$CheckBoxTeamViewerQS.Font        = $GlobalFont


###RenameComputer
$CheckBoxRenameComputer            = New-Object system.Windows.Forms.CheckBox
##edit
$CheckBoxRenameComputer.text        = "RenameComputer"
$CheckBoxRenameComputer.location    = New-Object System.Drawing.Point($CheckBoxVertical,100)
##notedit
$CheckBoxRenameComputer.AutoSize    = $true
$CheckBoxRenameComputer.Checked     = $false
$CheckBoxRenameComputer.width       = $CheckBoxWidth
$CheckBoxRenameComputer.height      = $CheckBoxHeight
$CheckBoxRenameComputer.Font        = $GlobalFont


###Firefox 
$CheckBoxFirefox             = New-Object system.Windows.Forms.CheckBox
##edit
$CheckBoxFirefox.text        = "Firefox"
$CheckBoxFirefox.location    = New-Object System.Drawing.Point($CheckBoxVertical,125)
##notedit
$CheckBoxFirefox.AutoSize    = $true
$CheckBoxFirefox.Checked     = $true
$CheckBoxFirefox.width       = $CheckBoxWidth
$CheckBoxFirefox.height      = $CheckBoxHeight
$CheckBoxFirefox.Font        = $GlobalFont




###7z
$CheckBox7z             = New-Object system.Windows.Forms.CheckBox
##edit
$CheckBox7z.text        = "7z"
$CheckBox7z.location    = New-Object System.Drawing.Point($CheckBoxVertical,150)
##notedit
$CheckBox7z.AutoSize    = $true
$CheckBox7z.Checked     = $true
$CheckBox7z.width       = $CheckBoxWidth
$CheckBox7z.height      = $CheckBoxHeight
$CheckBox7z.Font        = $GlobalFont

###Greenshot
$CheckBoxGreenshot             = New-Object system.Windows.Forms.CheckBox
##edit
$CheckBoxGreenshot.text        = "Greenshot"
$CheckBoxGreenshot.location    = New-Object System.Drawing.Point($CheckBoxVertical,175)
##notedit
$CheckBoxGreenshot.AutoSize    = $true
$CheckBoxGreenshot.Checked     = $true
$CheckBoxGreenshot.width       = $CheckBoxWidth
$CheckBoxGreenshot.height      = $CheckBoxHeight
$CheckBoxGreenshot.Font        = $GlobalFont

###CDburnerXP
$CheckBoxCDburnerXP             = New-Object system.Windows.Forms.CheckBox
##edit
$CheckBoxCDburnerXP.text        = "CDburnerXP"
$CheckBoxCDburnerXP.location    = New-Object System.Drawing.Point($CheckBoxVertical,200)
##notedit
$CheckBoxCDburnerXP.AutoSize    = $true
$CheckBoxCDburnerXP.Checked     = $true
$CheckBoxCDburnerXP.width       = $CheckBoxWidth
$CheckBoxCDburnerXP.height      = $CheckBoxHeight
$CheckBoxCDburnerXP.Font        = $GlobalFont

###O365ProPlusRetail
$CheckBoxO365ProPlusRetail             = New-Object system.Windows.Forms.CheckBox
##edit
$CheckBoxO365ProPlusRetail.text        = "O365ProPlusRetail"
$CheckBoxO365ProPlusRetail.location    = New-Object System.Drawing.Point($CheckBoxVertical,225)
##notedit
$CheckBoxO365ProPlusRetail.AutoSize    = $true
$CheckBoxO365ProPlusRetail.Checked     = $true
$CheckBoxO365ProPlusRetail.width       = $CheckBoxWidth
$CheckBoxO365ProPlusRetail.height      = $CheckBoxHeight
$CheckBoxO365ProPlusRetail.Font        = $GlobalFont

###AcrobatReaderDC
$CheckBoxAcrobatReaderDC           = New-Object system.Windows.Forms.CheckBox
##edit
$CheckBoxAcrobatReaderDC.text      = "AcrobatReaderDC"
$CheckBoxAcrobatReaderDC.location  = New-Object System.Drawing.Point($CheckBoxVertical,250)
##notedit
$CheckBoxAcrobatReaderDC.AutoSize  = $true
$CheckBoxAcrobatReaderDC.Checked   = $true
$CheckBoxAcrobatReaderDC.width     = $CheckBoxWidth
$CheckBoxAcrobatReaderDC.height    = $CheckBoxHeight
$CheckBoxAcrobatReaderDC.Font      = $GlobalFont

###BRZclient
$CheckBoxBRZclient           = New-Object system.Windows.Forms.CheckBox
##edit
$CheckBoxBRZclient.text      = "BRZclient"
$CheckBoxBRZclient.location  = New-Object System.Drawing.Point($CheckBoxVertical,275)
##notedit
$CheckBoxBRZclient.AutoSize  = $true
$CheckBoxBRZclient.Checked   = $true
$CheckBoxBRZclient.width     = $CheckBoxWidth
$CheckBoxBRZclient.height    = $CheckBoxHeight
$CheckBoxBRZclient.Font      = $GlobalFont

###GoogleEarth
$CheckBoxGoogleEarth             = New-Object system.Windows.Forms.CheckBox
##edit
$CheckBoxGoogleEarth.text        = "GoogleEarth"
$CheckBoxGoogleEarth.location    = New-Object System.Drawing.Point($CheckBoxVertical,300)
##notedit
$CheckBoxGoogleEarth.AutoSize    = $true
$CheckBoxGoogleEarth.Checked     = $true
$CheckBoxGoogleEarth.width       = $CheckBoxWidth
$CheckBoxGoogleEarth.height      = $CheckBoxHeight
$CheckBoxGoogleEarth.Font        = $GlobalFont

###GoogleChrome
$CheckBoxGoogleChrome             = New-Object system.Windows.Forms.CheckBox
##edit
$CheckBoxGoogleChrome.text        = "GoogleChrome"
$CheckBoxGoogleChrome.location    = New-Object System.Drawing.Point($CheckBoxVertical,325)
##notedit
$CheckBoxGoogleChrome.AutoSize    = $true
$CheckBoxGoogleChrome.Checked     = $true
$CheckBoxGoogleChrome.width       = $CheckBoxWidth
$CheckBoxGoogleChrome.height      = $CheckBoxHeight
$CheckBoxGoogleChrome.Font        = $GlobalFont



##### TextBox



###InstallPath
$TextBoxPermInstallPath            = New-Object system.Windows.Forms.TextBox
##edit
$TextBoxPermInstallPath.location   = New-Object System.Drawing.Point($TextBoxVertical,25)
$TextBoxPermInstallPath.text       = "$InstallPath"
##noedit
$TextBoxPermInstallPath.multiline  = $false
$TextBoxPermInstallPath.width      = $TextBoxWidth
$TextBoxPermInstallPath.height     = $TextBoxHeight
$TextBoxPermInstallPath.Font       = $GlobalFont
#$form1.TextBoxPermInstallPath = $TextBoxPermInstallPath
$form1.Controls.Add($TextBoxPermInstallPath)



###InstallPathBRZclient
$TextBoxPermInstallPathBRZclient           = New-Object system.Windows.Forms.TextBox
##edit
$TextBoxPermInstallPathBRZclient.location   = New-Object System.Drawing.Point($TextBoxVertical,50)
$TextBoxPermInstallPathBRZclient.text       = "$InstallPathBRZclient"
##noedit
$TextBoxPermInstallPathBRZclient.multiline  = $false
$TextBoxPermInstallPathBRZclient.width      = $TextBoxWidth
$TextBoxPermInstallPathBRZclient.height     = $TextBoxHeight
$TextBoxPermInstallPathBRZclient.Font       = $GlobalFont
#$form1.TextBoxPermInstallPathBRZclient = $TextBoxPermInstallPathBRZclient
$form1.Controls.Add($TextBoxPermInstallPathBRZclient)



##### Buttons

$ButtonStart                     = New-Object system.Windows.Forms.Button
$ButtonStart.text                = "Start"
$ButtonStart.width               = 120
$ButtonStart.height              = 50
$ButtonStart.location            = New-Object System.Drawing.Point($CheckBoxVertical,350)
$ButtonStart.Font                = 'Microsoft Sans Serif,10,style=Bold'

$ButtonDownload                     = New-Object system.Windows.Forms.Button
$ButtonDownload.text                = "Download"
$ButtonDownload.width               = 120
$ButtonDownload.height              = 50
$ButtonDownload.location            = New-Object System.Drawing.Point($CheckBoxVertical,525)
$ButtonDownload.Font                = 'Microsoft Sans Serif,10'


####Progressbar

$ProgressBar1                    = New-Object system.Windows.Forms.ProgressBar
$ProgressBar1.width              = 575
$ProgressBar1.height             = 30
$ProgressBar1.location           = New-Object System.Drawing.Point(200,550)
$ProgressBar1.Value              = 0

function timestamp {
    return $(Get-Date -Format G)
}

function Test-Admin {
  $currentUser = New-Object Security.Principal.WindowsPrincipal $([Security.Principal.WindowsIdentity]::GetCurrent())
  $currentUser.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
}

function TestPath{
    Test-Path -Path $TextBoxPermInstallPath.text 
}



#
#
#
#
#
#
#########################Start of only Download
#
#
#
#
#
#
function DownloadTeamViewerQS{

# Download the installer
    If (Test-Path -Path $destinationTeamViewerQS){ 
    Write-Host "$destinationTeamViewerQS exists" -ForegroundColor Green
    $LogWindow.text += "$(timestamp):   $destinationTeamViewerQS inst vorhanden. `r`n"
    }
    ELSE{ 

    # Download the installer
    $JobDownloadTeamViewerQS = Start-BitsTransfer -Source $sourceTeamViewerQS -Destination $destinationTeamViewerQS -Asynchronous
    $LogWindow.text += "$(timestamp):   $destinationTeamViewerQS wird herunter geladen. `r`n"

    While( ($JobDownloadTeamViewerQS.JobState.ToString() -eq 'Transferring') -or ($JobDownloadTeamViewerQS.JobState.ToString() -eq 'Connecting') ){

    $ProgressBar1.Value = [int](($JobDownloadTeamViewerQS.BytesTransferred*100) / $JobDownloadTeamViewerQS.BytesTotal)
    }
    Complete-BitsTransfer -BitsJob $JobDownloadTeamViewerQS
                                #Invoke-WebRequest $sourceTeamViewerQS -OutFile $destinationTeamViewerQS
    

        If (Test-Path -Path $destinationTeamViewerQS){
                #Invoke-WebRequest $sourceTeamViewerQS -OutFile $destinationTeamViewerQS
            Write-Host "$destinationTeamViewerQS downloaded" -ForegroundColor Green
            $LogWindow.text += "$(timestamp):   $destinationTeamViewerQS download abgeschlossen. `r`n"
        }
        else {
            Write-Host "$destinationTeamViewerQS not downloaded" -ForegroundColor Green
            $LogWindow.text += "$(timestamp):   $destinationTeamViewerQS download abgbrochen. `r`n"    
        }
    }   

    



}

function DownloadFirefox {

    # Download the installer
        If (Test-Path -Path $destinationFirefox){ 
        Write-Host "$destinationFirefox exists" -ForegroundColor Green
        $LogWindow.text += "$(timestamp):   $destinationFirefox inst vorhanden. `r`n"
        }
        ELSE{ 
    
        # Download the installer
        $JobDownloadFirefox = Start-BitsTransfer -Source $sourceFirefox -Destination $destinationFirefox -Asynchronous
        $LogWindow.text += "$(timestamp):   $destinationFirefox wird herunter geladen. `r`n"
    
        While( ($JobDownloadFirefox.JobState.ToString() -eq 'Transferring') -or ($JobDownloadFirefox.JobState.ToString() -eq 'Connecting') ){
    
        $ProgressBar1.Value = [int](($JobDownloadFirefox.BytesTransferred*100) / $JobDownloadFirefox.BytesTotal)
        }
        Complete-BitsTransfer -BitsJob $JobDownloadFirefox
        
    
            If (Test-Path -Path $destinationFirefox){
                    #Invoke-WebRequest $sourceFirefox -OutFile $destinationFirefox
                Write-Host "$destinationFirefox downloaded" -ForegroundColor Green
                $LogWindow.text += "$(timestamp):   $destinationFirefox download abgeschlossen. `r`n"
            }
            else {
                Write-Host "$destinationFirefox not downloaded" -ForegroundColor Green
                $LogWindow.text += "$(timestamp):   $destinationFirefox download abgbrochen. `r`n"    
            }
        }   
    
        
    }

#Erledigt
function Download7z {

# Download the installer
    If (Test-Path -Path $destination7z){ 
    Write-Host "$destination7z exists" -ForegroundColor Green
    $LogWindow.text += "$(timestamp):   $destination7z inst vorhanden. `r`n"
    }
    ELSE{ 

    # Download the installer
    $JobDownload7z = Start-BitsTransfer -Source $source7z -Destination $destination7z -Asynchronous
    $LogWindow.text += "$(timestamp):   $destination7z wird herunter geladen. `r`n"

    While( ($JobDownload7z.JobState.ToString() -eq 'Transferring') -or ($JobDownload7z.JobState.ToString() -eq 'Connecting') ){

    $ProgressBar1.Value = [int](($JobDownload7z.BytesTransferred*100) / $JobDownload7z.BytesTotal)
    }
    Complete-BitsTransfer -BitsJob $JobDownload7z
    

        If (Test-Path -Path $destination7z){
                #Invoke-WebRequest $source7z -OutFile $destination7z
            Write-Host "$destination7z downloaded" -ForegroundColor Green
            $LogWindow.text += "$(timestamp):   $destination7z download abgeschlossen. `r`n"
        }
        else {
            Write-Host "$destination7z not downloaded" -ForegroundColor Green
            $LogWindow.text += "$(timestamp):   $destination7z download abgbrochen. `r`n"    
        }
    }   

    
}

#Erledigt
function RenameComputer {
    $delimiter = '-'
    $pwqusers = quser | Select-Object -skip 1
    $pwGetDate = Get-Date -Format "yyyy"
    $pwquserfmt = $pwqusers.substring(1,22).Trim() -join ","
    $pwqyearfmt = $pwGetDate.substring(2,2).Trim() -join ","
    $NewName = $pwquserfmt + $delimiter + $pwqyearfmt
    IF($pwquserfmt -eq "Administrator"){
    $LogWindow.text += "$(timestamp):   $NewName ist kein gültiger Name `r`n"
    }
    Else{
    Rename-Computer -NewName $NewName
    }
}

#Erledigt
function DownloadGreenshot{

# Download the installer
    If (Test-Path -Path $destinationGreenshot){ 
    Write-Host "$destinationGreenshot exists" -ForegroundColor Green
    $LogWindow.text += "$(timestamp):   $destinationGreenshot inst vorhanden. `r`n"
    }
    ELSE{ 

    # Download the installer
    $JobDownloadGreenshot = Start-BitsTransfer -Source $sourceGreenshot -Destination $destinationGreenshot -Asynchronous
    $LogWindow.text += "$(timestamp):   $destinationGreenshot wird herunter geladen. `r`n"

    While( ($JobDownloadGreenshot.JobState.ToString() -eq 'Transferring') -or ($JobDownloadGreenshot.JobState.ToString() -eq 'Connecting') ){

    $ProgressBar1.Value = [int](($JobDownloadGreenshot.BytesTransferred*100) / $JobDownloadGreenshot.BytesTotal)
    }
    Complete-BitsTransfer -BitsJob $JobDownloadGreenshot
    

        If (Test-Path -Path $destinationGreenshot){
                                #Invoke-WebRequest $sourceGreenshot -OutFile $destinationGreenshot
            Write-Host "$destinationGreenshot downloaded" -ForegroundColor Green
            $LogWindow.text += "$(timestamp):   $destinationGreenshot download abgeschlossen. `r`n"
        }
        else {
            Write-Host "$destinationGreenshot not downloaded" -ForegroundColor Green
            $LogWindow.text += "$(timestamp):   $destinationGreenshot download abgbrochen. `r`n"    
        }
    }
}




    




function DownloadCDburnerXP{

    # Download the installer
    If (Test-Path -Path $destinationCDburnerXP){ 
    Write-Host "$destinationCDburnerXP exists" -ForegroundColor Green
    $LogWindow.text += "$(timestamp):   $destinationCDburnerXP inst vorhanden. `r`n"
    }
    ELSE{ 

    # Download the installer
    $JobDownloadCDburnerXP = Start-BitsTransfer -Source $sourceCDburnerXP -Destination $destinationCDburnerXP -Asynchronous
    $LogWindow.text += "$(timestamp):   $destinationCDburnerXP wird herunter geladen. `r`n"

    While( ($JobDownloadCDburnerXP.JobState.ToString() -eq 'Transferring') -or ($JobDownloadCDburnerXP.JobState.ToString() -eq 'Connecting') ){

    $ProgressBar1.Value = [int](($JobDownloadCDburnerXP.BytesTransferred*100) / $JobDownloadCDburnerXP.BytesTotal)
    }
    Complete-BitsTransfer -BitsJob $JobDownloadCDburnerXP
    

        If (Test-Path -Path $destinationCDburnerXP){
                #Invoke-WebRequest $sourceCDburnerXP -OutFile $destinationCDburnerXP
            Write-Host "$destinationCDburnerXP downloaded" -ForegroundColor Green
            $LogWindow.text += "$(timestamp):   $destinationCDburnerXP download abgeschlossen. `r`n"
        }
        else {
            Write-Host "$destinationCDburnerXP not downloaded" -ForegroundColor Green
            $LogWindow.text += "$(timestamp):   $destinationCDburnerXP download abgbrochen. `r`n"    
        }
    }


}





    



#Erledigt
function DownloadO365ProPlusRetail{


    If (Test-Path -Path $destinationO365ProPlusRetail){ 
    Write-Host "$destinationO365ProPlusRetail exists" -ForegroundColor Green
    $LogWindow.text += "$(timestamp):   $destinationO365ProPlusRetail inst vorhanden. `r`n"
    }
    ELSE{ 

    # Download the installer
    $JobDownloadO365ProPlusRetail = Start-BitsTransfer -Source $sourceO365ProPlusRetail -Destination $destinationO365ProPlusRetail -Asynchronous
    $LogWindow.text += "$(timestamp):   $destinationO365ProPlusRetail wird herunter geladen. `r`n"

    While( ($JobDownloadO365ProPlusRetail.JobState.ToString() -eq 'Transferring') -or ($JobDownloadO365ProPlusRetail.JobState.ToString() -eq 'Connecting') ){

    $ProgressBar1.Value = [int](($JobDownloadO365ProPlusRetail.BytesTransferred*100) / $JobDownloadO365ProPlusRetail.BytesTotal)
    }
    Complete-BitsTransfer -BitsJob $JobDownloadO365ProPlusRetail
    

        If (Test-Path -Path $destinationO365ProPlusRetail){
                #Invoke-WebRequest $sourceO365ProPlusRetail -OutFile $destinationO365ProPlusRetail
            Write-Host "$destinationO365ProPlusRetail downloaded" -ForegroundColor Green
            $LogWindow.text += "$(timestamp):   $destinationO365ProPlusRetail download abgeschlossen. `r`n"
        }
        else {
            Write-Host "$destinationO365ProPlusRetail not downloaded" -ForegroundColor Green
            $LogWindow.text += "$(timestamp):   $destinationO365ProPlusRetail download abgbrochen. `r`n"    
        }
    }





}


function DownloadAcrobatreaderDC{

 # Download the installer
    If (Test-Path -Path $destinationAcrobatreaderDC){ 
    Write-Host "$destinationAcrobatreaderDC exists" -ForegroundColor Green
    $LogWindow.text += "$(timestamp):   $destinationAcrobatreaderDC inst vorhanden. `r`n"
    }
    ELSE{ 

    # Download the installer
    $JobDownloadAcrobatreaderDC = Start-BitsTransfer -Source $sourceAcrobatreaderDC -Destination $destinationAcrobatreaderDC -Asynchronous
    $LogWindow.text += "$(timestamp):   $destinationAcrobatreaderDC wird herunter geladen. `r`n"

    While( ($JobDownloadAcrobatreaderDC.JobState.ToString() -eq 'Transferring') -or ($JobDownloadAcrobatreaderDC.JobState.ToString() -eq 'Connecting') ){

    $ProgressBar1.Value = [int](($JobDownloadAcrobatreaderDC.BytesTransferred*100) / $JobDownloadAcrobatreaderDC.BytesTotal)
    }
    Complete-BitsTransfer -BitsJob $JobDownloadAcrobatreaderDC
    

        If (Test-Path -Path $destinationAcrobatreaderDC){
                #Invoke-WebRequest $sourceAcrobatreaderDC -OutFile $destinationAcrobatreaderDC
            Write-Host "$destinationAcrobatreaderDC downloaded" -ForegroundColor Green
            $LogWindow.text += "$(timestamp):   $destinationAcrobatreaderDC download abgeschlossen. `r`n"
        }
        else {
            Write-Host "$destinationAcrobatreaderDC not downloaded" -ForegroundColor Green
            $LogWindow.text += "$(timestamp):   $destinationAcrobatreaderDC download abgbrochen. `r`n"    
        }
    }


}


function CheckBRZclient{


If (Test-Path -Path $destinationBRZclient){ 
    Write-Host "$destinationBRZclient exists" -ForegroundColor Green
    $LogWindow.text += "$(timestamp):   $destinationBRZclient inst vorhanden. `r`n"
}
ELSE{ 

    # Download the installer

    Write-Host "$destinationBRZclient doesn´t exists" -ForegroundColor Red
    $LogWindow.text += "$(timestamp):   $destinationBRZclient nicht vorhanden. `r`n"
}
}


function DownloadGoogleEarth{

If (Test-Path -Path $destinationGoogleEarth){ 
    Write-Host "$destinationGoogleEarth exists" -ForegroundColor Green
    $LogWindow.text += "$(timestamp):   $destinationGoogleEarth inst vorhanden. `r`n"
    }
    ELSE{ 

    # Download the installer
    $JobDownloadGoogleEarth = Start-BitsTransfer -Source $sourceGoogleEarth -Destination $destinationGoogleEarth -Asynchronous
    $LogWindow.text += "$(timestamp):   $destinationGoogleEarth wird herunter geladen. `r`n"

    While( ($JobDownloadGoogleEarth.JobState.ToString() -eq 'Transferring') -or ($JobDownloadGoogleEarth.JobState.ToString() -eq 'Connecting') )
    {

    $ProgressBar1.Value = [int](($JobDownloadGoogleEarth.BytesTransferred*100) / $JobDownloadGoogleEarth.BytesTotal)
    }
    Complete-BitsTransfer -BitsJob $JobDownloadGoogleEarth
    

        If (Test-Path -Path $destinationGoogleEarth){
                #Invoke-WebRequest $sourceGoogleEarth -OutFile $destinationGoogleEarth
            Write-Host "$destinationGoogleEarth downloaded" -ForegroundColor Green
            $LogWindow.text += "$(timestamp):   $destinationGoogleEarth download abgeschlossen. `r`n"
        }
        else {
            Write-Host "$destinationGoogleEarth not downloaded" -ForegroundColor Green
            $LogWindow.text += "$(timestamp):   $destinationGoogleEarth download abgbrochen. `r`n"    
        }
    }





}


function DownloadGoogleChrome{

    If (Test-Path -Path $destinationGoogleChrome){ 
    Write-Host "$destinationGoogleChrome exists" -ForegroundColor Green
    $LogWindow.text += "$(timestamp):   $destinationGoogleChrome inst vorhanden. `r`n"
    }
    ELSE{ 

    # Download the installer
    $JobDownloadGoogleChrome = Start-BitsTransfer -Source $sourceGoogleChrome -Destination $destinationGoogleChrome -Asynchronous
    $LogWindow.text += "$(timestamp):   $destinationGoogleChrome wird herunter geladen. `r`n"

    While( ($JobDownloadGoogleChrome.JobState.ToString() -eq 'Transferring') -or ($JobDownloadGoogleChrome.JobState.ToString() -eq 'Connecting') )
    {
    $ProgressBar1.Value = [int](($JobDownloadGoogleChrome.BytesTransferred*100) / $JobDownloadGoogleChrome.BytesTotal)
    }
    Complete-BitsTransfer -BitsJob $JobDownloadGoogleChrome
    

        If (Test-Path -Path $destinationGoogleChrome){
                #Invoke-WebRequest $sourceGoogleChrome -OutFile $destinationGoogleChrome
            Write-Host "$destinationGoogleChrome downloaded" -ForegroundColor Green
            $LogWindow.text += "$(timestamp):   $destinationGoogleChrome download abgeschlossen. `r`n"
        }
        else {
            Write-Host "$destinationGoogleChrome not downloaded" -ForegroundColor Green
            $LogWindow.text += "$(timestamp):   $destinationGoogleChrome download abgbrochen. `r`n"    
        }
    }



}

function btnDownload { 

    Sources
    $LogWindow.text += " ------------------------   $LogWindowTitel   ------------------------ `r`n"
    $LogWindow.text += "$(timestamp)   : Vorraussetzungen pruefen............ `r`n"
    #### Testing Admin And Path
    
    
    If ((TestPath) -eq $false){
        #$path = $TextBoxPermInstallPath.text
        $LogWindow.text += "$(timestamp)   :FEHLER: $path  nicht vorhanden. `r`n"
    }
    Else{
        $LogWindow.text += "$(timestamp)   :   OK   : $path  vorhanden. `r`n"
    }
    $LogWindow.text += "$(timestamp)   : Vorraussetzungen erfüllt. `r`n"
    $LogWindow.text += "$(timestamp)   : Installation der Programme kann ausgeführt werden. `r`n"
    
    if ($CheckBoxTeamViewerQS.Checked){
        DownloadTeamViewerQS
    }
    if ($CheckBoxRenameComputer.Checked){
        #RenameComputer
    }
 
    if ($CheckBoxFirefox.Checked){
        DownloadFirefox
    }    

    if ($CheckBox7z.Checked){
        Download7z
    }
    if ($CheckBoxGreenshot.Checked){
        DownloadGreenshot
    }
    if ($CheckBoxCDburnerXP.Checked){
        DownloadCDburnerXP
    }

    if ($CheckBoxAcrobatReaderDC.Checked){
        DownloadAcrobatreaderDC
    }
    if ($CheckBoxBRZclient.Checked){
        
    }
    if ($CheckBoxGoogleEarth.Checked){
        DownloadGoogleEarth
    }
        if ($CheckBoxGoogleChrome.Checked){
        DownloadGoogleChrome
    }
        if ($CheckBoxO365ProPlusRetail.Checked){
        DownloadO365ProPlusRetail
    }

    
$LogWindow.text += "$(timestamp): Download der Programme abgeschlossen. `r`n"

$ProgressBar1.Value = 100
}
#
#
#
#
#
#
#########################End of only Download
#
#
#
#
#
#

#
#
#
#
#
#
#########################Start of Install and Download
#
#
#
#
#
#

function CopyTeamViewerQS{

# Download the installer

DownloadTeamViewerQS


# Start the installation

    $UserDesktop = "$env:HOMEDRIVE$env:HOMEPATH\Desktop\" 
    Copy-Item -Path "$destinationTeamViewerQS" -Destination "$UserDesktop"

# Wait XX Seconds for the installation to finish

Start-Sleep -s 2
    


}


#Erledigt
function RenameComputer {
    $delimiter = '-'
    $pwqusers = quser | Select-Object -skip 1
    $pwGetDate = Get-Date -Format "yyyy"
    $pwquserfmt = $pwqusers.substring(1,22).Trim() -join ","
    $pwqyearfmt = $pwGetDate.substring(2,2).Trim() -join ","
    $NewName = $pwquserfmt + $delimiter + $pwqyearfmt
    IF($pwquserfmt -eq "Administrator"){
    $LogWindow.text += "$(timestamp):   $NewName ist kein gültiger Name `r`n"
    }
    Else{
    Rename-Computer -NewName $NewName
    }
}


function SilentInstallFirefox {

    # Download the installer


DownloadFirefox


        # Start the installation
        
$TestPathFirefox = Test-Path 'HKLM:\SOFTWARE\Mozilla\Firefox'

If($TestPathFirefox -eq "true"){
    Write-Host "$destinationFirefox ist schon installiert" -ForegroundColor Green
    $LogWindow.text += "$(timestamp):   $destinationFirefox ist schon installiert. `r`n"
}

Else{
    $LogWindow.text += "$(timestamp):   $destinationFirefox installation wird gestartet. `r`n" 
    Start-Process -FilePath "$destinationFirefox" -ArgumentList "-ms"
    Start-Sleep -s 15
    $LogWindow.text += "$(timestamp):   $destinationFirefox installation abgeschlossen. `r`n" 
}

     
    

    
}

function SilentInstall7z {

    # Download the installer


Download7z


        # Start the installation
        
$TestPath7z = Test-Path 'HKLM:\SOFTWARE\7-Zip'

If($TestPath7z -eq "true"){
    Write-Host "$destination7z ist schon installiert" -ForegroundColor Green
    $LogWindow.text += "$(timestamp):   $destination7z ist schon installiert. `r`n"
}

Else{
    $LogWindow.text += "$(timestamp):   $destination7z installation wird gestartet. `r`n" 
    msiexec.exe /i $destination7z /qn 
    Start-Sleep -s 5
    $LogWindow.text += "$(timestamp):   $destination7z installation abgeschlossen. `r`n" 
}

     
    

    
}


#Erledigt
function SilentInstallGreenshot{

    # Download the installer

DownloadGreenshot    


# Start the installation


$TestPathGreenshot = Test-Path "C:\Program Files\Greenshot\Greenshot.exe"

If($TestPathGreenshot -eq "True"){
    Write-Host "$destinationGreenshot ist schon installiert" -ForegroundColor Green
    $LogWindow.text += "$(timestamp):   $destinationGreenshot ist schon installiert. `r`n"
}

Else{
    $LogWindow.text += "$(timestamp):   $destinationGreenshot installation wird gestartet. `r`n" 
    Start-Process -FilePath "$destinationGreenshot" -ArgumentList "/VERYSILENT /SUPPRESSMESSAGEBOXES /CLOSEAPPLICATIONS /RESTARTAPPLICATIONS /NORESTART"
    Start-Sleep -s 35
    $LogWindow.text += "$(timestamp):   $destinationGreenshot installation abgeschlossen. `r`n" 
}

    


}


function SilentInstallCDburnerXP{

    # Download the installer

DownloadCDburnerXP

# Start the installation


$TestPathCDburnerXP = Test-Path 'HKLM:\SOFTWARE\WOW6432Node\Canneverbe Limited\CDBurnerXP'

If($TestPathCDburnerXP -eq "True"){
    Write-Host "$destinationCDburnerXP ist schon installiert" -ForegroundColor Green
    $LogWindow.text += "$(timestamp):   $destinationCDburnerXP ist schon installiert. `r`n"
}

Else{
    $LogWindow.text += "$(timestamp):   $destinationCDburnerXP installation wird gestartet. `r`n" 
    Start-Process -FilePath "$destinationCDburnerXP" -ArgumentList "/VERYSILENT" 
    Start-Sleep -s 35
    $LogWindow.text += "$(timestamp):   $destinationCDburnerXP installation abgeschlossen. `r`n" 
}




    

}

#Erledigt
function SilentInstallO365ProPlusRetail{

DownloadO365ProPlusRetail

# Start the installation



$TestPathO365ProPlusRetailx86 = Test-Path 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office\16.0\Word'
$TestPathO365ProPlusRetailx64 = Test-Path 'HKLM:\SOFTWARE\Microsoft\Office\16.0\Word'

If($TestPathO365ProPlusRetailx86 -eq "True"){
    Write-Host "$destinationO365ProPlusRetail x86 ist schon installiert" -ForegroundColor Green
    $LogWindow.text += "$(timestamp):   $destinationO365ProPlusRetail x86 ist schon installiert. `r`n"
}
ElseIf($TestPathO365ProPlusRetailx64 -eq "True"){
    Write-Host "$destinationO365ProPlusRetail x64 ist schon installiert" -ForegroundColor Green
    $LogWindow.text += "$(timestamp):   $destinationO365ProPlusRetail x86 ist schon installiert. `r`n"

}
Else{
    $LogWindow.text += "$(timestamp):   $destinationO365ProPlusRetail installation wird gestartet. `r`n" 
    Start-Process -FilePath "$destinationO365ProPlusRetail" #-ArgumentList "/configure $destinationXML4O365ProPlusRetail"  
    Start-Sleep -s 350
    $LogWindow.text += "$(timestamp):   $destinationO365ProPlusRetail installation abgeschlossen. `r`n" 
}






}


function SilentInstallAcrobatreaderDC{
# Silent install Adobe Reader DC

DownloadAcrobatreaderDC

# Start the installation


$TestPathAdobeDC = Test-Path 'HKLM:\SOFTWARE\WOW6432Node\Adobe\Acrobat Reader\DC'

If($TestPathAdobeDC -eq "True"){
    Write-Host "$destinationAcrobatreaderDC ist schon installiert" -ForegroundColor Green
    $LogWindow.text += "$(timestamp):   $destinationAcrobatreaderDC ist schon installiert. `r`n"
}
Else{
    $LogWindow.text += "$(timestamp):   $destinationAcrobatreaderDC installation wird gestartet. `r`n" 
    Start-Process -FilePath "$destinationAcrobatreaderDC" -ArgumentList "/sPB /rs"
    
    Start-Sleep -s 35
    $LogWindow.text += "$(timestamp):   $destinationAcrobatreaderDC installation abgeschlossen. `r`n" 
}


}


function SilentInstallBRZclient{


CheckBRZclient

# Start the installation
    if($destinationBRZclientExists = "notexist"){
        $LogWindow.text += "$(timestamp):   Bitte Pfad zu $destinationBRZclientExists kontrollieren. `r`n"
    }
    else {
        Start-Process -FilePath "$destinationBRZclient" -ArgumentList " --auto=TRUE -L0x0407"  

        # Wait XX Seconds for the installation to finish
        
        Start-Sleep -s 35
        
    }


}


function SilentInstallGoogleEarth{

DownloadGoogleEarth
# Start the installation

$TestPathGoogleEarth = Test-Path 'HKLM:\SOFTWARE\Google\Google Earth Pro'

If($TestPathGoogleEarth -eq "True"){
    Write-Host "$destinationGoogleEarth ist schon installiert" -ForegroundColor Green
    $LogWindow.text += "$(timestamp):   $destinationGoogleEarth ist schon installiert. `r`n"
}
Else{
    $LogWindow.text += "$(timestamp):   $destinationGoogleEarth installation wird gestartet. `r`n" 
    Start-Process -FilePath "$destinationGoogleEarth" -ArgumentList "OMAHA=1"  
    
    Start-Sleep -s 35
    $LogWindow.text += "$(timestamp):   $destinationGoogleEarth installation abgeschlossen. `r`n" 
}




}


function SilentInstallGoogleChrome{


DownloadGoogleChrome

# Start the installation
$TestPathGoogleChrome = Test-Path 'HKLM:\SOFTWARE\Google\Chrome'

If($TestPathGoogleChrome -eq "True"){
    Write-Host "$destinationGoogleChrome ist schon installiert" -ForegroundColor Green
    $LogWindow.text += "$(timestamp):   $destinationGoogleChrome ist schon installiert. `r`n"
}
Else{
    $LogWindow.text += "$(timestamp):   $destinationGoogleChrome installation wird gestartet. `r`n" 
    msiexec.exe /i "$destinationGoogleChrome"  /q /norestart
    
    Start-Sleep -s 35
    $LogWindow.text += "$(timestamp):   $destinationGoogleChrome installation abgeschlossen. `r`n" 
}



}




###### On Apply Button ######
function btnStart { 

    Sources
    $LogWindow.text += " ------------------------   $LogWindowTitel   ------------------------ `r`n"
    $LogWindow.text += "$(timestamp)   : Vorraussetzungen pruefen............ `r`n"
    #### Testing Admin And Path
    if ((Test-Admin) -eq $false) {
        $LogWindow.text += "$(timestamp)   :FEHLER: Pfad nich verfügbar. `r`n"
        return
    }
    Else{
        $LogWindow.text += "$(timestamp)   :    OK    : Pfad vorhanden. `r`n"
    }
    
    If ((TestPath) -eq $false){
        #$path = $TextBoxPermInstallPath.text
        $LogWindow.text += "$(timestamp)   :FEHLER: $path  nicht vorhanden. `r`n"
        return
    }
    Else{
        $LogWindow.text += "$(timestamp)   :   OK   : $path  vorhanden. `r`n"
    }
    $LogWindow.text += "$(timestamp)   : Vorraussetzungen erfüllt. `r`n"
    $LogWindow.text += "$(timestamp)   : Installation der Programme kann ausgeführt werden. `r`n"
    
    if ($CheckBoxTeamViewerQS.Checked){
        CopyTeamViewerQS
    }
    if ($CheckBoxRenameComputer.Checked){
        RenameComputer
    }

    if ($CheckBoxFirefox.Checked){
        SilentInstallFirefox
    }

    if ($CheckBox7z.Checked){
        SilentInstall7z
    }
    if ($CheckBoxGreenshot.Checked){
        SilentInstallGreenshot
    }
    if ($CheckBoxCDburnerXP.Checked){
        SilentInstallCDburnerXP
    }

    if ($CheckBoxAcrobatReaderDC.Checked){
        SilentInstallAcrobatreaderDC
    }
    if ($CheckBoxBRZclient.Checked){
        SilentInstallBRZclient
    }
    if ($CheckBoxGoogleEarth.Checked){
        SilentInstallGoogleEarth
    }
        if ($CheckBoxGoogleChrome.Checked){
        SilentInstallGoogleChrome
    }
        if ($CheckBoxO365ProPlusRetail.Checked){
        SilentInstallO365ProPlusRetail
    }

    
$LogWindow.text += "$(timestamp): Installation der Programme abgeschlossen. `r`n"
$ProgressBar1.Value = 100
}
#
#
#
#
#
#
#########################End of Install and Download
#
#
#
#
#
#



##### in das Fenster einbinden

$form1.controls.AddRange(@($ProgressBar1,$CheckBoxTeamViewerQS,$CheckBoxRenameComputer,$CheckBoxFirefox,$CheckBox7z,$CheckBoxGreenshot,$CheckBoxCDburnerXP,$CheckBoxO365ProPlusRetail,$CheckBoxAcrobatReaderDC,$CheckBoxBRZclient,$CheckBoxGoogleEarth,$CheckBoxGoogleChrome,$TextBoxPermInstallPath,$TextBoxPermInstallPathBRZclient,$ButtonStart,$ButtonDownload,$LogWindow))

$ButtonStart.Add_Click({ btnStart })

$ButtonDownload.Add_Click({ btnDownload })

[void]$form1.ShowDialog()







