#######################################
# Configurable Variables
#--------------------------------------
$version = "1.0-beta.1"
$ProgramName = "MS365-PSToolkit"
########################################
$tempdir = "C:\INSTALL\$ProgramName"
#######################################
#######################################
# Fixed Variables
#--------------------------------------
$error.clear()
$ProgressPreference = 'SilentlyContinue'
$host.UI.RawUI.WindowTitle = "$ProgramName - Version $version"
$curDate = Get-Date -Format "dd-MM-yyyy-HH-mm-ss"
[console]::WindowWidth=200; [console]::WindowHeight=50; [console]::BufferWidth=[console]::WindowWidth
#######################################



function Logo {
    Write-Host " "
    Write-Host "  __  __ ___ ____  __ ___     ___  ___ _____         _ _   _ _   " -ForegroundColor Blue
    write-host " |  \/  / __|__ / / /| __|___| _ \/ __|_   _|__  ___| | |_(_) |_ " -ForegroundColor Blue
    Write-Host " | |\/| \__ \|_ \/ _ \__ \___|  _/\__ \ | |/ _ \/ _ \ | / / |  _|" -ForegroundColor Blue
    Write-Host " |_|  |_|___/___/\___/___/   |_|  |___/ |_|\___/\___/_|_\_\_|\__|" -ForegroundColor Blue
    Write-Host ""
    write-Host "v$version" -ForegroundColor Blue
}

function saveCSV {
    Write-Host "Save output to CSV ?  (Y/N)" -ForegroundColor Yellow
    Write-Host "Save CSV ? >>> " -NoNewline -ForegroundColor Yellow
    $readSaveCSV = Read-Host
    if ($readSaveCSV -eq "Y") {
        Write-Host "`nExporting Data to CSV" -ForegroundColor Yellow
        Invoke-Expression $functionCmd | Export-Csv -LiteralPath $tempdir\"$ProgramName-$functionName-$curDate.csv"
        Write-Host "Export Complete" -ForegroundColor Yellow
        Write-Host '`nPress any key to return to menu' -NoNewline -ForegroundColor Yellow
        $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
    }
}

function installModule {
    Write-Host "Checking if ExchangeOnlineManagement Module is installed" -ForegroundColor Yellow
    if(-not (Get-Module ExchangeOnlineManagement -ListAvailable)){
        Write-Host "Installing Module" -ForegroundColor Yellow
        Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force
        }
    Start-Sleep -Seconds 1
    Write-Host "Module Installed" -ForegroundColor Green
    Start-Sleep -Seconds 2
}

function isConnected {
    $Connection = Get-ConnectionInformation
    If (-Not ( $Connection.State -match 'Connected' ) ){
        Write-Host "Connecting to Microsoft 365 - Exchange Online" -ForegroundColor Yellow
        Connect-ExchangeOnline -ShowBanner:$False
    }
    else {
        Write-Host "Disconnecting Microsoft 365 - Exchange Online" -ForegroundColor Yellow
        Start-Sleep -Seconds 2
        Disconnect-ExchangeOnline -Confirm:$false
    }
}

function listMailbox {
    $functionName = "listMailbox" 
    $functionCmd = "Get-Mailbox -Resultsize Unlimited | Select-Object Name, PrimarySMTPAddress"
    Invoke-Expression "Get-Mailbox | ft name,PrimarySMTPAddress -autosize" -Debug
    Write-Host ""
    saveCSV
}

function listMailboxPermissions {
    $functionName = "listMailboxPermissions" 
    $functionCmd = "Get-mailbox | Get-MailboxPermission | Select-object Identity, User, AccessRights"
    Invoke-Expression "Get-mailbox | Get-MailboxPermission | ft Identity, User, AccessRights -autosize" -Debug
    Write-Host ""
    saveCSV
}

function listMailboxSize {
    $functionName = "listMailboxSize" 
    $functionCmd = "get-mailbox | Get-MailboxStatistics | Select-Object DisplayName, MailboxTypeDetail, ItemCount, TotalItemSize, SystemMessageSizeShutoffQuota"
    Invoke-Expression "get-mailbox | Get-MailboxStatistics | ft DisplayName, MailboxTypeDetail, ItemCount, TotalItemSize, SystemMessageSizeShutoffQuota -autosize" -Debug
    Write-Host ""
    saveCSV
}


function customCmd {
    $whileLoopVarCustomCmd = 1
    Logo
    while ($whileLoopVarCustomCmd -eq 1) {
        write-host ""
        Write-Host "Exchange Online Shell >>> " -NoNewline -ForegroundColor Yellow
        $readCustomCmd = Read-Host
        if ($readCustomCmd -eq "exit"){
            $whileLoopVarCustomCmd = 0
        }
        else {
            Invoke-Expression $readCustomCmd -Debug
            Write-Host ""
        }       
    }    
}


function UI_EXIT {
    Clear-Host
    Write-Host "BYE BYE !!!" -ForegroundColor Green
    Write-Host "`n############################################" -ForegroundColor Magenta
    Write-Host "Last Errors :" -ForegroundColor Yellow
    Write-host "$error" -ForegroundColor Red
    Write-Host "############################################" -ForegroundColor Magenta
    $curDateFinished = Get-Date -Format "dd/MM/yyyy HH/mm/ss"
    Write-Host "`nFinished : $curDateFinished" -ForegroundColor Yellow
    Write-Host " "
    Write-Host " "
    Write-Host "################ LOG END ################"-ForegroundColor Magenta
    Stop-Transcript
    Start-Sleep -Seconds 3
    Exit 0
    
}



#####################################################################
# Check Install Folder
if (-Not (Test-Path $tempdir)){
    New-Item -ItemType Directory -Path $tempdir
}
#####################################################################
#####################################################################                                
# Start Logging in Install directory (Tempdir)
$CurDate = Get-Date -Format "dd-MM-yyyy-HH-mm-ss"
Start-Transcript -Path "$tempdir\$ProgramName-$version-$CurDate.log" | Out-Null
Write-Host " "
Write-Host " "
Write-Host "################ LOG BEGIN ################" -ForegroundColor Magenta
Clear-Host
Start-Sleep -Milliseconds 250
#####################################################################

#####################################################################
# Main Code --- Main Code --- Main Code --- Main Code --- Main Code #
#-------------------------------------------------------------------#
Logo
Write-Host
Write-Host "Microsoft 365 Powershell Toolkit" -ForegroundColor Yellow
Write-Host "`nMATRIXNET ~ Vincent" -ForegroundColor Yellow
Write-Host "INCONEL BV ~ Vincent" -ForegroundColor Yellow
Write-Host
Write-Host "----------------------------" -ForegroundColor Magenta
write-Host "| Always trust the process |" -ForegroundColor Magenta
Write-Host "----------------------------" -ForegroundColor Magenta
Start-Sleep -Seconds 3
#####################################################################





#########################
# Check admin rights    #
#-----------------------#
function isadmin
 {
 # Returns true/false
   ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
 }
 if (isadmin -eq "True") {
     Write-Host "`nGot Administrator Permissions" -ForegroundColor Green
     Start-Sleep -Seconds 1
 }
 else {
     Write-Host "`nThis script needs Administrator Privileges to work it's magic" -ForegroundColor Red
     Start-Sleep -Seconds 3
 }
 if (-Not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] 'Administrator')) {
  if ([int](Get-CimInstance -Class Win32_OperatingSystem | Select-Object -ExpandProperty BuildNumber) -ge 6000) {
   $CommandLine = "-File `"" + $MyInvocation.MyCommand.Path + "`" " + $MyInvocation.UnboundArguments
   Start-Process -FilePath PowerShell.exe -Verb Runas -ArgumentList $CommandLine
   Exit
  }
}
#####################################################################





##################################
# Begin Loop
$WhileLoopVarMenu = 1
while ($WhileLoopVarMenu -eq 1){
    
    #Check Connection State
    $Connection = Get-ConnectionInformation
    If ( $Connection.State -match 'Connected' ){
        Write-Host "Connected to Microsoft 365 - Exchange Online" -ForegroundColor Green
        $connectionStatus = "Green"
    }
    else {
        $connectionStatus = "Red"
    }
#####################################################################
# Interactive Menu #
##################################
#Menu items
$list = @('Install Modules','Connect/Disconnect MS365','List Mailbox','List Mailbox Permissions', 'List Mailbox Size','Custom Command','EXIT')
 
#menu offset to allow space to write a message above the menu
$xmin = 3
$ymin = 15
 
#Write Menu
Clear-Host
Logo
Write-Host ""
Write-Host "Connection : " -NoNewline -ForegroundColor Yellow
Write-Host "(#)" -ForegroundColor $connectionStatus
Write-Host "`n"
Write-Host ""
Write-Host "  Use the up / down arrow to navigate and Enter to make a selection" -ForegroundColor Yellow
[Console]::SetCursorPosition(0, $ymin)
foreach ($name in $List) {
    for ($i = 0; $i -lt $xmin; $i++) {
        Write-Host " " -NoNewline
    }
    Write-Host "   " + $name
}
 
#Highlights the selected line
function Write-Highlighted {
 
    [Console]::SetCursorPosition(1 + $xmin, $cursorY + $ymin)
    Write-Host ">" -BackgroundColor Yellow -ForegroundColor Black -NoNewline
    Write-Host " " + $List[$cursorY] -BackgroundColor Yellow -ForegroundColor Black
    [Console]::SetCursorPosition(0, $cursorY + $ymin)     
}
 
#Undoes highlight
function Write-Normal {
    [Console]::SetCursorPosition(1 + $xmin, $cursorY + $ymin)
    Write-Host "  " + $List[$cursorY]  
}
 
#highlight first item by default
$cursorY = 0
Write-Highlighted
 
$selection = ""
$menu_active = $true
while ($menu_active) {
    if ([console]::KeyAvailable) {
        $x = $Host.UI.RawUI.ReadKey()
        [Console]::SetCursorPosition(1, $cursorY)
        Write-Normal
        switch ($x.VirtualKeyCode) { 
            38 {
                #down key
                if ($cursorY -gt 0) {
                    $cursorY = $cursorY - 1
                }
            }
 
            40 {
                #up key
                if ($cursorY -lt $List.Length - 1) {
                    $cursorY = $cursorY + 1
                }
            }
            13 {
                #enter key
                $selection = $List[$cursorY]
                $menu_active = $false
            }
        }
        Write-Highlighted
    }
    Start-Sleep -Milliseconds 5 #Prevents CPU usage from spiking while looping
}
 
Clear-Host
switch ($selection) {
    "Install Modules" {installModule}
    "Connect/Disconnect MS365" {isConnected}
    "List Mailbox" {listMailbox}
    "List Mailbox Permissions" {listMailboxPermissions}
    "List Mailbox Size" {listMailboxSize}
    "Custom Command" {customCmd}
    "EXIT" {UI_EXIT}
}
#####################################################################
}