#######################################
# Configurable Variables
#--------------------------------------
$version = "2.0-alpha"
$ProgramName = "MS365-PSToolkit"
########################################
#$tempdir = "C:\INSTALL\$ProgramName-$version"
#######################################
#######################################
# Fixed Variables
#--------------------------------------
$error.clear()
$ProgressPreference = 'SilentlyContinue'
$host.UI.RawUI.WindowTitle = "$ProgramName - Version $version"
#$curDate = Get-Date -Format "dd-MM-yyyy-HH-mm-ss"
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



function Show-MainMenu {
    Write-Host "#### MAIN MENU ####" -ForegroundColor Cyan
    Write-Host "`nSelect Microsoft 365 Service To Connect With"
    Write-Host "`n(1) - Microsoft Graph"
    Write-Host "(2) - Exchange Online"
    Write-Host "(3) - Sharepoint Online"
    Write-Host "`n(9) - Exit"
    Write-Host ""
}

function Show-MsGraphMenu {
    Write-Host "#### Microsoft Graph Menu ####" -ForegroundColor Cyan
    Write-Host "(1) - Install Required Modules"
    Write-Host "(2) - Check required permission scopes for MsGraph Command"
    Write-Host "(3) - Tool - Forced password change at next login - all users"
    Write-Host "(4) - Tool - Forced password change at next login - single user"
    Write-Host "(5) - Report - Create MFA status report"
    Write-Host "`n(9) - Main Menu"
    Write-Host ""
}

function Show-ExchangeOnlineMenu {
    Write-Host "#### Microsoft Exchange Online Menu ####" -ForegroundColor Cyan
    Write-Host "`n(1) - List Mailbox"
    Write-Host "(2) - List Mailbox Permissions"
    Write-Host "(3) - List Specific Mailbox Permissions"
    Write-Host "(4) - List Mailbox Size"
    Write-Host "(8) - Custom Command"
    Write-Host "`n(9) - Main Menu"
    Write-Host ""
}


function installExchangeOnlineModule {
    if(-not (Get-Module ExchangeOnlineManagement -ListAvailable)){
        Write-Warning "Module ExchangeOnlineManagement not installed"
        Write-Host "Installing ExchangeOnlineManagement Please Wait" -ForegroundColor Green
        Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force 
        }
    Start-Sleep -Seconds 1
    Write-Host "Module Already Installed" -ForegroundColor Green
    Start-Sleep -Seconds 2
}

function installMicrosoftGraphModule {
    if(-not (Get-Module Microsoft.Graph -ListAvailable)){
        Write-Warning "Module Microsoft.Graph not installed"
        Write-Host "Installing Microsoft.Graph Please Wait" -ForegroundColor Green
        Install-Module Microsoft.Graph -Scope CurrentUser -Force
        Install-Module Microsoft.Graph.Beta -Scope CurrentUser -Force
        }
    Start-Sleep -Seconds 1
    Write-Host "Module Already 
    Installed" -ForegroundColor Green
    Start-Sleep -Seconds 2
}





####################################
function CheckMsGraphDelegatedPermissions {
    $exit = 0
    while ($exit -eq 0){
        Write-Host "Enter command to see the required permission scopes"
        $cmd = Read-Host "Command :"
        if ($cmd -eq "exit"){
            $exit = 1 
        }
        else {
            Find-MgGraphCommand -Command "$cmd" | Select-Object -ExpandProperty Permissions | Select-Object -Unique name
        }
    }
}

function MsGraphForcePasswordResetAllUsers {
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

    #Disconnect from the Microsoft Graph If already connected
    if (Get-MgContext) {
        Write-Host Disconnecting from the previous sesssion.... -ForegroundColor Yellow
        Disconnect-MgGraph | Out-Null
    }
    
    Write-Host "This script will force every account to change password at next sign in." -ForegroundColor Yellow
    Write-Host "`nContinue ? [y/n]"
    $continue = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    if ($continue.Character -eq "n"){
        Exit
    }
    
    # Connect to Microsoft Graph API
    Connect-MgGraph -Scopes "User.Read.All", "User-PasswordProfile.ReadWrite.All", "UserAuthenticationMethod.ReadWrite.All" -NoWelcome
    
    Write-Host "Would you like to exclude an account from this script? [y/n]"
    $exclude = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    if ($exclude.Character -eq "y"){
        Write-Host "Enter UPN of account to exclude (example@domain.com)"
        $excludedUPN = Read-Host "UPN :"
        $excludedAccount = Get-MgBetaUser -UserId $excludedUPN
    }
    
    
    # Get all Microsoft Entra ID users using the Microsoft Graph Beta API
    $users = Get-MgBetaUser -All
    
    # Initialize progress counter
    $counter = 0
    $totalUsers = $users.Count
    
    # Loop through each user account
    foreach ($user in $users) {
        if ($user.Id -eq $excludedAccount.Id){continue} 
        $counter++
    
        # Calculate percentage completion
        $percentComplete = [math]::Round(($counter / $totalUsers) * 100)
    
        # Define progress bar parameters with user principal name
        $progressParams = @{
            Activity        = "Processing Users"
            Status          = "User $($counter) of $totalUsers - $($user.UserPrincipalName) - $percentComplete% Complete"
            PercentComplete = $percentComplete
        }
    
        Write-Progress @progressParams
        try {
            Update-MgUser -UserId $user.id -PasswordProfile @{ ForceChangePasswordNextSignIn=$true}
        }
        catch {
            Write-Warning "Error updating $($User.UserPrincipalName)"
        }
        
    }
    # Clear progress bar
    Write-Progress -Activity "Processing Users" -Completed
    
}

####################################








function isExchangeConnected {
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


function customExchangeCmd {
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





#####################################################################
# Main Code --- Main Code --- Main Code --- Main Code --- Main Code #
#-------------------------------------------------------------------#
Logo
Write-Host
Write-Host "Microsoft 365 Powershell Toolkit" -ForegroundColor Yellow
Write-Host "`nMATRIXNET ~ Vincent" -ForegroundColor Yellow
Write-Host
Write-Host "----------------------------" -ForegroundColor Magenta
write-Host "| Always trust the process |" -ForegroundColor Magenta
Write-Host "----------------------------" -ForegroundColor Magenta
Start-Sleep -Seconds 3
Set-ExecutionPolicy Bypass -Force -Scope Process -Confirm:$false
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






# Begin Loop
$WhileLoopVarMainMenu = 1
while ($WhileLoopVarMainMenu -eq 1){
    Clear-Host
    Start-Sleep -Milliseconds 500
    Show-MainMenu
    $MainMenuChoice = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown").VirtualKeyCode
    switch ($MainMenuChoice) {
        49 {
            $WhileLoopVarMsGraphMenu = 1
            while ($WhileLoopVarMsGraphMenu -eq 1){
                Clear-Host
                Start-Sleep -Milliseconds 500
                Show-MsGraphMenu
                $MsGraphMenuChoice = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown").VirtualKeyCode
                    switch ($MsGraphMenuChoice) {
                        49 {installMicrosoftGraphModule}
                        50 {CheckMsGraphDelegatedPermissions}
                        51 {MsGraphForcePasswordResetAllUsers}
                        57 {
                            Write-Host "`nReturning to Main Menu" -ForegroundColor Cyan
                            $WhileLoopVarMsGraphMenu = 0
                        }
                        default {
                            Write-Host "`nInvalid selection, please try again." -ForegroundColor Red
                            Start-Sleep -Seconds 1
                        }
                    }
            }
        }


        50 {
            $WhileLoopVarExchangeOnlineMenu = 1
            while ($WhileLoopVarExchangeOnlineMenu -eq 1){
                Clear-Host
                Start-Sleep -Milliseconds 500
                Show-MsGraphMenu
                $MsGraphMenuChoice = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown").VirtualKeyCode
                    switch ($MsGraphMenuChoice) {
                        49 {installExchangeOnlineModule}
                        50 {}
                        51 {}
                        57 {
                            Write-Host "`nReturning to Main Menu" -ForegroundColor Cyan
                            $WhileLoopVarExchangeOnlineMenu = 0
                        }
                        default {
                            Write-Host "`nInvalid selection, please try again." -ForegroundColor Red
                            Start-Sleep -Seconds 1
                        }
                    }
            }
        }
        51 { # ASCII for "3"
            Write-Host "3"
        }
        57 { # ASCII for "9"
            Write-Host "`nExiting... Goodbye!" -ForegroundColor Cyan
            $WhileLoopVarMainMenu = 0
            exit
        }
        default {
            Write-Host "`nInvalid selection, please try again." -ForegroundColor Red
            Start-Sleep -Seconds 1
        }
    }

    if ($MainMenuChoice -ne 57) {
        Write-Host ""
        Start-Sleep -Seconds 2
    }
}
exit