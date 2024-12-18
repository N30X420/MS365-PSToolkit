#######################################
# Configurable Variables
#--------------------------------------
$version = "2.0.1"
$ProgramName = "MS365-PSToolkit"
$tempdir = "C:\INSTALL\$ProgramName-$version"
$GithubRepo = "https://github.com/N30X420/MS365-PSToolkit"
#######################################
#######################################
$error.clear()
$ProgressPreference = 'Continue'
$host.UI.RawUI.WindowTitle = "$ProgramName - Version $version"
[console]::WindowWidth=200; [console]::WindowHeight=50; [console]::BufferWidth=[console]::WindowWidth
$script:ExportCSV = $false
#######################################


function Logo {
    Write-Host " "
    Write-Host "  __  __ ___ ____  __ ___     ___  ___ _____         _ _   _ _   " -ForegroundColor Blue
    write-host " |  \/  / __|__ / / /| __|___| _ \/ __|_   _|__  ___| | |_(_) |_ " -ForegroundColor Blue
    Write-Host " | |\/| \__ \|_ \/ _ \__ \___|  _/\__ \ | |/ _ \/ _ \ | / / |  _|" -ForegroundColor Blue
    Write-Host " |_|  |_|___/___/\___/___/   |_|  |___/ |_|\___/\___/_|_\_\_|\__|" -ForegroundColor Blue
    Write-Host ""
    write-Host "v$version " -ForegroundColor Blue -NoNewline
    Write-Host " $Script:NewVersionAvailable" -ForegroundColor Green
    Write-Host
    Write-Host "Microsoft 365 Powershell Toolkit" -ForegroundColor Cyan
    Write-Host "MATRIXNET ~ Vincent" -ForegroundColor Cyan
    Write-Host ""
}
function Show-MainMenu {
    Write-Host "################## MAIN MENU ##################" -ForegroundColor DarkCyan
    Write-Host "Select Microsoft 365 Service To Connect With"
    Write-Host "`n(1) - Microsoft Graph"
    Write-Host "(2) - Exchange Online"
    Write-Host "(3) - Sharepoint Online - Coming Soon" -ForegroundColor Red
    Write-Host "`n(8) - Check For Updates"
    Write-Host "(9) - Exit"
    Write-Host "###############################################" -ForegroundColor DarkCyan
    Write-Host ""
}
function CheckPowershellVersion {
    if ($PSVersionTable.PSVersion.Major -ne 7){
        Write-Warning "This program requires powershell 7. You are running version $($PSVersionTable.PSVersion.Major).$($PSVersionTable.PSVersion.Minor) Please Update to Powershell 7.x"
        exit
    }
}
function CheckAdminPrivs {
    function isadmin {
        # Returns true/false
        ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
    }
        if (isadmin -eq "True") {
            Write-Host "`nGot Administrator Permissions" -ForegroundColor Green
            Start-Sleep -Seconds 1
        }
        else {
            Write-Error "This script needs Administrator Privileges to work it's magic"
            Start-Sleep -Seconds 3
        }
        if (-Not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] 'Administrator')) {
            if ([int](Get-CimInstance -Class Win32_OperatingSystem | Select-Object -ExpandProperty BuildNumber) -ge 6000) {
                $CommandLine = "-File `"" + $MyInvocation.MyCommand.Path + "`" " + $MyInvocation.UnboundArguments
                Start-Process -FilePath PowerShell.exe -Verb Runas -ArgumentList $CommandLine
                Exit
            }
        }
}
function CheckForUpdates {
    try {
        $Releases = Invoke-RestMethod -Uri "https://api.github.com/repos/N30X420/MS365-PSToolkit/releases"
		$ReleaseInfo = ($Releases | Sort-Object id -desc)[0]
		$LatestVersion = [version[]]$ReleaseInfo.Name.Trim('v')
		if ($LatestVersion -gt $version){ $Script:NewVersionAvailable = "v$LatestVersion is available $(Format-Hyperlink -Uri "$GithubRepo" -Label "v$LatestVersion")"}
    }
    catch {
        Write-Warning "Error while checking for updates"
        Start-Sleep -Seconds 2
    }
}
function Format-Hyperlink {
    param(
      [Parameter(ValueFromPipeline = $true, Position = 0)]
      [ValidateNotNullOrEmpty()]
      [Uri] $Uri,
  
      [Parameter(Mandatory=$false, Position = 1)]
      [string] $Label
    )
  
    if (($PSVersionTable.PSVersion.Major -lt 6 -or $IsWindows) -and -not $Env:WT_SESSION) {
      # Fallback for Windows users not inside Windows Terminal
      if ($Label) {
        return "($Uri)"
      }
      return "$Uri"
    }
  
    if ($Label) {
      return "`e]8;;$Uri`e\$Label`e]8;;`e\"
    }
  
    return "$Uri"
  }
function CatchError {
    Start-Sleep -Seconds 1
    Write-Host -NoNewLine "Press any key to return to the menu..." -ForegroundColor Yellow
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
}
function PromptExportToCSV {
    Write-Host "Export data to CSV ? [y/n]" -ForegroundColor Yellow
    $UserInput = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    If ($UserInput -eq "y"){
        Write-Host "Exporting Data" -ForegroundColor Yellow
        $script:ExportCSV = $true
    }
    else {
        Start-Sleep -Seconds 1
        Break
    }
}

####################################
#  MsGraph  #
#-------------
function Show-MsGraphMenu {
    Write-Host "############# Microsoft Graph Menu #############" -ForegroundColor DarkCyan
    Write-Host "(1) - Install Required Modules"
    Write-Host "(2) - Check required permission scopes for MsGraph Command"
    Write-Host "(3) - Tool - Forced password change at next login - all users"
    Write-Host "(4) - Tool - Forced password change at next login - single user"
    Write-Host "(5) - Report - Create MFA status report"
    Write-Host "`n(9) - Main Menu"
    Write-Host "################################################" -ForegroundColor DarkCyan
    Write-Host ""
}
function installMicrosoftGraphModule {
    if(-not (Get-Module Microsoft.Graph -ListAvailable)){
        Write-Warning "Module Microsoft.Graph not installed"
        Write-Host "Installing Microsoft.Graph Please Wait" -ForegroundColor Green
        Install-Module Microsoft.Graph -Scope CurrentUser -Force
        Install-Module Microsoft.Graph.Beta -Scope CurrentUser -Force
        }
    Start-Sleep -Seconds 1
    Write-Host "Module Already Installed" -ForegroundColor Green
    Start-Sleep -Seconds 2
}
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
    
    Write-Warning "This script will force every account to change password at next sign in."
    Write-Host "`nContinue ? [y/n]"
    $continue = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    if ($continue.Character -ine "y"){
        Break
    }
    
    # Connect to Microsoft Graph API
    Write-Host "`nA new browser window will open for you to sign in using your Microsoft 365 Global Admin Account" -ForegroundColor Yellow
    Start-Sleep -Seconds 2
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
function MsGraphForcePasswordResetSingleUser {
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    if (Get-MgContext) {
        Write-Host Disconnecting from the previous sesssion.... -ForegroundColor Yellow
        Disconnect-MgGraph | Out-Null
    }

    Write-Host "`nA new browser window will open for you to sign in using your Microsoft 365 Global Admin Account" -ForegroundColor Yellow
    Start-Sleep -Seconds 2
    
    Connect-MgGraph -Scopes "User.Read.All", "User-PasswordProfile.ReadWrite.All", "UserAuthenticationMethod.ReadWrite.All" -NoWelcome
    Write-Host "Enter UPN of account to configure forced password change at next login (example@domain.com)"
        $UPN = Read-Host "UPN :"
        $MSAccount = Get-MgBetaUser -UserId $UPN
    try {
        Update-MgUser -UserId $MSAccount.id -PasswordProfile @{ForceChangePasswordNextSignIn=$true}
        Write-Host "$($UMSAccount.UserPrincipalName) Configured" -ForegroundColor Green
        Start-Sleep -Seconds 2
    }
    catch {
        Write-Warning "Error updating $($UMSAccount.UserPrincipalName)"
    }
    Disconnect-MgGraph | Out-Null
}
function CreateMFAStatusReport {
    #Disconnect from the Microsoft Graph If already connected
    if (Get-MgContext) {
        Write-Host Disconnecting from the previous sesssion.... -ForegroundColor Yellow
        Disconnect-MgGraph | Out-Null
    }

    Write-Host "`nA new browser window will open for you to sign in using your Microsoft 365 Global Admin Account" -ForegroundColor Yellow
    Start-Sleep -Seconds 2

    # Connect to Microsoft Graph API
    Connect-MgGraph -Scopes "User.Read.All", "UserAuthenticationMethod.Read.All", "UserAuthenticationMethod.ReadWrite.All" -NoWelcome

    # Create variable for the date stamp
    $LogDate = Get-Date -f yyyyMMddhhmm

    # Define CSV file export location variable
    $Csvfile = "$tempdir\MFAUsers_$LogDate.csv"

    # Get all Microsoft Entra ID users using the Microsoft Graph Beta API
    $users = Get-MgBetaUser -All

    # Initialize a List to store the data
    $Report = [System.Collections.Generic.List[Object]]::new()

    # Initialize progress counter
    $counter = 0
    $totalUsers = $users.Count

    # Loop through each user account
    foreach ($user in $users) {
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

        # Create an object to store user MFA information
        $ReportLine = [PSCustomObject]@{
            DisplayName               = "-"
            UserPrincipalName         = "-"
            MFAstatus                 = "Disabled"
            DefaultMFAMethod          = "-"
            Email                     = "-"
            Fido2                     = "-"
            MicrosoftAuthenticatorApp = "-"
            Phone                     = "-"
            SoftwareOath              = "-"
            TemporaryAccessPass       = "-"
            WindowsHelloForBusiness   = "-"
        }

        $ReportLine.UserPrincipalName = $user.UserPrincipalName
        $ReportLine.DisplayName = $user.DisplayName

        # Check authentication methods for each user
        $MFAData = Get-MgBetaUserAuthenticationMethod -UserId $user.Id

        # Retrieve the default MFA method
        $DefaultMFAUri = "https://graph.microsoft.com/beta/users/$($user.Id)/authentication/signInPreferences"
        try {
            $DefaultMFAMethod = Invoke-MgGraphRequest -Uri $DefaultMFAUri -Method GET
            if ($DefaultMFAMethod.userPreferredMethodForSecondaryAuthentication) {
                $ReportLine.DefaultMFAMethod = $DefaultMFAMethod.userPreferredMethodForSecondaryAuthentication
            }
            else {
                $ReportLine.DefaultMFAMethod = "Not set"
            }
        }
        catch {
            Write-Warning "Failed to retrieve default MFA method for $($user.UserPrincipalName): $_"
            $ReportLine.DefaultMFAMethod = "Error"
        }

        foreach ($method in $MFAData) {

            Switch ($method.AdditionalProperties["@odata.type"]) {
                "#microsoft.graph.emailAuthenticationMethod" {
                    $ReportLine.Email = $true
                    $ReportLine.MFAstatus = "Enabled"
                }
                "#microsoft.graph.fido2AuthenticationMethod" {
                    $ReportLine.Fido2 = $true
                    $ReportLine.MFAstatus = "Enabled"
                }
                "#microsoft.graph.microsoftAuthenticatorAuthenticationMethod" {
                    $ReportLine.MicrosoftAuthenticatorApp = $true
                    $ReportLine.MFAstatus = "Enabled"
                }
                "#microsoft.graph.phoneAuthenticationMethod" {
                    $ReportLine.Phone = $true
                    $ReportLine.MFAstatus = "Enabled"
                }
                "#microsoft.graph.softwareOathAuthenticationMethod" {
                    $ReportLine.SoftwareOath = $true
                    $ReportLine.MFAstatus = "Enabled"
                }
                "#microsoft.graph.temporaryAccessPassAuthenticationMethod" {
                    $ReportLine.TemporaryAccessPass = $true
                    $ReportLine.MFAstatus = "Enabled"
                }
                "#microsoft.graph.windowsHelloForBusinessAuthenticationMethod" {
                    $ReportLine.WindowsHelloForBusiness = $true
                    $ReportLine.MFAstatus = "Enabled"
                }
            }
        }
        # Add the report line to the List
        $Report.Add($ReportLine)
    }

    # Clear progress bar
    Write-Progress -Activity "Processing Users" -Completed

    # Export user information to CSV
    $Report | Export-Csv -Path $Csvfile -NoTypeInformation -Encoding UTF8

    Write-Host "Script completed. Results exported to $Csvfile." -ForegroundColor Cyan
    Start-Process $Csvfile
    
    Disconnect-MgGraph | Out-Null

    Write-Host -NoNewLine "Press any key to return to the menu..." -ForegroundColor Yellow
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
    
}
####################################

####################################
#  Exchange Online  #
#-----------------------
function Show-ExchangeOnlineMenu {
    Write-Host "######## Microsoft Exchange Online Menu ########" -ForegroundColor DarkCyan
    Write-Host "Exchange Online Connection : " -NoNewline -ForegroundColor Yellow
    Write-Host "$connectionStatus" -ForegroundColor $connectionStatusColor
    Write-Host "`n(1) - Install Required Modules"
    Write-Host "(2) - Connect / Disconnect Exchange Online"
    Write-Host "(3) - List Mailbox - Coming Soon" -ForegroundColor Red
    Write-Host "(4) - List Mailbox Permissions - Coming Soon" -ForegroundColor Red
    Write-Host "(5) - List Specific Mailbox Permissions - Coming Soon" -ForegroundColor Red
    Write-Host "(6) - List Mailbox Size - Coming Soon" -ForegroundColor Red
    Write-Host "(7) - Open Tools Menu"
    Write-Host "(8) - Custom Command"
    Write-Host "`n(9) - Main Menu"
    Write-Host "################################################" -ForegroundColor DarkCyan
    Write-Host ""
}
function Show-ExchangeOnlineToolsMenu {
    Write-Host "######## Microsoft Exchange Online Tools Menu ########" -ForegroundColor DarkCyan
    Write-Host "Exchange Online Connection : " -NoNewline -ForegroundColor Yellow
    Write-Host "$connectionStatus" -ForegroundColor $connectionStatusColor
    Write-Host "`n(1) - Enable Delegate Sent Items Shared Mailbox"
    Write-Host "`n(9) - Return to Exchange Online Menu"
    Write-Host "#####################################################" -ForegroundColor DarkCyan
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
function ExchangeOnlineConnection {
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

function ExchangeOnlineListMailbox {
    Get-Recipient -ResultSize Unlimited -RecipientTypeDetails UserMailbox,SharedMailbox | select-object Displayname,RecipientType,PrimarySMTPAddress,EmailAddresses
    PromptExportToCSV
    if ($script:ExportCSV -eq $true){
        $RecipientType={'RoomMailbox', 'LinkedRoomMailbox', 'EquipmentMailbox', 'SchedulingMailbox', 
        'LegacyMailbox', 'LinkedMailbox', 'UserMailbox', 'MailContact', 'DynamicDistributionGroup', 'MailForestContact', 'MailNonUniversalGroup', 'MailUniversalDistributionGroup', 'MailUniversalSecurityGroup', 
        'RoomList', 'MailUser', 'GuestMailUser', 'GroupMailbox', 'DiscoveryMailbox', 'PublicFolder', 'TeamMailbox', 'SharedMailbox', 'RemoteUserMailbox', 'RemoteRoomMailbox', 'RemoteEquipmentMailbox', 
        'RemoteTeamMailbox', 'RemoteSharedMailbox', 'PublicFolderMailbox', 'SharedWithMailUser'}

        $ExportResult=""   
        $ExportResults=@()  
        $OutputCSV="$tempdir\Office365EmailAddressesReport_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv"
        $Count=0
        #Get all Email addresses in Microsoft 365
        Get-Recipient -ResultSize Unlimited -RecipientTypeDetails $RecipientType | ForEach-Object {
            $Count++
            $DisplayName=$_.DisplayName
            Write-Progress -Activity "`n     Retrieving email addresses of $DisplayName.."`n" Processed count: $Count"
            $RecipientTypeDetails=$_.RecipientTypeDetails
            $PrimarySMTPAddress=$_.PrimarySMTPAddress
            $Alias= ($_.EmailAddresses | Where-Object {$_ -clike "smtp:*"} | ForEach-Object {$_ -replace "smtp:",""}) -join ","
            If($Alias -eq "")
            {
            $Alias="-"
            }

            #Export result to CSV file
            $ExportResult=@{'Display Name'=$DisplayName;'Recipient Type Details'=$RecipientTypeDetails;'Primary SMTP Address'=$PrimarySMTPAddress;'Alias'=$Alias}
            $ExportResults= New-Object PSObject -Property $ExportResult  
            $ExportResults | Select-Object 'Display Name','Recipient Type Details','Primary SMTP Address','Alias' | Export-Csv -Path $OutputCSV -Notype -Append
        }
    }
    Write-Host "Script completed. Results exported to  $OutputCSV." -ForegroundColor Cyan
}
function ExchangeOnlineSaveSentItemsInSharedMailboxAll {
    $mailboxes = (get-mailbox -ResultSize Unlimited -RecipientTypeDetails Sharedmailbox)
    $counter = 0
    $totalmailboxes = $mailboxes.Count
    foreach ($mailbox in $mailboxes){ 
        $counter++
        $percentComplete = [math]::Round(($counter / $totalmailboxes) * 100)
        $progressParams = @{
            Activity        = "Processing Mailbox"
            Status          = "Mailbox $($counter) of $totalmailboxes - $($Mailbox.Name) - $percentComplete% Complete"
            PercentComplete = $percentComplete
        }
        Write-Progress @progressParams
        Set-mailbox -Identity $mailbox.Identity -MessageCopyForSendOnBehalfEnabled $true -MessageCopyForSentAsEnabled $true
    } 
    Write-Progress -Activity "Processing Mailboxes" -Completed
        
}
function customExchangeCmd {
    $whileLoopVarCustomCmd = 1
    Clear-Host
    Start-Sleep -Milliseconds 250
    Logo
    while ($whileLoopVarCustomCmd -eq 1) {
        write-host ""
        Write-Host "Exchange Online Shell >>> " -NoNewline -ForegroundColor Yellow
        $ExchangeOnlineCustomCmd = Read-Host
        if ($ExchangeOnlineCustomCmd -eq "exit"){
            $whileLoopVarCustomCmd = 0
        }
        else {
            Invoke-Expression $ExchangeOnlineCustomCmd -Debug
            Write-Host ""
        }       
    }    
}
####################################







#####################################################################
# Main Code --- Main Code --- Main Code --- Main Code --- Main Code #
#-------------------------------------------------------------------#
Logo
Start-Sleep -Seconds 1
Set-ExecutionPolicy Bypass -Force -Scope Process -Confirm:$false
#####################################################################

#####################################################################

######################
# Requirements Check #
######################
CheckPowershellVersion
CheckAdminPrivs
if (!(Test-Path $tempdir)) {New-Item -itemType Directory -Path $tempdir | Out-Null}
CheckForUpdates
######################

##############
# Loop Menus #
##############
$WhileLoopVarMainMenu = 1
while ($WhileLoopVarMainMenu -eq 1){
    Clear-Host
    Start-Sleep -Milliseconds 250
    Logo
    Start-Sleep -Milliseconds 250
    Show-MainMenu
    $MainMenuChoice = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown").VirtualKeyCode
    switch ($MainMenuChoice) {
        49 {
            $WhileLoopVarMsGraphMenu = 1
            while ($WhileLoopVarMsGraphMenu -eq 1){
                Clear-Host
                Start-Sleep -Milliseconds 250
                Logo
                Start-Sleep -Milliseconds 250
                Show-MsGraphMenu
                $MsGraphMenuChoice = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown").VirtualKeyCode
                switch ($MsGraphMenuChoice) {
                    49 {try {installMicrosoftGraphModule}
                        catch {Write-Error "Error Running Module Installer"
                            CatchError}}
                    50 {try {CheckMsGraphDelegatedPermissions}
                        catch {Write-Error "Error Running Script"
                            CatchError}}
                    51 {try {MsGraphForcePasswordResetAllUsers}
                        catch {Write-Error "Error Running Script"
                            CatchError}}
                    52 {try {MsGraphForcePasswordResetSingleUser}
                        catch {Write-Error "Error Running Script"
                            CatchError}}
                    53 {try {CreateMFAStatusReport}
                        catch {Write-Error "Error Running Script"
                            CatchError}}
                    57 {
                        Write-Host "`nReturning to Main Menu" -ForegroundColor Cyan
                        $WhileLoopVarMsGraphMenu = 0}
                        
                        default {Write-Host "`nInvalid selection, please try again." -ForegroundColor Red
                            Start-Sleep -Seconds 1}
                }
            }
        }


        50 {
            $WhileLoopVarExchangeOnlineMenu = 1
            while ($WhileLoopVarExchangeOnlineMenu -eq 1){
                Clear-Host
                Start-Sleep -Milliseconds 250
                Logo
                Start-Sleep -Milliseconds 250
                $connectionStatus = "Disconnected"
                $connectionStatusColor = "Red"
                if (Get-Module ExchangeOnlineManagement -ListAvailable){
                    $ExchangeOnlineConnectionState = Get-ConnectionInformation
                    If ( $ExchangeOnlineConnectionState.State -match 'Connected' ){
                        $connectionStatus = "Connected"
                        $connectionStatusColor = "Green"
                    }
                    else {
                        $connectionStatus = "Disconnected"
                        $connectionStatusColor = "Red"
                    }
                }
                Show-ExchangeOnlineMenu
                $MsGraphMenuChoice = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown").VirtualKeyCode
                switch ($MsGraphMenuChoice) {
                    49 {try {installExchangeOnlineModule}
                        catch {Write-Error "Error Running Module Installer"
                            CatchError}}
                    50 {try {ExchangeOnlineConnection}
                        catch {Write-Error "Error while connecting"
                            CatchError}}
                    51 {try {ExchangeOnlineListMailbox}
                        catch {CatchError}}
                    52 {}
                    53 {}
                    54 {
                        $WhileLoopVarExchangeOnlineToolsMenu = 1
                        while ($WhileLoopVarExchangeOnlineToolsMenu -eq 1){
                            Clear-Host
                            Start-Sleep -Milliseconds 250
                            Logo
                            Start-Sleep -Milliseconds 250
                            Show-ExchangeOnlineToolsMenu
                            $MsGraphMenuChoice = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown").VirtualKeyCode
                            switch ($MsGraphMenuChoice) {
                                49 {try {ExchangeOnlineSaveSentItemsInSharedMailboxAll}
                                    catch {CatchError}}
                                57 {Write-Host "`nReturning to Exchange Online Menu" -ForegroundColor Cyan
                                $WhileLoopVarExchangeOnlineToolsMenu = 0}
                                
                                default {Write-Host "`nInvalid selection, please try again." -ForegroundColor Red
                                    Start-Sleep -Seconds 1}
                            }
                        }
                    }
                    56 {customExchangeCmd}
                    57 {Write-Host "`nReturning to Main Menu" -ForegroundColor Cyan
                        $WhileLoopVarExchangeOnlineMenu = 0}
                        
                        default {Write-Host "`nInvalid selection, please try again." -ForegroundColor Red
                            Start-Sleep -Seconds 1}
                }
            }
        }
        51 {Write-Host "3"}
        56 {CheckForUpdates}
        57 {
            Write-Host "`nExiting... Goodbye!" -ForegroundColor Cyan
            $WhileLoopVarMainMenu = 0
            Start-Sleep -Seconds 2
            exit
        }
        default {Write-Host "`nInvalid selection, please try again." -ForegroundColor Red
            Start-Sleep -Seconds 1}
    }

    if ($MainMenuChoice -ne 57) {
        Write-Host ""
        Start-Sleep -Milliseconds 250
    }
}
exit
##############