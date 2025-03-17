#######################################
# Configurable Variables
#--------------------------------------
$version = "2.3.0"
$ProgramName = "MS365-PSToolkit"
$programdir = "C:\MATRIXNET\$ProgramName-$version"
$GithubRepo = "https://github.com/N30X420/MS365-PSToolkit"
$ErrorLog = "$programdir\error.log"
#######################################
#######################################
$error.clear()
$ProgressPreference = 'Continue'
$host.UI.RawUI.WindowTitle = "$ProgramName - Version $version"
[console]::WindowWidth=200; [console]::WindowHeight=50; [console]::BufferWidth=[console]::WindowWidth
$LicenseFriendlyNameList = @{
"AAD_BASIC"="Microsoft Entra Basic"
"AAD_PREMIUM"="Microsoft Entra ID P1"
"AAD_PREMIUM_P1"="Microsoft Entra ID P1"
"AAD_PREMIUM_P2"="Microsoft Entra ID P2"
"ADALLOM_O365"="Office 365 Advanced Security Management"
"ADALLOM_STANDALONE"="Microsoft Cloud App Security"
"ADALLOM_S_O365"="POWER BI STANDALONE"
"ADALLOM_S_STANDALONE"="Microsoft Cloud App Security"
"ATA"="Azure Advanced Threat Protection for Users"
"ATP_ENTERPRISE"="Exchange Online Advanced Threat Protection"
"ATP_ENTERPRISE_FACULTY"="Exchange Online Advanced Threat Protection"
"BI_AZURE_P0"="Power BI (free)"
"BI_AZURE_P1"="Power BI Reporting and Analytics"
"BI_AZURE_P2"="Power BI Pro"
"CCIBOTS_PRIVPREV_VIRAL"="Dynamics 365 AI for Customer Service Virtual Agents Viral SKU"
"CRMINSTANCE"="Microsoft Dynamics CRM Online Additional Production Instance (Government Pricing)"
"CRMIUR"="CRM for Partners"
"CRMPLAN1"="Microsoft Dynamics CRM Online Essential (Government Pricing)"
"CRMPLAN2"="Dynamics CRM Online Plan 2"
"CRMSTANDARD"="CRM Online"
"CRMSTORAGE"="Microsoft Dynamics CRM Online Additional Storage"
"CRMTESTINSTANCE"="CRM Test Instance"
"DESKLESS"="Microsoft StaffHub"
"DESKLESSPACK"="Office 365 (Plan K1)"
"DESKLESSPACK_GOV"="Microsoft Office 365 (Plan K1) for Government"
"DESKLESSPACK_YAMMER"="Office 365 Enterprise K1 with Yammer"
"DESKLESSWOFFPACK"="Office 365 (Plan K2)"
"DESKLESSWOFFPACK_GOV"="Microsoft Office 365 (Plan K2) for Government"
"DEVELOPERPACK"="Office 365 Enterprise E3 Developer"
"DEVELOPERPACK_E5"="Microsoft 365 E5 Developer(without Windows and Audio Conferencing)"
"DMENTERPRISE"="Microsoft Dynamics Marketing Online Enterprise"
"DYN365_ENTERPRISE_CUSTOMER_SERVICE"="Dynamics 365 for Customer Service Enterprise Edition"
"DYN365_ENTERPRISE_P1_IW"="Dynamics 365 P1 Trial for Information Workers"
"DYN365_ENTERPRISE_PLAN1"="Dynamics 365 Plan 1 Enterprise Edition"
"DYN365_ENTERPRISE_SALES"="Dynamics 365 for Sales Enterprise Edition"
"DYN365_ENTERPRISE_SALES_CUSTOMERSERVICE"="Dynamics 365 for Sales and Customer Service Enterprise Edition"
"DYN365_ENTERPRISE_TEAM_MEMBERS"="Dynamics 365 for Team Members Enterprise Edition"
"DYN365_FINANCIALS_BUSINESS_SKU"="Dynamics 365 for Financials Business Edition"
"DYN365_MARKETING_USER"="Dynamics 365 for Marketing USL"
"DYN365_MARKETING_APP"="Dynamics 365 Marketing"
"DYN365_SALES_INSIGHTS"="Dynamics 365 AI for Sales"
"D365_SALES_PRO"="Dynamics 365 for Sales Professional"
"Dynamics_365_for_Operations"="Dynamics 365 Unf Ops Plan Ent Edition"
"ECAL_SERVICES"="ECAL"
"EMS"="Enterprise Mobility + Security E3"
"EMSPREMIUM"="Enterprise Mobility + Security E5"
"ENTERPRISEPACK"="Office 365 Enterprise E3"
"ENTERPRISEPACKLRG"="Office 365 Enterprise E3 LRG"
"ENTERPRISEPACKWITHOUTPROPLUS"="Office 365 Enterprise E3 without ProPlus Add-on"
"ENTERPRISEPACK_B_PILOT"="Office 365 (Enterprise Preview)"
"ENTERPRISEPACK_FACULTY"="Office 365 (Plan A3) for Faculty"
"ENTERPRISEPACK_GOV"="Microsoft Office 365 (Plan G3) for Government"
"ENTERPRISEPACK_STUDENT"="Office 365 (Plan A3) for Students"
"ENTERPRISEPREMIUM"="Enterprise E5 (with Audio Conferencing)"
"ENTERPRISEPREMIUM_NOPSTNCONF"="Enterprise E5 (without Audio Conferencing)"
"ENTERPRISEWITHSCAL"="Office 365 Enterprise E4"
"ENTERPRISEWITHSCAL_FACULTY"="Office 365 (Plan A4) for Faculty"
"ENTERPRISEWITHSCAL_GOV"="Microsoft Office 365 (Plan G4) for Government"
"ENTERPRISEWITHSCAL_STUDENT"="Office 365 (Plan A4) for Students"
"EOP_ENTERPRISE"="Exchange Online Protection"
"EOP_ENTERPRISE_FACULTY"="Exchange Online Protection for Faculty"
"EQUIVIO_ANALYTICS"="Office 365 Advanced Compliance"
"EQUIVIO_ANALYTICS_FACULTY"="Office 365 Advanced Compliance for Faculty"
"ESKLESSWOFFPACK_GOV"="Microsoft Office 365 (Plan K2) for Government"
"EXCHANGEARCHIVE"="Exchange Online Archiving"
"EXCHANGEARCHIVE_ADDON"="Exchange Online Archiving for Exchange Online"
"EXCHANGEDESKLESS"="Exchange Online Kiosk"
"EXCHANGEENTERPRISE"="Exchange Online (Plan 2)"
"EXCHANGEENTERPRISE_FACULTY"="Exch Online Plan 2 for Faculty"
"EXCHANGEENTERPRISE_GOV"="Microsoft Office 365 Exchange Online (Plan 2) only for Government"
"EXCHANGEESSENTIALS"="Exchange Online Essentials"
"EXCHANGESTANDARD"="Exchange Online (Plan 1)"
"EXCHANGESTANDARD_GOV"="Microsoft Office 365 Exchange Online (Plan 1) only for Government"
"EXCHANGESTANDARD_STUDENT"="Exchange Online (Plan 1) for Students"
"EXCHANGETELCO"="Exchange Online POP"
"EXCHANGE_ANALYTICS"="Microsoft MyAnalytics"
"EXCHANGE_L_STANDARD"="Exchange Online (Plan 1)"
"EXCHANGE_S_ARCHIVE_ADDON_GOV"="Exchange Online Archiving"
"EXCHANGE_S_DESKLESS"="Exchange Online Kiosk"
"EXCHANGE_S_DESKLESS_GOV"="Exchange Kiosk"
"EXCHANGE_S_ENTERPRISE"="Exchange Online (Plan 2) Ent"
"EXCHANGE_S_ENTERPRISE_GOV"="Exchange Plan 2G"
"EXCHANGE_S_ESSENTIALS"="Exchange Online Essentials"
"EXCHANGE_S_FOUNDATION"="Exchange Foundation for certain SKUs"
"EXCHANGE_S_STANDARD"="Exchange Online (Plan 2)"
"EXCHANGE_S_STANDARD_MIDMARKET"="Exchange Online (Plan 1)"
"FLOW_FREE"="Microsoft Flow (Free)"
"FLOW_O365_P2"="Flow for Office 365"
"FLOW_O365_P3"="Flow for Office 365"
"FLOW_P1"="Microsoft Flow Plan 1"
"FLOW_P2"="Microsoft Flow Plan 2"
"FORMS_PLAN_E3"="Microsoft Forms (Plan E3)"
"FORMS_PLAN_E5"="Microsoft Forms (Plan E5)"
"INFOPROTECTION_P2"="Azure Information Protection Premium P2"
"INTUNE_A"="Intune"
"INTUNE_A_D"="Microsoft Intune Device"
"INTUNE_A_VL"="Intune (Volume License)"
"INTUNE_O365"="Mobile Device Management for Office 365"
"INTUNE_STORAGE"="Intune Extra Storage"
"IT_ACADEMY_AD"="Microsoft Imagine Academy"
"LITEPACK"="Office 365 (Plan P1)"
"LITEPACK_P2"="Office 365 Small Business Premium"
"LOCKBOX"="Customer Lockbox"
"LOCKBOX_ENTERPRISE"="Customer Lockbox"
"M365_Copilot"="Microsoft Copilot for Microsoft 365"
"MCOCAP"="Command Area Phone"
"MCOEV"="Skype for Business Cloud PBX"
"MCOIMP"="Skype for Business Online (Plan 1)"
"MCOLITE"="Lync Online (Plan 1)"
"MCOMEETADV"="PSTN conferencing"
"MCOPLUSCAL"="Skype for Business Plus CAL"
"MCOPSTN1"="Skype for Business Pstn Domestic Calling"
"MCOPSTN2"="Skype for Business Pstn Domestic and International Calling"
"MCOSTANDARD"="Skype for Business Online Standalone Plan 2"
"MCOSTANDARD_GOV"="Lync Plan 2G"
"MCOSTANDARD_MIDMARKET"="Lync Online (Plan 1)"
"MCVOICECONF"="Lync Online (Plan 3)"
"MDE_SMB"="Microsoft Defender for Business"
"MDM_SALES_COLLABORATION"="Microsoft Dynamics Marketing Sales Collaboration"
"MEE_FACULTY"="Minecraft Education Edition Faculty"
"MEE_STUDENT"="Minecraft Education Edition Student"
"MEETING_ROOM"="Meeting Room"
"MFA_PREMIUM"="Azure Multi-Factor Authentication"
"MICROSOFT_BUSINESS_CENTER"="Microsoft Business Center"
"Microsoft_365_Copilot"="Copilot for Microsoft 365"
"MICROSOFT_REMOTE_ASSIST"="Dynamics 365 Remote Assist"
"MIDSIZEPACK"="Office 365 Midsize Business"
"MINECRAFT_EDUCATION_EDITION"="Minecraft Education Edition Faculty"
"MS-AZR-0145P"="Azure"
"MS_TEAMS_IW"="Microsoft Teams"
"NBPOSTS"="Microsoft Social Engagement Additional 10k Posts (minimum 100 licenses) (Government Pricing)"
"NBPROFESSIONALFORCRM"="Microsoft Social Listening Professional"
"O365_BUSINESS"="Microsoft 365 Apps for business"
"O365_BUSINESS_ESSENTIALS"="Microsoft 365 Business Basic"
"O365_BUSINESS_PREMIUM"="Microsoft 365 Business Standard"
"OFFICE365_MULTIGEO"="Multi-Geo Capabilities in Office 365"
"OFFICESUBSCRIPTION"="Microsoft 365 Apps for enterprise"
"OFFICESUBSCRIPTION_FACULTY"="Office 365 ProPlus for Faculty"
"OFFICESUBSCRIPTION_GOV"="Office ProPlus"
"OFFICESUBSCRIPTION_STUDENT"="Office ProPlus Student Benefit"
"OFFICE_FORMS_PLAN_2"="Microsoft Forms (Plan 2)"
"OFFICE_PRO_PLUS_SUBSCRIPTION_SMBIZ"="Office ProPlus"
"ONEDRIVESTANDARD"="OneDrive"
"PAM_ENTERPRISE "="Exchange Primary Active Manager"
"PLANNERSTANDALONE"="Planner Standalone"
"POWERAPPS_INDIVIDUAL_USER"="Microsoft PowerApps and Logic flows"
"POWERAPPS_O365_P2"="PowerApps"
"POWERAPPS_O365_P3"="PowerApps for Office 365"
"POWERAPPS_VIRAL"="PowerApps (Free)"
"POWERFLOW_P1"="Microsoft PowerApps Plan 1"
"POWERFLOW_P2"="Microsoft PowerApps Plan 2"
"POWER_BI_ADDON"="Office 365 Power BI Addon"
"POWER_BI_INDIVIDUAL_USE"="Power BI Individual User"
"POWER_BI_INDIVIDUAL_USER"="Power BI for Office 365 Individual"
"POWER_BI_PRO"="Power BI Pro"
"POWER_BI_STANDALONE"="Power BI Standalone"
"POWER_BI_STANDARD"="Power-BI Standard"
"PREMIUM_ADMINDROID"="AdminDroid Office 365 Reporter"
"PROJECTCLIENT"="Project Professional"
"PROJECTESSENTIALS"="Project Lite"
"PROJECTONLINE_PLAN_1"="Project Online (Plan 1)"
"PROJECTONLINE_PLAN_1_FACULTY"="Project Online for Faculty Plan 1"
"PROJECTONLINE_PLAN_1_STUDENT"="Project Online for Students Plan 1"
"PROJECTONLINE_PLAN_2"="Project Online and PRO"
"PROJECTONLINE_PLAN_2_FACULTY"="Project Online for Faculty Plan 2"
"PROJECTONLINE_PLAN_2_STUDENT"="Project Online for Students Plan 2"
"PROJECTPREMIUM"="Project Online Premium"
"PROJECTPROFESSIONAL"="Project Online Pro"
"PROJECTWORKMANAGEMENT"="Office 365 Planner Preview"
"PROJECT_CLIENT_SUBSCRIPTION"="Project Pro for Office 365"
"PROJECT_ESSENTIALS"="Project Lite"
"PROJECT_MADEIRA_PREVIEW_IW_SKU"="Dynamics 365 for Financials for IWs"
"PROJECT_ONLINE_PRO"="Project Online Plan 3"
"RIGHTSMANAGEMENT"="Azure Rights Management Premium"
"RIGHTSMANAGEMENT_ADHOC"="Windows Azure Rights Management"
"RIGHTSMANAGEMENT_STANDARD_FACULTY"="Azure Rights Management for faculty"
"RIGHTSMANAGEMENT_STANDARD_STUDENT"="Information Rights Management for Students"
"RMS_S_ENTERPRISE"="Azure Active Directory Rights Management"
"RMS_S_ENTERPRISE_GOV"="Windows Azure Active Directory Rights Management"
"RMS_S_PREMIUM"="Azure Information Protection Plan 1"
"RMS_S_PREMIUM2"="Azure Information Protection Premium P2"
"SCHOOL_DATA_SYNC_P1"="School Data Sync (Plan 1)"
"SHAREPOINTDESKLESS"="SharePoint Online Kiosk"
"SHAREPOINTDESKLESS_GOV"="SharePoint Online Kiosk"
"SHAREPOINTENTERPRISE"="SharePoint Online (Plan 2)"
"SHAREPOINTENTERPRISE_EDU"="SharePoint Plan 2 for EDU"
"SHAREPOINTENTERPRISE_GOV"="SharePoint Plan 2G"
"SHAREPOINTENTERPRISE_MIDMARKET"="SharePoint Online (Plan 1)"
"SHAREPOINTLITE"="SharePoint Online (Plan 1)"
"SHAREPOINTPARTNER"="SharePoint Online Partner Access"
"SHAREPOINTSTANDARD"="SharePoint Online Plan 1"
"SHAREPOINTSTANDARD_EDU"="SharePoint Plan 1 for EDU"
"SHAREPOINTSTORAGE"="SharePoint Online Storage"
"SHAREPOINTWAC"="Office Online"
"SHAREPOINTWAC_EDU"="Office Online for Education"
"SHAREPOINTWAC_GOV"="Office Online for Government"
"SHAREPOINT_PROJECT"="SharePoint Online (Plan 2) Project"
"SHAREPOINT_PROJECT_EDU"="Project Online Service for Education"
"SMB_APPS"="Business Apps (free)"
"SMB_BUSINESS"="Microsoft 365 Apps for Business"
"SMB_BUSINESS_ESSENTIALS"="Microsoft 365 Business Basic"
"SMB_BUSINESS_PREMIUM"="Microsoft 365 Business Standard - Prepaid Legacy"
"SPZA IW"="Microsoft PowerApps Plan 2 Trial"
"SPB"="Microsoft 365 Business Premium"
"SPE_E3"="Microsoft 365 E3"
"SPE_E5"="Microsoft 365 E5"
"SQL_IS_SSIM"="Power BI Information Services"
"STANDARDPACK"="Office 365 (Plan E1)"
"STANDARDPACK_FACULTY"="Office 365 (Plan A1) for Faculty"
"STANDARDPACK_GOV"="Microsoft Office 365 (Plan G1) for Government"
"STANDARDPACK_STUDENT"="Office 365 (Plan A1) for Students"
"STANDARDWOFFPACK"="Office 365 (Plan E2)"
"STANDARDWOFFPACKPACK_FACULTY"="Office 365 (Plan A2) for Faculty"
"STANDARDWOFFPACKPACK_STUDENT"="Office 365 (Plan A2) for Students"
"STANDARDWOFFPACK_FACULTY"="Office 365 Education E1 for Faculty"
"STANDARDWOFFPACK_GOV"="Microsoft Office 365 (Plan G2) for Government"
"STANDARDWOFFPACK_IW_FACULTY"="Office 365 Education for Faculty"
"STANDARDWOFFPACK_IW_STUDENT"="Office 365 Education for Students"
"STANDARDWOFFPACK_STUDENT"="Microsoft Office 365 (Plan A2) for Students"
"STANDARD_B_PILOT"="Office 365 (Small Business Preview)"
"STREAM"="Microsoft Stream"
"STREAM_O365_E3"="Microsoft Stream for O365 E3 SKU"
"STREAM_O365_E5"="Microsoft Stream for O365 E5 SKU"
"SWAY"="Sway"
"TEAMS1"="Microsoft Teams"
"TEAMS_COMMERCIAL_TRIAL"="Microsoft Teams Commercial Cloud Trial"
"THREAT_INTELLIGENCE"="Office 365 Threat Intelligence"
"VIDEO_INTEROP "="Skype Meeting Video Interop for Skype for Business"
"VISIO_PLAN1_DEPT"="Visio Plan 1"
"VISIO_PLAN2_DEPT"="Visio Plan 2"
"VISIOCLIENT"="Visio Online Plan 2"
"VISIOONLINE_PLAN1"="Visio Online Plan 1"
"VISIO_CLIENT_SUBSCRIPTION"="Visio Pro for Office 365"
"WACONEDRIVEENTERPRISE"="OneDrive for Business (Plan 2)"
"WACONEDRIVESTANDARD"="OneDrive for Business with Office Online"
"WACSHAREPOINTSTD"="Office Online STD"
"WHITEBOARD_PLAN3"="White Board (Plan 3)"
"WIN_DEF_ATP"="Microsoft Defender for Endpoint"
"WIN10_PRO_ENT_SUB"="Windows 10 Enterprise E3"
"WIN10_VDA_E3"="Windows E3"
"WIN10_VDA_E5"="Windows E5"
"WINDOWS_STORE"="Windows Store"
"YAMMER_EDU"="Yammer for Academic"
"YAMMER_ENTERPRISE"="Yammer for the Starship Enterprise"
"YAMMER_ENTERPRISE_STANDALONE"="Yammer Enterprise"
"YAMMER_MIDSIZE"="Yammer"
}
#######################################
function SplashLogo {
    write-host ""
    write-host "███╗   ███╗ █████╗ ████████╗██████╗ ██╗██╗  ██╗" -ForegroundColor White -NoNewline
    Write-host "███╗   ██╗███████╗████████╗" -ForegroundColor Red
    write-host "████╗ ████║██╔══██╗╚══██╔══╝██╔══██╗██║╚██╗██╔╝" -ForegroundColor white -NoNewline
    Write-Host "████╗  ██║██╔════╝╚══██╔══╝" -ForegroundColor Red
    write-host "██╔████╔██║███████║   ██║   ██████╔╝██║ ╚███╔╝ " -ForegroundColor White -NoNewline
    Write-Host "██╔██╗ ██║█████╗     ██║   " -ForegroundColor Red
    write-host "██║╚██╔╝██║██╔══██║   ██║   ██╔══██╗██║ ██╔██╗ " -ForegroundColor White -NoNewline
    Write-Host "██║╚██╗██║██╔══╝     ██║   " -ForegroundColor Red
    write-host "██║ ╚═╝ ██║██║  ██║   ██║   ██║  ██║██║██╔╝ ██╗" -ForegroundColor White -NoNewline
    write-host "██║ ╚████║███████╗   ██║   " -ForegroundColor Red
    write-host "╚═╝     ╚═╝╚═╝  ╚═╝   ╚═╝   ╚═╝  ╚═╝╚═╝╚═╝  ╚═╝" -ForegroundColor White -NoNewline
    write-host "╚═╝  ╚═══╝╚══════╝   ╚═╝   " -ForegroundColor Red                                         
}
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
    Write-Host "(3) - Sharepoint Online"
    Write-Host "`n(8) - Check For Updates"
    Write-Host "(9) - Exit"
    Write-Host "###############################################" -ForegroundColor DarkCyan
    Write-Host ""
}
function CheckPowershellVersion {
    try {
        if ($PSVersionTable.PSVersion.Major -ne 7) {
            throw "This program requires PowerShell 7.x You are running version $($PSVersionTable.PSVersion.Major).$($PSVersionTable.PSVersion.Minor). Please update to PowerShell 7.x"
        }
        else {
            Write-Host "Running Powershell v$($PSVersionTable.PSVersion.Major).$($PSVersionTable.PSVersion.Minor).$($PSVersionTable.PSVersion.Patch)" -ForegroundColor Green
        }
    } catch {
        Write-ErrorLog $_.Exception.Message
        Write-Warning $_.Exception.Message
        exit
    }
}
function CheckAdminPrivs {
    function isadmin {
        # Returns true/false
        ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
    }
    try {
        if (isadmin -eq "True") {
            Write-Host "Got Administrator Permissions" -ForegroundColor Green
            Start-Sleep -Seconds 1
        } else {
            throw "This script needs Administrator Privileges to work its magic"
        }
        if (-Not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] 'Administrator')) {
            if ([int](Get-CimInstance -Class Win32_OperatingSystem | Select-Object -ExpandProperty BuildNumber) -ge 6000) {
                $CommandLine = "-File `"" + $MyInvocation.MyCommand.Path + "`" " + $MyInvocation.UnboundArguments
                Start-Process -FilePath PowerShell.exe -Verb Runas -ArgumentList $CommandLine
                Exit
            }
        }
    } catch {
        Write-ErrorLog $_.Exception.Message
        Write-Error $_.Exception.Message
        Start-Sleep -Seconds 3
        exit
    }
}
function CheckForUpdates {
    try {
        $Releases = Invoke-RestMethod -Uri "https://api.github.com/repos/N30X420/MS365-PSToolkit/releases"
		$ReleaseInfo = ($Releases | Sort-Object id -desc)[0]
		$LatestVersion = [version[]]$ReleaseInfo.Name.Trim('v')
		if ($LatestVersion -gt $version){ $Script:NewVersionAvailable = "v$LatestVersion is available $(Format-Hyperlink -Uri "$GithubRepo" -Label "v$LatestVersion")"}
        else {
            Write-Host "You are running the latest version - " -ForegroundColor Green -NoNewline
            Write-Host "v$version" -ForegroundColor Blue
            Start-Sleep -Seconds 2} 
    }
    catch {
        Write-ErrorLog $_.Exception.Message
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
    Write-Error "Runtime Error: $($_.Exception.Message)"
    Write-Host -NoNewLine "Press any key to return to the menu..." -ForegroundColor Yellow
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
    Write-ErrorLog "Runtime Error: $($_.Exception.Message)"
}
function Write-ErrorLog {
    param (
        [string]$Message
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "$timestamp - $Message"
    Add-Content -Path $ErrorLog -Value $logMessage
}
function PromptExportToCSV {
    $script:ExportCSV = 0
    Write-Host "Export data to CSV ? [y/n]" -ForegroundColor Yellow
    $UserInput = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    If ($UserInput.Character -eq "y"){
        Write-Host "Exporting Data" -ForegroundColor Yellow
        $script:ExportCSV = 1
    }
    else {
        Break
    }
}
function PromptPressKeyToContinue {
    Write-Host -NoNewLine "Press any key to return to the menu..." -ForegroundColor Yellow
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
}
function DisplayBug {
    Write-Warning "If no data is displayed rerun the command. CSV export does work even when no data is displayed."
}

####################################
#  MsGraph  #
#-------------
function Show-MsGraphMenu {
    Write-Host "############# Microsoft Graph Menu #############" -ForegroundColor DarkCyan
    Write-Host "(1) - Install Required Modules"
    Write-Host "(2) - Check required permission scopes for MsGraph Command"
    Write-Host "(3) - Report - Create MFA status report (CSV Export)"
    Write-Host "(4) - Report - Create license usage report (CSV Export)"
    Write-Host "(5) - Show users with and without licenses"
    Write-Host "(8) - Tools Menu"
    Write-Host "`n(9) - Main Menu"
    Write-Host "################################################" -ForegroundColor DarkCyan
    Write-Host ""
}
function Show-MsGraphToolsMenu {
    Write-Host "############ Microsoft Graph Tools Menu ############" -ForegroundColor DarkCyan
    Write-Host "`n(1) - Tool - Remove obsolete MS365-PSToolkit entra applications"
    Write-Host "(2) - Tool - Forced password change at next login - " -NoNewline 
    write-host "single user" -ForegroundColor Yellow
    Write-Host "(3) - Tool - Forced password change at next login - " -NoNewline 
    write-host "all users" -ForegroundColor Red
    Write-Host "(4) - "
    Write-Host "(5) - "
    Write-Host "(6) - "
    Write-Host "(7) - "
    Write-Host "`n(9) - Return to Microsoft Graph Menu"
    Write-Host "####################################################" -ForegroundColor DarkCyan
    Write-Host ""
}

function installMicrosoftGraphModule {
    if(-not (Get-Module Microsoft.Graph -ListAvailable)){
        Write-Warning "Module Microsoft.Graph not installed"
        Write-Host "Installing Microsoft.Graph Please Wait" -ForegroundColor Green
        Install-Module Microsoft.Graph -Scope CurrentUser -Force
        Install-Module Microsoft.Graph.Beta -Scope CurrentUser -Force -Confirm:$false
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
    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

        # Disconnect from the Microsoft Graph If already connected
        if (Get-MgContext) {
            Write-Host Disconnecting from the previous session.... -ForegroundColor Yellow
            Disconnect-MgGraph | Out-Null
        }

        Write-Warning "This script will force every account to change password at next sign in."
        Write-Host "`nContinue ? [y/n]"
        $continue = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        if ($continue.Character -ine "y") {
            Break
        }

        # Connect to Microsoft Graph API
        Write-Host "`nA new browser window will open for you to sign in using your Microsoft 365 Global Admin Account" -ForegroundColor Yellow
        Start-Sleep -Seconds 2
        Connect-MgGraph -Scopes "User.Read.All", "User-PasswordProfile.ReadWrite.All", "UserAuthenticationMethod.ReadWrite.All" -NoWelcome

        Write-Host "Would you like to exclude an account from this script? [y/n]"
        $exclude = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        if ($exclude.Character -eq "y") {
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
            if ($user.Id -eq $excludedAccount.Id) { continue }
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
                Update-MgUser -UserId $user.id -PasswordProfile @{ ForceChangePasswordNextSignIn = $true }
            } catch {
                Write-ErrorLog "Error updating $($user.UserPrincipalName): $_.Exception.Message"
                Write-Warning "Error updating $($user.UserPrincipalName)"
            }
        }
        # Clear progress bar
        Write-Progress -Activity "Processing Users" -Completed
    } catch {
        Write-ErrorLog $_.Exception.Message
        Write-Error $_.Exception.Message
    }
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
function MsGraphRemoveObsoleteMS365ToolKitEntraApplication {
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

    #Disconnect from the Microsoft Graph If already connected
    if (Get-MgContext) {
        Write-Host Disconnecting from the previous sesssion.... -ForegroundColor Yellow
        Disconnect-MgGraph | Out-Null
    }
    Write-Host "`nA new browser window will open for you to sign in using your Microsoft 365 Global Admin Account" -ForegroundColor Yellow
    Start-Sleep -Seconds 2
    
    Connect-MgGraph -Scopes "Application.ReadWrite.All" -NoWelcome

    Write-Warning "This script will remove any obsolete MS365-PSToolkit Application in Microsoft Entra"
    Write-Host "`nContinue ? [y/n]"
    $continue = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    if ($continue.Character -ine "y"){
        Break
    }
    $EntraApplication = Get-MgApplication | Where-Object -Property DisplayName -Match "PnP-MS365*"
    $counter = 0
    $totalApps = $EntraApplication.Count
    foreach ($ID in $EntraApplication.Id){
        $counter++
        $percentComplete = [math]::Round(($counter / $totalApps) * 100)

        $progressParams = @{
            Activity        = "Processing Applications"
            Status          = "Application $($counter) of $totalApps - $($EntraApplication.DisplayName) - $percentComplete% Complete"
            PercentComplete = $percentComplete
        }

        Write-Progress @progressParams
        Remove-MgApplication -ApplicationId $ID
        Start-Sleep -Seconds 1
    }
    Write-Progress -Activity "Processing Applications" -Completed
}
function MsGraphCreateMFAStatusReport {
    try {
        # Disconnect from the Microsoft Graph If already connected
        if (Get-MgContext) {
            Write-Host Disconnecting from the previous session.... -ForegroundColor Yellow
            Disconnect-MgGraph | Out-Null
        }

        Write-Host "`nA new browser window will open for you to sign in using your Microsoft 365 Global Admin Account" -ForegroundColor Yellow
        Start-Sleep -Seconds 2

        # Connect to Microsoft Graph API
        Connect-MgGraph -Scopes "User.Read.All", "UserAuthenticationMethod.Read.All", "UserAuthenticationMethod.ReadWrite.All" -NoWelcome

        # Create variable for the date stamp
        $LogDate = Get-Date -f yyyyMMddhhmm

        # Define CSV file export location variable
        $Csvfile = "$programdir\MFA-Users-Report_$LogDate.csv"

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
                } else {
                    $ReportLine.DefaultMFAMethod = "Not set"
                }
            } catch {
                Write-ErrorLog "Failed to retrieve default MFA method for $($user.UserPrincipalName): $_.Exception.Message"
                Write-Warning "Failed to retrieve default MFA method for $($user.UserPrincipalName)"
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

        PromptPressKeyToContinue
    } 
    catch {
        Write-ErrorLog $_.Exception.Message
        Write-Error $_.Exception.Message
    }
}

function MsGraphCreate365LicenseReport {
    if (Get-MgContext) {
        Write-Host Disconnecting from the previous session.... -ForegroundColor Yellow
        Disconnect-MgGraph | Out-Null
    }

    Write-Host "`nA new browser window will open for you to sign in using your Microsoft 365 Global Admin Account" -ForegroundColor Yellow
    Start-Sleep -Seconds 2

    Connect-MgGraph -Scopes "Directory.ReadWrite.All" -NoWelcome

    $OutputCSVName = "Microsoft365LicenseUsageReport_$((Get-Date -format yyyy-MM-dd-HH-mm).ToString()).csv"
    $Csvfile = "$programdir\$OutputCSVName"
    $Result = ""  
    $Results = @()
    Write-Host "Generating Microsoft 365 license usage report..." -ForegroundColor Yellow
    $counter = 0
    $TotalLicenses = (Get-MgBetaSubscribedSku).Count
    Get-MgBetaSubscribedSku | ForEach-Object {
        $AccountSkuID = $_.SkuID
        $LicensePlan = $_.SkuPartNumber
        $ActiveUnits = $_.PrepaidUnits.Enabled
        $ConsumedUnits = $_.ConsumedUnits

        $counter++
        $percentComplete = [math]::Round(($counter / $TotalLicenses) * 100)

        $progressParams = @{
            Activity        = "Retrieving license info"
            Status          = "License $($counter) of $TotalLicenses - Currently Processing: $LicensePlan - $percentComplete% Complete"
            PercentComplete = $percentComplete
        }
        
        Write-Progress @progressParams
        $EasyName = $LicenseFriendlyNameList[$LicensePlan]

        if (!($EasyName))
            { $FriendlyName = $LicensePlan}
        else
            { $FriendlyName = $EasyName}
        
        $Result = @{'AccountSkuId' = $AccountSkuID; 'AccountSkuPartNumber' = $LicensePlan; 'License Plan_Friendly Name' = $FriendlyName; 'Active Units' = $ActiveUnits; 'Consumed Units' = $ConsumedUnits }
        $Results = New-Object PSObject -Property $Result
        $Results | select-object 'AccountSkuId','AccountSkuPartNumber', 'License Plan_Friendly Name', 'Active Units', 'Consumed Units' | Export-Csv -Path $Csvfile -Notype -Append
        Start-Sleep -Milliseconds 300
    }
    Write-Progress -Activity "Retrieving license info" -Completed

    Write-Host "Script completed. Results exported to $Csvfile." -ForegroundColor Cyan
    Start-Process $Csvfile
    Disconnect-MgGraph | Out-Null
    PromptPressKeyToContinue
}

function MsGraphShowUsersWithAndWithoutLicense {
    if (Get-MgContext) {
        Write-Host Disconnecting from the previous session.... -ForegroundColor Yellow
        Disconnect-MgGraph | Out-Null
    }

    Write-Host "`nA new browser window will open for you to sign in using your Microsoft 365 Global Admin Account" -ForegroundColor Yellow
    Start-Sleep -Seconds 2

    Connect-MgGraph -Scopes "Directory.ReadWrite.All" -NoWelcome

    Write-Host "Users with Licenses" -ForegroundColor Yellow
    Get-MgBetaUser -All | Where-Object {($_.AssignedLicenses.Count) -ne 0 } | Select-Object 'DisplayName', 'UserPrincipalName', 'AccountEnabled' | Format-Table -AutoSize
    Write-Host "Users without Licenses" -ForegroundColor Yellow
    Get-MgBetaUser -All | Where-Object {($_.AssignedLicenses.Count) -eq 0 } | Select-Object 'DisplayName', 'UserPrincipalName', 'AccountEnabled' | Format-Table -AutoSize
    PromptPressKeyToContinue
}
####################################

####################################
#  Exchange Online  #
#-----------------------
function Show-ExchangeOnlineMenu {
    Write-Host "######## Microsoft Exchange Online Menu ########" -ForegroundColor DarkCyan
    Write-Host "Exchange Online Connection : " -NoNewline -ForegroundColor Yellow
    Write-Host "$ExchangeOnlineConnectionStatus" -ForegroundColor $ExchangeOnlineConnectionStatusColor
    Write-Host "`n(1) - Install Required Modules"
    Write-Host "(2) - Connect / Disconnect Exchange Online"
    Write-Host "(3) - List Mailbox"
    Write-Host "(4) - List Mailbox Permissions"
    Write-Host "(5) - List Mailbox Permissions By User"
    Write-Host "(6) - List Mailbox Size"
    Write-Host "(7) - Open Tools Menu"
    Write-Host "(8) - Custom Command"
    Write-Host "`n(9) - Main Menu"
    Write-Host "################################################" -ForegroundColor DarkCyan
    Write-Host ""
    Import-Module ExchangeOnlineManagement -ErrorAction SilentlyContinue
}
function Show-ExchangeOnlineToolsMenu {
    Write-Host "######## Microsoft Exchange Online Tools Menu ########" -ForegroundColor DarkCyan
    Write-Host "Exchange Online Connection : " -NoNewline -ForegroundColor Yellow
    Write-Host "$ExchangeOnlineConnectionStatus" -ForegroundColor $ExchangeOnlineConnectionStatusColor
    Write-Host "`n(1) - Enable Delegate Sent Items Shared Mailbox"
    Write-Host "`n(9) - Return to Exchange Online Menu"
    Write-Host "#####################################################" -ForegroundColor DarkCyan
    Write-Host ""
}
function installExchangeOnlineModule {
    if(-not (Get-Module ExchangeOnlineManagement -ListAvailable)){
        Write-Warning "Module ExchangeOnlineManagement not installed"
        Write-Host "Installing ExchangeOnlineManagement Please Wait" -ForegroundColor Green
        Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force -Confirm:$false
        }
    Start-Sleep -Seconds 1
    Write-Host "Module Already Installed" -ForegroundColor Green
    Start-Sleep -Seconds 2
}
function ExchangeOnlineConnection {
    try {
        $ExchangeOnlineConnection = Get-ConnectionInformation
        If (-Not ($ExchangeOnlineConnection.State -match 'Connected')) {
            Write-Host "Connecting to Microsoft 365 - Exchange Online" -ForegroundColor Yellow
            Connect-ExchangeOnline -ShowBanner:$False
        } else {
            Write-Host "Disconnecting Microsoft 365 - Exchange Online" -ForegroundColor Yellow
            Start-Sleep -Seconds 2
            Disconnect-ExchangeOnline -Confirm:$false
        }
    } catch {
        Write-ErrorLog $_.Exception.Message
        Write-Error $_.Exception.Message
    }
}
function ExchangeOnlineListMailbox {
    $ListMailbox=$null
    $RecipientType='RoomMailbox', 'LinkedRoomMailbox', 'EquipmentMailbox', 'SchedulingMailbox', 
        'LegacyMailbox', 'LinkedMailbox', 'UserMailbox', 'MailContact', 'DynamicDistributionGroup', 'MailForestContact', 'MailNonUniversalGroup', 'MailUniversalSecurityGroup', 
        'RoomList', 'MailUser', 'GuestMailUser', 'GroupMailbox', 'PublicFolder', 'TeamMailbox', 'SharedMailbox', 'RemoteUserMailbox', 'RemoteRoomMailbox', 'RemoteEquipmentMailbox', 
        'RemoteTeamMailbox', 'RemoteSharedMailbox', 'PublicFolderMailbox', 'SharedWithMailUser'
    $ListMailbox = Get-Recipient -ResultSize Unlimited -RecipientTypeDetails UserMailbox,SharedMailbox | select-object Displayname,RecipientType,PrimarySMTPAddress,EmailAddresses | Sort-Object -Property DisplayName | Format-Table -AutoSize
    Start-Sleep -Milliseconds 250
    DisplayBug
    Write-Host "####################################################" -ForegroundColor Green
    Write-Output $ListMailbox
    Write-Host "####################################################" -ForegroundColor Green
    Start-Sleep -Seconds 3
    PromptExportToCSV
    if ($script:ExportCSV -eq 1){
        $ExportResult=""   
        $ExportResults=@()  
        $OutputCSV="$programdir\Office365EmailAddressesReport_$((Get-Date -format yyyy-MM-dd-HH-mm).ToString()).csv"
        # Initialize progress counter
        $Counter = 0
        $totalmailboxes = $ListMailbox.Count
        #Get all Email addresses in Microsoft 365
        $ListMailbox = Get-Recipient -ResultSize Unlimited -RecipientTypeDetails $RecipientType | select-object Displayname,RecipientType,PrimarySMTPAddress,EmailAddresses
        $ListMailbox | ForEach-Object {
            $Counter++
            # Calculate percentage completion
            $percentComplete = [math]::Round(($counter / $totalmailboxes) * 100)

            # Define progress bar parameters with user principal name
            $progressParams = @{
                Activity        = "Processing Mailboxes"
                Status          = "Mailbox $($counter) of $totalMailboxes - $($_.Displayname) - $percentComplete% Complete"
                PercentComplete = $percentComplete
            }

            Write-Progress @progressParams

            $DisplayName=$_.DisplayName
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
            Start-Sleep -Milliseconds 100
        }
    }
    # Clear progress bar
    Write-Progress -Activity "Processing Users" -Completed
    Write-Host "Script completed. Results exported to  $OutputCSV." -ForegroundColor Cyan
}
function ExchangeOnlineListPermissions {
    $ListMailboxPermissions = Get-mailbox -RecipientTypeDetails UserMailbox,SharedMailbox -ResultSize Unlimited | Get-MailboxPermission -ErrorAction SilentlyContinue | Sort-Object -Property Identity | where-object -Property User -NotMatch "NT AUTHORITY" | Format-Table identity, accessrights, User -AutoSize
    Start-Sleep -Milliseconds 250
    DisplayBug
    Write-Host "####################################################" -ForegroundColor Green
    Write-Output $ListMailboxPermissions
    Write-Host "####################################################" -ForegroundColor Green
    Start-Sleep -Seconds 3
    PromptExportToCSV
    if ($script:ExportCSV -eq 1){
        $ExportCSV="$programdir\SharedMBPermissionReport_$((Get-Date -format yyyy-MM-dd-HH-mm).ToString()).csv"
        $Result=""
        $Results=@()
        $Counter = 0
        $ListMailbox = Get-mailbox -RecipientTypeDetails UserMailbox,SharedMailbox -ResultSize Unlimited
        $totalmailboxes = $ListMailbox.Count
        $ListMailbox | ForEach-Object{
            $upn=$_.UserPrincipalName
            $DisplayName=$_.Displayname
            $PrimarySMTPAddress=$_.PrimarySMTPAddress
            
            $Counter++
            # Calculate percentage completion
            $percentComplete = [math]::Round(($Counter / $totalmailboxes) * 100)
            # Define progress bar parameters with user principal name
            $progressParams = @{
                Activity        = "Processing Mailboxes"
                Status          = "Mailbox $($counter) of $totalMailboxes - $($DisplayName) - $percentComplete% Complete"
                PercentComplete = $percentComplete
            }
            Write-Progress @progressParams
            
            #Getting delegated Fullaccess permission for mailbox
            $FullAccessPermissions=(Get-MailboxPermission -Identity $upn | Where-Object { ($_.AccessRights -contains "FullAccess") -and ($_.IsInherited -eq $false) -and -not ($_.User -match "NT AUTHORITY" -or $_.User -match "S-1-5-21") }).User
            if([string]$FullAccessPermissions -ne ""){
                $UserWithAccess=""
                $AccessType="FullAccess"
                foreach($FullAccessPermission in $FullAccessPermissions){
                if($UserWithAccess -ne ""){
                    $UserWithAccess=$UserWithAccess+","
                }
                $UserWithAccess=$UserWithAccess+$FullAccessPermission
                }
                $Result = @{'Display Name'=$_.Displayname;'User PrinciPal Name'=$upn;'Primary SMTP Address'=$PrimarySMTPAddress;'Access Type'=$AccessType;'User With Access'=$userwithAccess}
                $Results = New-Object PSObject -Property $Result
                $Results | select-object 'Display Name','User PrinciPal Name','Primary SMTP Address','Access Type','User With Access' | Export-Csv -Path $ExportCSV -Notype -Append
            }
            
            #Getting delegated SendAs permission for mailbox
            $SendAsPermissions=(Get-RecipientPermission -Identity $upn | Where-Object{ -not (($_.Trustee -match "NT AUTHORITY") -or ($_.Trustee -match "S-1-5-21"))}).Trustee
            if([string]$SendAsPermissions -ne ""){
                $UserWithAccess=""
                $AccessType="SendAs"
                foreach($SendAsPermission in $SendAsPermissions){
                if($UserWithAccess -ne ""){
                    $UserWithAccess=$UserWithAccess+","
                }
                $UserWithAccess=$UserWithAccess+$SendAsPermission
                }
                $Result = @{'Display Name'=$_.Displayname;'User PrinciPal Name'=$upn;'Primary SMTP Address'=$PrimarySMTPAddress;'Access Type'=$AccessType;'User With Access'=$userwithAccess}
                $Results = New-Object PSObject -Property $Result
                $Results |select-object 'Display Name','User PrinciPal Name','Primary SMTP Address','Access Type','User With Access' | Export-Csv -Path $ExportCSV -Notype -Append
            }
            
            #Getting delegated SendOnBehalf permission for mailbox
            $SendOnBehalfPermissions=$_.GrantSendOnBehalfTo
            if([string]$SendOnBehalfPermissions -ne ""){
                $UserWithAccess=""
                $AccessType="SendOnBehalf"
                foreach($SendOnBehalfPermissionDN in $SendOnBehalfPermissions){
                if($UserWithAccess -ne ""){
                    $UserWithAccess=$UserWithAccess+","
                }
                #$SendOnBehalfPermission=(Get-Mailbox -Identity $SendOnBehalfPermissionDN).UserPrincipalName
                $UserWithAccess=$UserWithAccess+$SendOnBehalfPermissionDN
            }
                $Result = @{'Display Name'=$_.Displayname;'User PrinciPal Name'=$upn;'Primary SMTP Address'=$PrimarySMTPAddress;'Access Type'=$AccessType;'User With Access'=$userwithAccess}
                $Results = New-Object PSObject -Property $Result
                $Results |select-object 'Display Name','User PrinciPal Name','Primary SMTP Address','Access Type','User With Access' | Export-Csv -Path $ExportCSV -Notype -Append
            }
            Start-Sleep -Milliseconds 100
        }
    }
    Write-Progress -Activity "Processing Mailboxes" -Completed
    Write-Host "Script completed. Results exported to  $OutputCSV." -ForegroundColor Cyan
}
function ExchangeOnlineListPermissionsByUser {
    Write-Host "Enter email address to check FullAccess rights" -ForegroundColor Yellow
    $ExchangeUser = Read-Host "Username"
    DisplayBug
    Write-Host "####################################################" -ForegroundColor Green
    Get-Mailbox -ResultSize Unlimited | Get-MailboxPermission -User $ExchangeUser -ErrorAction SilentlyContinue | Sort-Object -Property Identity | format-table Identity ,AccessRights, User -AutoSize
    Write-Host "####################################################" -ForegroundColor Green
    PromptPressKeyToContinue

}
function ExchangeOnlineListMailboxSize {
    $OutputCSV="$programdir\MailboxSizeReport_$((Get-Date -format yyyy-MM-dd-HH-mm).ToString()).csv" 
    $Result=""   
    $Results=@()  
    $MBCounter=0
    $ListMailboxSize = get-mailbox -ResultSize Unlimited | Get-MailboxStatistics -ErrorAction SilentlyContinue | Select-Object DisplayName, MailboxTypeDetail, ItemCount, TotalItemSize, SystemMessageSizeShutoffQuota | Format-Table -AutoSize
    Start-Sleep -Milliseconds 250
    DisplayBug
    Write-Host "####################################################" -ForegroundColor Green
    Write-Output $ListMailboxSize
    Write-Host "####################################################" -ForegroundColor Green
    Start-Sleep -Seconds 3
    PromptExportToCSV
    if ($script:ExportCSV -eq 1){
        Get-Mailbox -ResultSize Unlimited | ForEach-Object {
        $UPN=$_.UserPrincipalName
        $Mailboxtype=$_.RecipientTypeDetails
        $DisplayName=$_.DisplayName
        $PrimarySMTPAddress=$_.PrimarySMTPAddress
        $IssueWarningQuota=$_.IssueWarningQuota -replace "\(.*",""
        $ProhibitSendQuota=$_.ProhibitSendQuota -replace "\(.*",""
        $ProhibitSendReceiveQuota=$_.ProhibitSendReceiveQuota -replace "\(.*",""

        $totalmailboxes = $ListMailboxSize.Count
        #Get all Email addresses in Microsoft 365
            $MBCounter++
            # Calculate percentage completion
            $percentComplete = [math]::Round(($MBcounter / $totalmailboxes) * 100)

            # Define progress bar parameters with user principal name
            $progressParams = @{
                Activity        = "Processing Mailboxes"
                Status          = "Mailbox $($MBcounter) of $totalMailboxes - $($DisplayName) - $percentComplete% Complete"
                PercentComplete = $percentComplete
            }

            Write-Progress @progressParams

        if($SharedMBOnly.IsPresent -and ($Mailboxtype -ne "SharedMailbox")){
            return }
        if($UserMBOnly.IsPresent -and ($MailboxType -ne "UserMailbox")){
            return}  
        #Check for archive enabled mailbox
        if(($null -eq $_.ArchiveDatabase) -and ($_.ArchiveDatabaseGuid -eq $_.ArchiveGuid)){
            $ArchiveStatus = "Disabled"}
        else{$ArchiveStatus= "Active"}

        $Stats=Get-MailboxStatistics -Identity $UPN
        $ItemCount=$Stats.ItemCount
        $TotalItemSize=$Stats.TotalItemSize
        $TotalItemSizeinBytes= $TotalItemSize –replace “(.*\()|,| [a-z]*\)”, “”
        $TotalSize=$stats.TotalItemSize.value -replace "\(.*",""
        $DeletedItemCount=$Stats.DeletedItemCount
        $TotalDeletedItemSize=$Stats.TotalDeletedItemSize
        
        #Export result to csv
        $Result=@{'Display Name'=$DisplayName;'User Principal Name'=$upn;'Mailbox Type'=$MailboxType;'Primary SMTP Address'=$PrimarySMTPAddress;'Archive Status'=$Archivestatus;'Item Count'=$ItemCount;'Total Size'=$TotalSize;'Total Size (Bytes)'=$TotalItemSizeinBytes;'Deleted Item Count'=$DeletedItemCount;'Deleted Item Size'=$TotalDeletedItemSize;'Issue Warning Quota'=$IssueWarningQuota;'Prohibit Send Quota'=$ProhibitSendQuota;'Prohibit send Receive Quota'=$ProhibitSendReceiveQuota}
        $Results= New-Object PSObject -Property $Result  
        $Results | Select-Object 'Display Name','User Principal Name','Mailbox Type','Primary SMTP Address','Item Count','Total Size','Total Size (Bytes)','Archive Status','Deleted Item Count','Deleted Item Size','Issue Warning Quota','Prohibit Send Quota','Prohibit Send Receive Quota' | Export-Csv -Path $OutputCSV -Notype -Append 

        Start-Sleep -Milliseconds 100
        }
        Write-Progress -Activity "Processing Mailboxes" -Completed
        Write-Host "Script completed. Results exported to  $OutputCSV." -ForegroundColor Cyan
    }
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
            Invoke-Expression $ExchangeOnlineCustomCmd
            Write-Host ""
        }       
    }    
}
####################################

####################################
#  Sharepoint Online  #
#-----------------------
function Show-SharePointOnlineMenu {
    Write-Host "######## Microsoft SharePoint Online Menu ########" -ForegroundColor DarkCyan
    Write-Host "SharePoint Online Connection : " -NoNewline -ForegroundColor Yellow
    Write-Host "$SharePointConnectionStatus" -ForegroundColor $SharePointConnectionStatusColor
    Write-Host "`n(1) - Install Required Modules"
    Write-Host "(2) - Connect / Disconnect Sharepoint Online"
    Write-Host "(3) - List SharePoint Sites"
    Write-Host "(8) - Custom Command - Coming Soon" -ForegroundColor Red
    Write-Host "`n(9) - Main Menu"
    Write-Host "################################################" -ForegroundColor DarkCyan
    Write-Host ""
    Import-Module PnP.PowerShell -ErrorAction SilentlyContinue
}
function installSharepointOnlineModule {
    if(-not (Get-Module PnP.PowerShell -ListAvailable)){
        Write-Warning "PnP.PowerShell not installed"
        Write-Host "PnP.PowerShell Please Wait" -ForegroundColor Green
        Install-Module PnP.PowerShell -Scope CurrentUser -Force -Confirm:$false
        }
    Start-Sleep -Seconds 1
    Write-Host "Module Already Installed" -ForegroundColor Green
    Start-Sleep -Seconds 2
}
function SharePointOnlineConnection {
    try {
        Get-PnPConnection | Out-Null
        Write-Host "Disconnecting Microsoft 365 - PnP Online" -ForegroundColor Yellow
        Start-Sleep -Seconds 2
        Disconnect-PnPOnline
    } catch {
        Write-Host "Connecting to Microsoft 365 - PnP Online" -ForegroundColor Yellow
        Write-Host "Enter Organisation Name (First part of xxxx.onmicrosoft.com)"
        $OrgName = Read-Host "Name"
        $ClientID = Register-PnPEntraIDAppForInteractiveLogin -ApplicationName "PnP-MS365-PSToolkit-$((Get-Date -format yyyyMMddhhmm).ToString())" -Tenant "$OrgName.onmicrosoft.com" -Interactive -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
        Connect-PnPOnline -url "$OrgName.sharepoint.com" -Interactive -ClientId $clientID."AzureAppId/ClientId"

    }
}
function SharePointOnlineListSites {
    $ListSharePointSites = Get-PnPTenantSite | Format-Table url,status,archivestatus -autosize
    Start-Sleep -Milliseconds 250
    DisplayBug
    Write-Host "####################################################" -ForegroundColor Green
    Write-Output $ListSharePointSites
    Write-Host "####################################################" -ForegroundColor Green
    PromptPressKeyToContinue
}

####################################

#####################################################################
# Main Code --- Main Code --- Main Code --- Main Code --- Main Code #
#-------------------------------------------------------------------#
SplashLogo
Start-Sleep -Seconds 3
Clear-Host
start-sleep -Milliseconds 250
Logo
Start-Sleep -Seconds 1
Set-ExecutionPolicy Bypass -Force -Scope Process -Confirm:$false
#####################################################################

#####################################################################

######################
# Requirements Check #
######################
if (!(Test-Path $programdir)) {New-Item -itemType Directory -Path $programdir | Out-Null}
CheckPowershellVersion
CheckAdminPrivs
CheckForUpdates
Get-MgApplication -ErrorAction SilentlyContinue | Out-Null
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
                    51 {try {MsGraphCreateMFAStatusReport}
                        catch {Write-Error "Error Running Script"
                            CatchError}}
                    52 {try {MsGraphCreate365LicenseReport}
                        catch {Write-Error "Error Running Script"
                            CatchError}}
                    53 {try {MsGraphShowUsersWithAndWithoutLicense}
                        catch {Write-Error "Error Running Script"
                            CatchError}}
                    55 {}
                    56 {$WhileLoopVarMsGraphToolsMenu = 1
                        while ($WhileLoopVarMsGraphToolsMenu -eq 1){
                            Clear-Host
                            Start-Sleep -Milliseconds 250
                            Logo
                            Start-Sleep -Milliseconds 250
                            Show-MsGraphToolsMenu
                            $MsGraphToolsMenuChoice = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown").VirtualKeyCode
                            switch ($MsGraphToolsMenuChoice) {
                                49 {try {MsGraphRemoveObsoleteMS365ToolKitEntraApplication}
                                    catch {CatchError}}
                                50 {try {MsGraphForcePasswordResetSingleUser}
                                    catch {Write-Error "Error Running Script"
                                        CatchError}}
                                51 {try {MsGraphForcePasswordResetAllUsers}
                                    catch {Write-Error "Error Running Script"
                                        CatchError}}
                                57 {Write-Host "`nReturning to Microsoft Graph Menu" -ForegroundColor Cyan
                                $WhileLoopVarMsGraphToolsMenu = 0}
                                
                                default {Write-Host "`nInvalid selection, please try again." -ForegroundColor Red
                                    Start-Sleep -Seconds 1}
                            }
                        }
                    }
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
                $ExchangeOnlineConnectionStatus = "Disconnected"
                $ExchangeOnlineConnectionStatusColor = "Red"
                if (Get-Module ExchangeOnlineManagement -ListAvailable){
                    $ExchangeOnlineConnectionState = Get-ConnectionInformation
                    If ( $ExchangeOnlineConnectionState.State -match 'Connected' ){
                        $ExchangeOnlineConnectionStatus = "Connected"
                        $ExchangeOnlineConnectionStatusColor = "Green"
                    }
                    else {
                        $ExchangeOnlineConnectionStatus = "Disconnected"
                        $ExchangeOnlineConnectionStatusColor = "Red"
                    }
                }
                Show-ExchangeOnlineMenu
                $ExchangeOnlineMenuChoice = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown").VirtualKeyCode
                switch ($ExchangeOnlineMenuChoice) {
                    49 {try {installExchangeOnlineModule}
                        catch {Write-Error "Error Running Module Installer"
                            CatchError}}
                    50 {try {ExchangeOnlineConnection}
                        catch {Write-Error "Error while connecting"
                            CatchError}}
                    51 {try {ExchangeOnlineListMailbox}
                        catch {Write-Error "Error Running Script"
                            CatchError}}
                    52 {try{ExchangeOnlineListPermissions}
                        catch {Write-Error "Error Running Script"
                            CatchError}}
                    53 {try{ExchangeOnlineListPermissionsByUser}
                        catch {Write-Error "Error Running Script"
                            CatchError}}
                    54 {try {ExchangeOnlineListMailboxSize}
                        catch {Write-Error "Error Running Script"
                            CatchError}}
                    55 {
                        $WhileLoopVarExchangeOnlineToolsMenu = 1
                        while ($WhileLoopVarExchangeOnlineToolsMenu -eq 1){
                            Clear-Host
                            Start-Sleep -Milliseconds 250
                            Logo
                            Start-Sleep -Milliseconds 250
                            Show-ExchangeOnlineToolsMenu
                            $ExchangeOnlineToolsMenuChoice = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown").VirtualKeyCode
                            switch ($ExchangeOnlineToolsMenuChoice) {
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

        51 {$WhileLoopVarSharepointMenu = 1
            while ($WhileLoopVarSharepointMenu -eq 1){
                Clear-Host
                Start-Sleep -Milliseconds 250
                Logo
                Start-Sleep -Milliseconds 250
                $SharePointConnectionStatus = "Disconnected"
                $SharePointConnectionStatusColor = "Red"
                if (Get-Module PnP.Powershell -ListAvailable){
                    try{ Get-PnPConnection | Out-Null
                        $SharePointConnectionStatus = "Connected"
                        $SharePointConnectionStatusColor = "Green" 
                    }
                    
                    catch {
                        $SharePointConnectionStatus = "Disconnected"
                        $SharePointConnectionStatusColor = "Red"
                    }
                }
                Show-SharepointOnlineMenu
                $SharepointMenuChoice = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown").VirtualKeyCode
                switch ($SharepointMenuChoice) {
                    49 {try {installSharepointOnlineModule}
                        catch {Write-Error "Error Running Module Installer"
                            CatchError}}
                    50 {try {SharePointOnlineConnection}
                        catch {Write-Error "Error Running Script"
                            CatchError}}
                    51 {try {SharePointOnlineListSites}
                        catch {Write-Error "Error Running Script"
                            CatchError}}
                    52 {try {}
                        catch {Write-Error "Error Running Script"
                            CatchError}}
                    53 {try {}
                        catch {Write-Error "Error Running Script"
                            CatchError}}
                    56 {}
                    57 {
                        Write-Host "`nReturning to Main Menu" -ForegroundColor Cyan
                        $WhileLoopVarSharepointMenu = 0}
                        
                        default {Write-Host "`nInvalid selection, please try again." -ForegroundColor Red
                            Start-Sleep -Seconds 1}
                }
            }
        }
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