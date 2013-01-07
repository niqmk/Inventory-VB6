Attribute VB_Name = "mdlGlobal"
Option Explicit

Public Enum ReminderType
    NoneType
    FromMaster
    FromDate
    FromFirstDate
    FromTransaction
End Enum

Public Enum ValidateType
    NoneValidate
    OnceMonth
    MonthSequence
    DaySequence
End Enum

Public Enum FunctionMode
    ViewMode
    AddMode
    UpdateMode
    DeleteMode
    PrintMode
End Enum

Public Const curHeavy As Currency = 0.52
Public Const curNonHeavy As Currency = 0.4

Public Const intPort As Integer = 13300

Public Const PUBLIC_KEY As String = "ADMINIST"

Public Const SERVER_REGISTRY As String = "Server"
Public Const USERID_REGISTRY As String = "User_Id"
Public Const PASSWORD_REGISTRY As String = "Password"
Public Const DATABASE_REGISTRY As String = "Database"
Public Const COMPANY_REGISTRY As String = "Company"
Public Const ADDRESS_REGISTRY As String = "Address"
Public Const WEBSITE_REGISTRY As String = "Website"
Public Const EMAIL_REGISTRY As String = "Email"
Public Const PHONE_REGISTRY As String = "Phone"
Public Const FAX_REGISTRY As String = "Fax"
Public Const NPWP_REGISTRY As String = "NPWP"
Public Const LOGO_REGISTRY As String = "Logo"
Public Const WALLPAPER_REGISTRY As String = "Wallpaper"
Public Const COMMPORT_REGISTRY As String = "CommPort"
Public Const AUTORUN_REGISTRY As String = "AutoRun_Inventory"

Public Const LOGO_IMAGE_FILE As String = "logo"
Public Const WALLPAPER_IMAGE_FILE As String = "wallpaper"

Public Const strInventory As String = "INVENTORY"
Public Const strFinance As String = "FINANCE"
Public Const strAccounting As String = "ACCOUNTING"
Public Const strAdministrator As String = "ADMINISTRATOR"
Public Const strUser As String = "USER"

Public Const strTemplateFolder As String = "Template"

Public Const strYes As String = "Y"
Public Const strNo As String = "N"

Public Const strHeavy As String = "H"
Public Const strNonHeavy As String = "NH"

Public objDatabaseInit As SQLDATABASE

Public UserAuthority As clsAuthority

Public fso As FileSystemObject

Public conInventory As ADODB.Connection
Public conAccounting As ADODB.Connection
Public conFinance As ADODB.Connection

Public strServerInit As String
Public strUserIdInit As String
Public strPasswordInit As String

Public strPath As String
Public strFormatDate As String

Public strCompanyText As String
Public strAddressText As String
Public strWebsiteText As String
Public strEmailText As String
Public strPhoneText As String
Public strNPWPText As String
Public strFaxText As String
Public strLogoImageText As String
Public strWallpaperImageText As String

Public blnFill As Boolean
Public blnChat As Boolean
