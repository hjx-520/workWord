Public Class GlobalVarable

    Public Shared AppPath As String
    Public Shared INI_DataFormat As String
    Public Shared ErrLogFile As String
    Public Shared FileNum As Integer
    Public Shared INI_App As String
    Public Shared INI_DataSource As String
    Public Shared INI_Email As String
    Public Shared INI_ErrCode As String
    Public Shared AppName As String = "DataImport"
    Public Shared RunDate As String
    Public Shared Prev_Bus_Date As String
    Public Shared Current_Run_Date As String
    Public Shared DataFilePath As String

    Public Shared MailEnable As String
    Public Shared MailServer As String
    Public Shared MailServerPort As String
    Public Shared MailFrom As String
    Public Shared SUCCESSMailToList As String
    Public Shared SUCCESSMailCCList As String
    Public Shared ERRMailToList As String
    Public Shared ERRMailCCList As String
    Public Shared MailSubject As String
    Public Shared MailErrSubject As String
    Public Shared MailBody As String

    Public Shared ProgramStatus As Boolean = False
    Public Shared DATA_RETENTION_PERIOD As Integer
    Public Shared DATA_FILENAME_DATE_FORMAT As String

End Class
