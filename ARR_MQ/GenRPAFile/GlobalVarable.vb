Public Class GlobalVarable

    Public Const MAX_DB_CONNECTION_ATTEMPTS As Byte = 10
    Public Const MAX_MQ_CONNECTION_ATTEMPTS As Byte = 10
    Public Const MQ_CHAR_SET As Integer = 437
    Public Const MQ_Expire_Seconds As Integer = 7200

    Public Shared AppPath As String
    Public Shared ErrLogFile As String
    'Public Shared FileNum As Integer
    Public Shared INI_App As String
    Public Shared INI_DataSource As String
    Public Shared INI_Email As String
    Public Shared INI_ErrCode As String
    Public Shared AppName As String = "GenRPAFile"
    Public Shared RunDate As String
    Public Shared Output_File_Path As String
    Public Shared TXT_File_Name As String
    Public Shared CSV_File_Name As String
    Public Shared Header_Trailer_TXT_Flag As String = "Y"
    Public Shared Header_Trailer_CSV_Flag As String = "N"

    Public Shared CSV_DELIMTER As String
    Public Shared CSV_Col_Header As String
    Public Shared CSV_Col_Header_Caption As String()

    Public Shared FILE_FORMAT_CSV As String
    Public Shared FILE_FORMAT_TXT As String
    Public Shared TXT_Col_Width As String()
    Public Shared TXT_Col_Type As String()
    Public Shared TXT_FILLER_WIDTH As Integer


    Public Shared TXT_File_Header As String = ""
    Public Shared TXT_File_Trailer As String = ""
    Public Shared CSV_File_Header As String = ""
    Public Shared CSV_File_Trailer As String = ""

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
    Public Shared Attachment As String

    Public Shared ProgramStatus As Boolean = False
    Public Shared DATA_RETENTION_PERIOD As Integer
    Public Shared DATA_FILENAME_DATE_FORMAT_S As String
    Public Shared DATA_FILENAME_DATE_FORMAT_L As String

    Public Shared MQ_SERVER As String
    Public Shared MQ_CHANNEL As String
    Public Shared MQ_PORT As String
    Public Shared MQ_Manager As String
    Public Shared MQ_QUEUE_REQ As String
    Public Shared MQ_QUEUE_RPLY As String
    Public Shared MQ_TRANS_TYPE As String
    Public Shared MQ_MSG_Expiry As Integer
    Public Shared MQ_Wait_Interval As Integer

End Class
