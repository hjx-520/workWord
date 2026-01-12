Imports log4net
Imports System.Data.OleDb
Imports System.IO
Imports BEA.TFS.Common.Entities
Imports BEA.TFS.Common.ExtensionMethods
Imports BEA.TFS.CommonUtility

Public Module ARRDataImport

    Private ReadOnly _logger As ILog = LogManager.GetLogger(GetType(ARRDataImport))
    ''' <summary>
    ''' This list mimics TI440 behaviour where all functions will try-catch every exception, log the error message and continue processing with best effort. 
    ''' This list will be send in an email at the end of this program.
    ''' </summary>
    Friend _errorMessageList As New List(Of String)
    Public Function Main(Optional args As String() = Nothing) As Integer
        Try
            _logger.Info($"Start {NameOf(ARRDataImport)}...")
            _logger.Info("Initializing...")

            Dim dsINIFile As String = Path.GetFullPath(My.Settings.DATA_SOURCE_INI_FILE)
            Dim iniFile As String = Path.GetFullPath(My.Settings.INI_FILE)
            Dim emailINIFile As String = Path.GetFullPath(My.Settings.EMAIL_INI_FILE)
            _logger.Info("Reading datasource config...")
            Dim dbConfig As DataSourceConfig = GetDataSourceConfig(dsINIFile)
            _logger.Info("Reading email config...")
            Dim emailConfig As EmailConfig = GetEmailConfig(emailINIFile)
            _logger.Info("Reading program config...")
            Dim config As INIConfig = GetConfig(iniFile)
            Dim rptConnStr As String = GetConnectionStr(dbConfig.RptDSN, dbConfig.RptUserID, dbConfig.RptPwd)

            Using conn As New OleDbConnection(rptConnStr)
                conn.Open()
                ' Do not start transaction here, start transaction for each module and commit on success/rollback on error
                Try
                    ' Get working date
                    Dim rptRunDate As RptEntities.RptRunDate = RptEntities.RptRunDate.Find(conn, Nothing)
                    TryLogPropertyValues(rptRunDate)

                    Dim workingDate As Date = rptRunDate.RunDate.Value
                    Dim nextWorkingDate As Date = rptRunDate.NextBusinessDate.Value

                    ' Note: Using ResolveFileName() will remove all '<' or '>'. Perform manual replace first.
                    ' For detailed test, refer To case 2 of unit test "Test_ResolveFileName_Resolving_Two_Different_Dates()"
                    config.FILE_ARR = config.FILE_ARR.Replace("<yyMMdd_end>", workingDate.ToString("yyMMdd"))
                    config.FILE_ARR = ResolveFileName(config.FILE_ARR, workingDate)


                    _logger.Debug($"Completed updating date element in files, logging file names...")
                    config.TryLogPropertyValues()


                    ' ARR
                    ImportARR(config, conn, workingDate)

                    _logger.Info($"End {NameOf(ARRDataImport)}...")


                    'Delibrately not roll back due to each module has different behaviour.
                Finally

                    ' This is equivalent to the logic "If (error_flag) Then"
                    ' If any error occurred, send email with error messages in email body.
                    If _errorMessageList.Count > 0 Then
                        _logger.Info($"Sending error notification email...")
                        SendErrorEmail(emailConfig, config, _errorMessageList)
                        _logger.Info($"Error notification sent.")
                    End If
                End Try
            End Using


        Catch ex As Exception
            _logger.Error($"Error occured in {NameOf(ARRDataImport)} : ", ex)
            Return -1
        End Try

        Return 0

    End Function

    Public Function GetConfig(iniFile As String) As INIConfig
        Dim config As New INIConfig With {
            .EmailTo = GetProfileStringVal("ARRDataImport", "EmailTo", iniFile),
            .EmailCCTo = GetProfileStringVal("ARRDataImport", "EmailCCTo", iniFile),
            .EmailSubject = GetProfileStringVal("ARRDataImport", "EmailSubject", iniFile),
            .EmailToErr = GetProfileStringVal("ARRDataImport", "EmailToErr", iniFile),
            .EmailCCToErr = GetProfileStringVal("ARRDataImport", "EmailCCToErr", iniFile),
            .EmailSubjectErr = GetProfileStringVal("ARRDataImport", "EmailSubjectErr", iniFile),
            .FILE_ARR = GetProfileStringVal("ARRDataImport", "FILE_ARR", iniFile),
            .SFTPINFOLDER = GetProfileStringVal("ARRDataImport", "SFTPINFOLDER", iniFile),
            .SFTPOUTFOLDER = GetProfileStringVal("ARRDataImport", "SFTPOUTFOLDER", iniFile)
        }

        Return config
    End Function
End Module
