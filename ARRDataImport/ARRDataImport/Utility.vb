Imports BEA.TFS.Common.ExtensionMethods
Imports BEA.TFS.CommonUtility
Imports log4net
Imports System.Data.OleDb
Imports System.IO


Module Utility
    Private ReadOnly _logger As ILog = LogManager.GetLogger(GetType(Utility))
    '''' <summary>
    '''' Get the month end date from report DB.
    '''' </summary>
    '''' <returns>Get the from LASTCALENDARDATE from GDR_RUN_DATE as month end date.</returns>
    'Friend Function GetMonthEndDate(conn As IDbConnection, Optional tran As IDbTransaction = Nothing) As Date

    '    Dim monthEndDate As Date = Nothing
    '    Using cmd As IDbCommand = conn.CreateCommand,
    '                rs As IDataReader = cmd.ExecuteReader(
    '                    AppConst.SELECT_LAST_CALENDAR_DATE,
    '                    New IDbDataParameter() {},
    '                    tran
    '                )
    '        If rs.Read() Then
    '            monthEndDate = rs.GetDateTime("LASTCALENDARDATE")
    '        Else
    '            _errorMessageList.Add("Error occurred while reading LASTCALENDARDATE from GDR_RUN_DATE. Table is empty.")
    '        End If
    '    End Using

    '    Return monthEndDate
    'End Function

    ''' <summary>
    ''' Housekeep and import GDR ARR data.
    ''' </summary>
    ''' <param name="config"></param>
    ''' <param name="conn"></param>
    ''' <param name="workingDate"></param>
    Friend Sub ImportARR(config As INIConfig, conn As OleDbConnection, workingDate As Date)
        Using ARRTran As OleDbTransaction = conn.BeginTransaction()

            _logger.Debug($"Deleting data from table ARR_RATE...")
            Utility.HouseKeepARRTable(conn, ARRTran)

            _logger.Info($"Importing ARR Files...")
            Try
                Dim isImportARRSuccessful As Boolean = Utility.ImportARRFiles(config, workingDate, conn, ARRTran)
                If Not isImportARRSuccessful Then

                    _logger.Error($"Import was not successful due to file not found or error has occurred.")
                End If

                ARRTran.Commit()
            Catch ex As Exception
                ARRTran.Rollback()
                Dim errorMessage As String = "Error occurred while importing ARR files."
                _errorMessageList.Add(errorMessage)
                _errorMessageList.Add(ex.Message)
            End Try
        End Using
    End Sub

    ''' <summary>
    ''' Import ARR
    ''' </summary>
    ''' <param name="config">Config read from INI file.</param>
    ''' <param name="workingDate">GDR_RUN_DATE.RUN_DATE.</param>
    ''' <param name="conn">Opened connection to DB.</param>
    ''' <param name="tran">Transaction to DB of the <paramref name="conn"/>.</param>
    ''' <returns>Returns true only if no error has occurred while loading files into DB.</returns>
    Friend Function ImportARRFiles(config As INIConfig, workingDate As Date, conn As IDbConnection, Optional tran As IDbTransaction = Nothing) As Boolean

        Dim isLoadARRSuccess As Boolean = False



        'INSERT DATA OF ARR_RATE = ARR FILE
        Dim ARRFilePath As String = Path.Combine(config.SFTPINFOLDER, config.FILE_ARR)
        _logger.Info($"Finding {ARRFilePath}")

        If File.Exists(ARRFilePath) Then
            isLoadARRSuccess = LoadARRFile(ARRFilePath, workingDate, conn, tran)
        Else
            _logger.Info("File of ARR not exists, skip to load.")
            isLoadARRSuccess = False
        End If

        Return isLoadARRSuccess
    End Function

    Friend Sub HouseKeepARRTable(conn As IDbConnection, Optional tran As IDbTransaction = Nothing)

        ' Clear records from GDR_CRC where LastCalendarDate = GDR_RUN_DATE.LastCalendarDate
        Using cmd As IDbCommand = conn.CreateCommand
            cmd.ExecuteNonQuery(
                AppConst.DELETE_FROM_ARR,
                New IDbDataParameter() {},
                tran
            )
        End Using
    End Sub

    Friend Function LoadARRFile(ARRFilePath As String, workingDate As Date, conn As IDbConnection, Optional tran As IDbTransaction = Nothing) As Boolean

        Try

            ' Import file 'ARR' into table ARR_RATE
            Using sr As New StreamReader(ARRFilePath)
                While Not sr.EndOfStream()
                    Dim line As String = sr.ReadLine()

                    If line.Substring(0, 2) = "50" Then

                        ' Read the body line into file object
                        Dim ARRFileObject As ARRFile = ARRFile.ReadLine(line)
                        ARRFileObject.RUN_DATE = workingDate
                        ' Import file object into database
                        GDR_ARR.ImportARRObject(ARRFileObject, conn, tran)
                    End If

                End While
                _logger.Info("File of ARR loaded.")
            End Using

            Return True

        Catch ex As Exception
            _logger.Error($"Error occured when importing file ARR to table ARR_RATE.", ex)
            _errorMessageList.Add("Error occured when importing file ARR to table ARR_RATE.")
            _errorMessageList.Add(ex.Message)
            Return False
        End Try
    End Function



    Friend Sub SendErrorEmail(emailConfig As EmailConfig, config As INIConfig, errorMessages As IEnumerable(Of String))
        Try
            Dim emailTo() As String = config.EmailToErr?.ToString.Split({";"c}, StringSplitOptions.RemoveEmptyEntries)
            Dim emailCc() As String = config.EmailCCToErr?.ToString.Split({";"c}, StringSplitOptions.RemoveEmptyEntries)
            Dim emailBody As String = String.Join(vbNewLine, errorMessages.ToArray())

            SendEmail(
                server:=emailConfig.Server,
                port:=emailConfig.PortNo,
                [from]:=emailConfig.MailFrom,
                to:=emailTo,
                cc:=emailCc,
                subject:=ResolveFileName(config.EmailSubjectErr),
                body:=emailBody
            )
        Catch ex As Exception
            _logger.Error($"Error occured while sending error notification email.", ex)
            For Each errMessage In errorMessages
                _logger.Error(errMessage)
            Next
            ' Explicitly kill the program if error email could not be sent
            Throw New Exception($"Error occured while sending error notification email.", ex)
        End Try
    End Sub
End Module
