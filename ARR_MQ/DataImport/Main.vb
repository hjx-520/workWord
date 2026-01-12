Imports System
Imports System.IO
Imports System.Configuration
Imports System.Text
Imports System.Data
Imports System.Data.OleDb
Imports System.Net.Mail


Module Main

    Sub Main()

        Dim myResult As Boolean = False
        Dim myLogger As New Logger
        Dim myFieldList As New ImportFields
        Dim myTableList As New ImportTables
        Dim aArgs() As String

        Dim myINIFile, myErrINI, myDSINI, myEmailINI As IniFile
        Dim myDB As New DBProcOLEDB
        Dim config As DataSourceConfig


        Dim SecName As String = ""
        Dim KeyName As String = ""
        Dim ColPrefix As String = ""
        Dim FileName, TableName, FileHeader, FileTrailer As String
        Dim FieldNum As Integer = 0
        Dim i, j, k As Integer
        Dim ErrMsg As String = ""
        Dim tmpFileName As String = ""
        Dim tmpStr As String = ""
        Dim ErrEmailContent As String = ""
        Dim EmailContent As String = ""

        GlobalVarable.ProgramStatus = True

        Try
            GlobalVarable.AppPath = My.Application.Info.DirectoryPath & "\"
            'GlobalVarable.AppPath = "D:\Development\ARR\"

            'GlobalVarable.INI_DataFormat = ConfigurationManager.AppSettings("INI_DataFormat")
            GlobalVarable.ErrLogFile = GlobalVarable.AppPath & "DataImport" & ".Log"
            GlobalVarable.INI_DataSource = ConfigurationManager.AppSettings("INI_DataSource")
            GlobalVarable.INI_Email = ConfigurationManager.AppSettings("INI_Email")
            GlobalVarable.INI_ErrCode = GlobalVarable.AppPath & ConfigurationManager.AppSettings("INI_ErrCode")

            'Setup and assign values to variables
            'argument checking
            aArgs = System.Environment.GetCommandLineArgs
            'aArgs = {"1", "ARRDataFormat.ini"}
            'aArgs = {"1", "TERMDataFormat.ini"}
            If aArgs.Count < 2 Then
                myResult = WriteToLog("NOSECNAME", "Missing or Invalid parameter!")
                myResult = WriteToLog("NOSECNAME", "Example: DataImport.exe ARRDataFormat.ini")
                End
                Exit Sub
            End If

            'GlobalVarable.INI_App = Path.GetFullPath(aArgs(1))
            GlobalVarable.INI_App = GlobalVarable.AppPath & aArgs(1)

            Console.WriteLine(GlobalVarable.INI_App)
            Console.WriteLine(GlobalVarable.AppPath)

            myINIFile = New IniFile(GlobalVarable.INI_App)
            'myINIFile = New IniFile(GlobalVarable.AppPath & ConfigurationManager.AppSettings("INI_App"))

            myErrINI = New IniFile(GlobalVarable.INI_ErrCode)
            config = New DataSourceConfig()

        Catch ex As Exception
            myLogger.AppendToFile(GlobalVarable.ErrLogFile, myLogger.FormatString(True, GlobalVarable.AppName, "SetGlobalVariable", "SetGlobalVariable error, please check!"))
            Exit Sub
        End Try

        myResult = WriteToLog("GENERAL", "ERR00100101")

        If Not File.Exists(GlobalVarable.INI_App) Then
            myResult = WriteToLog("INITIALIZATION", "ERR00200104")
        Else
            'Setup a Master and Sub List to store the Import Data Structure
            myResult = WriteToLog("INITIALIZATION", "ERR00200105")

            tmpStr = ""
            tmpStr = Trim(myINIFile.GetString("GENERAL", "FILE_NO", ""))
            If IsNumeric(tmpStr) Then
                GlobalVarable.FileNum = Int(Val(tmpStr))
            Else
                GlobalVarable.FileNum = 0
            End If

            tmpStr = ""
            tmpStr = IIf(Right(Trim(myINIFile.GetString("GENERAL", "DATA_FILE_PATH", "")), 1) = "\", Trim(myINIFile.GetString("GENERAL", "DATA_FILE_PATH", "")), Trim(myINIFile.GetString("GENERAL", "DATA_FILE_PATH", "")) & "\")
            If System.IO.Directory.Exists(tmpStr) Then
                GlobalVarable.DataFilePath = tmpStr
            Else
                myResult = WriteToLog("NOSECNAME", "DATA_FILE_PATH not exist, please check!")
                Exit Sub
            End If

            tmpStr = ""
            tmpStr = IIf(Right(Trim(myINIFile.GetString("GENERAL", "DATA_FILENAME_DATE_FORMAT", "")), 1) = "", "yyMMdd", Trim(myINIFile.GetString("GENERAL", "DATA_FILENAME_DATE_FORMAT", "")))
            If (tmpStr = "") Or (Not {"yyMMdd", "yyyyMMdd"}.Contains(tmpStr)) Then
                'myLogger.AppendToFile(GlobalVarable.ErrLogFile, myLogger.FormatString(True, GlobalVarable.AppName, "SetGlobalVariable", "DATA_FILENAME_DATE_FORMAT not correct, please check!"))
                myResult = WriteToLog("NOSECNAME", "DATA_FILENAME_DATE_FORMAT not correct, please check!")
                Exit Sub
            Else
                GlobalVarable.DATA_FILENAME_DATE_FORMAT = tmpStr
            End If

            tmpStr = ""
            tmpStr = IIf(myINIFile.GetString("GENERAL", "DATA_RETENTION_PERIOD", "") = "", "24", myINIFile.GetString("GENERAL", "DATA_RETENTION_PERIOD", ""))
            If IsNumeric(tmpStr) Then
                GlobalVarable.DATA_RETENTION_PERIOD = Int(Val(tmpStr))
            Else
                GlobalVarable.DATA_RETENTION_PERIOD = 24
            End If

            GlobalVarable.MailEnable = IIf(Trim(myINIFile.GetString("MAIL", "MAILENABLE", "")) = "", "N", Trim(myINIFile.GetString("MAIL", "MAILENABLE", "")))
            GlobalVarable.MailServer = IIf(Trim(myINIFile.GetString("MAIL", "MAILSERVER", "")) = "", "SMTPEx.intranet.hkbea.com", Trim(myINIFile.GetString("MAIL", "MAILSERVER", "")))
            GlobalVarable.MailServerPort = IIf(Trim(myINIFile.GetString("MAIL", "MAILPORT", "")) = "", "25", Trim(myINIFile.GetString("MAIL", "MAILPORT", "")))
            GlobalVarable.MailFrom = IIf(Trim(myINIFile.GetString("MAIL", "MAILFROM", "")) = "", "TFS_Scheduler@hkbea.com", Trim(myINIFile.GetString("MAIL", "MAILFROM", "")))
            GlobalVarable.SUCCESSMailToList = IIf(Trim(myINIFile.GetString("MAIL", "SUCCESSMAILTO", "")) = "", "hkg-tfs-sic@hkbea.com", Trim(myINIFile.GetString("MAIL", "SUCCESSMAILTO", "")))
            GlobalVarable.SUCCESSMailCCList = IIf(Trim(myINIFile.GetString("MAIL", "SUCCESSMAILCC", "")) = "", "", Trim(myINIFile.GetString("MAIL", "SUCCESSMAILCC", "")))
            GlobalVarable.MailSubject = IIf(Trim(myINIFile.GetString("MAIL", "SUBJECT", "")) = "", "", Trim(myINIFile.GetString("MAIL", "SUBJECT", "")))
            GlobalVarable.ERRMailToList = IIf(Trim(myINIFile.GetString("MAIL", "ERRMAILTO", "")) = "", "hkg-tfs-sic@hkbea.com", Trim(myINIFile.GetString("MAIL", "ERRMAILTO", "")))
            GlobalVarable.ERRMailCCList = IIf(Trim(myINIFile.GetString("MAIL", "ERRMAILCC", "")) = "", "", Trim(myINIFile.GetString("MAIL", "ERRMAILCC", "")))

            GlobalVarable.MailErrSubject = IIf(Trim(myINIFile.GetString("MAIL", "ERRSUBJECT", "")) = "", "", Trim(myINIFile.GetString("MAIL", "ERRSUBJECT", "")))
            GlobalVarable.MailBody = IIf(Trim(myINIFile.GetString("MAIL", "BODY", "")) = "", "", Trim(myINIFile.GetString("MAIL", "BODY", "")))

            EmailContent = EmailContent & GlobalVarable.MailBody & vbNewLine

            For i = 1 To GlobalVarable.FileNum

                SecName = "FILE_" & Right("000" & i.ToString, 3)
                TableName = myINIFile.GetString(SecName, "TABLENAME", "")
                FileName = myINIFile.GetString(SecName, "FILENAME", "")

                FileHeader = IIf(UCase(Trim(myINIFile.GetString(SecName, "HEADER", ""))) = "", "N", UCase(Trim(myINIFile.GetString(SecName, "HEADER", ""))))
                FileTrailer = IIf(UCase(Trim(myINIFile.GetString(SecName, "TRAILER", ""))) = "", "N", UCase(Trim(myINIFile.GetString(SecName, "TRAILER", ""))))
                FieldNum = myINIFile.GetString(SecName, "COL_NO", "")

                myFieldList = New ImportFields With
                {
                    .FileName = FileName,
                    .TableName = TableName,
                    .FileHeader = FileHeader,
                    .FileTrailer = FileTrailer,
                    .FieldNum = FieldNum
                }
                For j = 1 To FieldNum
                    With myFieldList
                        .StartPos.Add(Val(IIf(myINIFile.GetString(SecName, "COL_" & Right("000" & j.ToString, 3) & "_STARTPOS", "") = "", "0", myINIFile.GetString(SecName, "COL_" & Right("000" & j.ToString, 3) & "_STARTPOS", ""))))
                        .FieldLength.Add(Val(IIf(myINIFile.GetString(SecName, "COL_" & Right("000" & j.ToString, 3) & "_LEN", "") = "", "0", myINIFile.GetString(SecName, "COL_" & Right("000" & j.ToString, 3) & "_LEN", ""))))
                        .FieldName.Add(myINIFile.GetString(SecName, "COL_" & Right("000" & j.ToString, 3) & "_FIELDNAME", ""))
                        .DataType.Add(myINIFile.GetString(SecName, "COL_" & Right("000" & j.ToString, 3) & "_DATATYPE", ""))
                        .FieldSource.Add(myINIFile.GetString(SecName, "COL_" & Right("000" & j.ToString, 3) & "_SOURCE", ""))
                    End With
                Next
                myTableList.ItemList.Add(myFieldList)
            Next
            myResult = WriteToLog("INITIALIZATION", "ERR00200102")


            '****************************
            '* Setup Data Source Config *
            '****************************
            Dim strEncryptedPW As String = ""
            myDSINI = New IniFile(GlobalVarable.INI_DataSource)
            myEmailINI = New IniFile(GlobalVarable.INI_Email)

            'TI
            config.TIDSN = myDSINI.GetString("DATASOURCE", "TIDSN", "")
            config.TIUserID = myDSINI.GetString("DATASOURCE", "TIUserID", "")
            strEncryptedPW = ""
            strEncryptedPW = myDSINI.GetString("DATASOURCE", "TIUesrPwd", "")
            config.TIPwd = DataSecurity.Decrypt(strEncryptedPW)
            'RPT
            config.RptDSN = myDSINI.GetString("DATASOURCE", "RptDSN", "")
            config.RptUserID = myDSINI.GetString("DATASOURCE", "RptUserID", "")
            strEncryptedPW = ""
            strEncryptedPW = myDSINI.GetString("DATASOURCE", "RptUesrPwd", "")
            config.RptPwd = DataSecurity.Decrypt(strEncryptedPW)
            'TI Global
            config.TIGlobalDSN = myDSINI.GetString("DATASOURCE", "TIGLOBALDSN", "")
            config.TIGlobalUserID = myDSINI.GetString("DATASOURCE", "TIGLOBALUserID", "")

            'Given that many dataSource has not TIGlobalUserPwd, it will fail to decrypt by default
            'so adding the following check will prevent the program dieing from lacking unused pw
            Dim globalEncryptedPW As String = myDSINI.GetString("DATASOURCE", "TIGLOBALUserPwd", "")
            If Not String.IsNullOrEmpty(config.TIGlobalDSN?.Trim()) AndAlso
               Not String.IsNullOrEmpty(config.TIGlobalUserID?.Trim()) AndAlso
               Not String.IsNullOrEmpty(globalEncryptedPW?.Trim()) Then
                config.TIGlobalPwd = DataSecurity.Decrypt(globalEncryptedPW)
            Else
                config.TIGlobalPwd = globalEncryptedPW
            End If

            '**************************
            'Build up a DB Connection *
            '**************************
            myResult = WriteToLog("DBCONNECTION", "ERR00500101")
            Dim rptConnStr As String = ""
            Dim strDataSource As String = ""
            'strDataSource = "(DESCRIPTION=(ADDRESS=(PROTOCOL=tcp)(HOST=10.129.84.144)(PORT=11521))(CONNECT_DATA=(SERVICE_NAME=TIIUUATR)))"
            strDataSource = config.RptDSN
            Console.WriteLine(strDataSource & "-" & config.RptUserID & "-" & config.RptPwd)
            rptConnStr = myDB.GetConnectionStr(strDataSource, config.RptUserID, config.RptPwd)

            Dim myConn As New OleDbConnection
            Dim myParam As New OleDb.OleDbParameter
            Dim myAdapter As New OleDbDataAdapter
            Dim myDataset As New DataSet
            Dim strSQL As String = ""
            Dim tmpRunDate As String = ""
            Dim mySR As StreamReader
            Dim DataItem As List(Of String)
            Dim DataItemList As List(Of List(Of String))
            Dim tmpDataStr As String = ""
            Dim myCmd As OleDbCommand
            Dim strSQL1, strSQL2, strSQL3 As String

            '*************************
            '* Upload Data File      *
            '*************************
            Try
                myResult = WriteToLog("DATAUPLOAD", "ERR00300101")
                myConn.ConnectionString = rptConnStr
                myConn.Open()
                strSQL = "select Run_Date, Prev_Business_Date from V_Get_Run_Date"
                myAdapter = New OleDbDataAdapter(strSQL, myConn)
                myAdapter.Fill(myDataset)
                If myDataset.Tables(0).Rows.Count > 0 Then
                    GlobalVarable.RunDate = myDataset.Tables(0).Rows(0).Item(0)
                    GlobalVarable.Prev_Bus_Date = myDataset.Tables(0).Rows(0).Item(1)
                End If
                If Not IsDate(GlobalVarable.RunDate) Then
                    myResult = WriteToLog("NOSECNAME", "Get Run Date Error : " & GlobalVarable.RunDate & " , please check!")
                    Exit Sub
                End If

                If Not IsDate(GlobalVarable.Prev_Bus_Date) Then
                    myResult = WriteToLog("NOSECNAME", "Get Prev_Business_Date Error : " & GlobalVarable.RunDate & " , please check!")
                    Exit Sub
                End If

                'This while loop used to handle if the number of date between the previous business and the run date larger than 1
                GlobalVarable.Current_Run_Date = DateAdd("d", 1, GlobalVarable.Prev_Bus_Date).ToString("yyyy-MM-dd")
                While DateDiff("d", GlobalVarable.Current_Run_Date, GlobalVarable.RunDate) >= 0
                    For i = 0 To (myTableList.ItemList.Count - 1)
                        '*****************************************
                        '* Errase Data out of Retention Perios   *
                        '*****************************************

                        'strSQL3 = "Delete from " & myTableList.ItemList(i).TableName & " Where LastCalendarDate <  To_Char(Add_Months(To_Date('" & GlobalVarable.RunDate & "', 'yyyy-mm-dd'),-" & GlobalVarable.DATA_RETENTION_PERIOD & "),'yyyy-mm-dd')"
                        'myCmd = New OleDbCommand
                        'myCmd.Connection = myConn
                        'myCmd.CommandText = strSQL3
                        'myCmd.ExecuteNonQuery()
                        'myCmd.Dispose()
                        'myResult = WriteToLog("NOSECNAME", myTableList.ItemList(i).TableName & ": Delete History Data older then " & GlobalVarable.DATA_RETENTION_PERIOD & " Months!")

                        myCmd = myConn.CreateCommand
                        myCmd = New OleDbCommand("sp_ARR_Rate_HouseKeep", myConn)
                        myCmd.CommandType = CommandType.StoredProcedure
                        myParam = New OleDb.OleDbParameter
                        myParam = myCmd.Parameters.Add("@I_RETENT_PERIOD", OleDbType.Integer)
                        myParam.Direction = ParameterDirection.Input
                        myParam.Value = GlobalVarable.DATA_RETENTION_PERIOD

                        Try
                            myCmd.ExecuteNonQuery()
                        Catch ex As Exception
                            myResult = WriteToLog("NOSECNAME", ex.Message)
                            GlobalVarable.ProgramStatus = False
                        End Try

                        '***************************
                        '* Erase RunDate's data   *
                        '***************************
                        strSQL3 = "Delete " & myTableList.ItemList(i).TableName & " Where LastCalendarDate = '" & GlobalVarable.Current_Run_Date & "'"
                        myCmd = New OleDbCommand
                        myCmd.Connection = myConn
                        myCmd.CommandText = strSQL3
                        myCmd.ExecuteNonQuery()
                        myCmd.Dispose()
                        myResult = WriteToLog("NOSECNAME", myTableList.ItemList(i).TableName & ": Delete run_date uploaded data!")
                        strSQL3 = ""
                    Next

                    For i = 0 To (myTableList.ItemList.Count - 1)
                        tmpFileName = myTableList.ItemList(i).FileName
                        If tmpFileName.IndexOf("<") <> -1 Then
                            FileName = Left(tmpFileName, tmpFileName.IndexOf("<")) & Date.Parse(GlobalVarable.Current_Run_Date).ToString("yyMMdd") & Right(tmpFileName, Len(tmpFileName) - (tmpFileName.IndexOf(">") + 1))
                        Else
                            FileName = tmpFileName
                        End If

                        If Not System.IO.File.Exists(GlobalVarable.DataFilePath & FileName) Then
                            myResult = WriteToLog("NOSECNAME", FileName & " not exists, please check!")
                            ErrEmailContent = ErrEmailContent & FileName & " not exists, please check!" & vbNewLine
                            'ErrMsg = SendMail(GlobalVarable.MailErrSubject, FileName & " not exists, please check!", False)
                            'Exit Sub
                        Else
                            mySR = New StreamReader(GlobalVarable.DataFilePath & FileName)
                            DataItemList = New List(Of List(Of String))
                            Dim tmpRowPointer As Long = 0

                            While Not mySR.EndOfStream
                                tmpDataStr = mySR.ReadLine
                                'If Left(tmpDataStr, 2) = "50" Then
                                DataItem = New List(Of String)
                                For j = 0 To (myTableList.ItemList(i).FieldNum - 1)
                                    If UCase(myTableList.ItemList(i).FieldSource(j)) <> "SYSTEM" Then
                                        DataItem.Add(Trim(Mid(tmpDataStr, myTableList.ItemList(i).StartPos(j), myTableList.ItemList(i).FieldLength(j))))
                                    Else
                                        DataItem.Add(GlobalVarable.Current_Run_Date)
                                    End If
                                Next
                                DataItemList.Add(DataItem)
                                'End If
                                tmpRowPointer = tmpRowPointer + 1
                            End While
                            mySR.Close()
                            If myTableList.ItemList(i).FileHeader = "Y" Then
                                DataItemList.Item(0).RemoveRange(0, myTableList.ItemList(i).FieldNum)
                            End If
                            If myTableList.ItemList(i).FileTrailer = "Y" Then
                                DataItemList.Item(tmpRowPointer - 1).RemoveRange(0, myTableList.ItemList(i).FieldNum)
                            End If

                            If DataItemList.Count > 0 Then

                                strSQL3 = ""
                                For k = 0 To (DataItemList.Count - 1)
                                    If DataItemList.Item(k).Count > 0 Then
                                        strSQL1 = "Insert into " & myTableList.ItemList(i).TableName & " ("
                                        strSQL2 = "Values ("
                                        For j = 0 To (myTableList.ItemList(i).FieldNum - 1)
                                            If j = (myTableList.ItemList(i).FieldNum - 1) Then
                                                strSQL1 = strSQL1 & myTableList.ItemList(i).FieldName(j)
                                                Select Case UCase(myTableList.ItemList(i).DataType(j))
                                                    Case "STRING"
                                                        strSQL2 = strSQL2 & "'" & DataItemList.Item(k).Item(j) & "'"
                                                    Case "DATE"
                                                        strSQL2 = strSQL2 & "to_date('" & DataItemList.Item(k).Item(j) & "','yyyy-mm-dd')"
                                                    Case "NUMERIC"
                                                        strSQL2 = strSQL2 & "" & DataItemList.Item(k).Item(j) & ""
                                                End Select
                                            Else
                                                strSQL1 = strSQL1 & myTableList.ItemList(i).FieldName(j) & ","
                                                Select Case UCase(myTableList.ItemList(i).DataType(j))
                                                    Case "STRING"
                                                        strSQL2 = strSQL2 & "'" & DataItemList.Item(k).Item(j) & "',"
                                                    Case "DATE"
                                                        strSQL2 = strSQL2 & "to_date('" & DataItemList.Item(k).Item(j) & "','yyyy-mm-dd'),"
                                                    Case "NUMERIC"
                                                        strSQL2 = strSQL2 & "" & DataItemList.Item(k).Item(j) & ","
                                                End Select
                                            End If
                                        Next
                                        strSQL1 = strSQL1 & ") "
                                        strSQL2 = strSQL2 & ") "
                                        strSQL3 = strSQL1 & strSQL2

                                        Try
                                            myCmd = New OleDbCommand
                                            myCmd.Connection = myConn
                                            myCmd.CommandText = strSQL3
                                            myCmd.ExecuteNonQuery()
                                            myCmd.Dispose()
                                        Catch ex As Exception
                                            myResult = WriteToLog("DATAUPLOAD", ex.Message & ": " & strSQL3)
                                        End Try

                                        strSQL1 = ""
                                        strSQL2 = ""
                                        strSQL3 = ""
                                    End If
                                Next
                            End If
                            EmailContent = EmailContent & "Upload " & GlobalVarable.DataFilePath & FileName & " Success!" & vbNewLine
                            myResult = WriteToLog("NOSECNAME", "Upload " & GlobalVarable.DataFilePath & FileName & " Success!")
                        End If
                    Next
                    GlobalVarable.Current_Run_Date = DateAdd("d", 1, GlobalVarable.Current_Run_Date).ToString("yyyy-MM-dd")
                End While


                myResult = WriteToLog("DATAUPLOAD", "ERR00300102")
            Catch ex As Exception
                myResult = WriteToLog("DATAUPLOAD", "ERR00300103")
                Console.WriteLine(ex.Message)
                GlobalVarable.ProgramStatus = False
            End Try

            '***************************
            '* Close the DB Connection *
            '***************************
            Try
                myAdapter.Dispose()
                myConn.Close()
                myConn = Nothing
                myResult = WriteToLog("DBCONNECTION", "ERR00500103")
            Catch ex As Exception
                myResult = WriteToLog("DBCONNECTION", "ERR00500102")
                Console.WriteLine(ex.Message)
                GlobalVarable.ProgramStatus = False
            End Try

        End If

        '************************************
        '* Send out Status Result via Email *
        '************************************

        If GlobalVarable.MailEnable = "Y" Then
            myResult = WriteToLog("EMAIL", "ERR00600101")
            If GlobalVarable.ProgramStatus Then
                ErrMsg = SendMail(GlobalVarable.MailSubject, EmailContent, True)
                If ErrEmailContent <> "" Then
                    ErrMsg = SendMail(GlobalVarable.MailErrSubject, ErrEmailContent, False)
                End If
            Else
                ErrMsg = SendMail(GlobalVarable.MailErrSubject, ErrEmailContent, False)
            End If
            myResult = WriteToLog("EMAIL", "ERR00600103")

        End If

        myResult = WriteToLog("GENERAL", "ERR00100102")

    End Sub

    Public Function WriteToLog(SecName As String, KeyName As String) As Boolean

        Dim ErrMsg As String = ""
        Dim myErrINI As New IniFile(GlobalVarable.INI_ErrCode)
        Dim myLogger As New Logger

        Try
            If Trim(SecName) <> "NOSECNAME" Then
                myErrINI = New IniFile(GlobalVarable.INI_ErrCode)
                ErrMsg = myErrINI.GetString(SecName, KeyName, "")
            Else
                ErrMsg = KeyName
            End If
            myLogger.AppendToFile(GlobalVarable.ErrLogFile, myLogger.FormatString(True, GlobalVarable.AppName, "", ErrMsg))
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            myLogger.AppendToFile(GlobalVarable.ErrLogFile, myLogger.FormatString(True, GlobalVarable.AppName, "", ErrMsg))
            Return False
        Finally
            Console.WriteLine(ErrMsg)
        End Try

        myErrINI = Nothing
        myLogger = Nothing

    End Function
    Public Function SetGlobalVariable() As Boolean

        Try

            'GlobalVarable.AppPath = "D:\Development\DataImport\"
            GlobalVarable.AppPath = My.Application.Info.DirectoryPath & "\"
            'GlobalVarable.INI_DataFormat = ConfigurationManager.AppSettings("INI_DataFormat")
            GlobalVarable.ErrLogFile = GlobalVarable.AppPath & "DataImport" & ".Log"
            GlobalVarable.INI_DataSource = ConfigurationManager.AppSettings("INI_DataSource")
            GlobalVarable.INI_Email = ConfigurationManager.AppSettings("INI_Email")
            GlobalVarable.INI_ErrCode = GlobalVarable.AppPath & ConfigurationManager.AppSettings("INI_ErrCode")
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    Public Function SendMail(ByVal strEmailSubject As String, ByVal strEmailContent As String, ByVal SuccessStatusFlag As Boolean) As String
        Dim mail As New MailMessage()
        Dim arrTO As [String]()
        Dim arrCC As [String]()
        Dim strResult As String = "OK"

        Try
            mail.From = New MailAddress(GlobalVarable.MailFrom, "")
            If SuccessStatusFlag = True Then
                If GlobalVarable.SUCCESSMailToList <> "" Then
                    arrTO = GlobalVarable.SUCCESSMailToList.Split(";")
                    For Each toPt As String In arrTO
                        If toPt.Trim = "" Then
                            Continue For
                        End If

                        Try
                            mail.To.Add(toPt)
                        Catch ex As Exception
                            strResult = "Error occur: " & ex.Message
                        End Try
                    Next
                End If

                If GlobalVarable.SUCCESSMailCCList <> "" Then
                    arrCC = GlobalVarable.SUCCESSMailCCList.Split(";")
                    For Each cc In arrCC
                        If cc.Trim = "" Then
                            Continue For
                        End If

                        Try
                            mail.CC.Add(cc)
                        Catch ex As Exception
                            strResult = "Error occur: " & ex.Message
                        End Try
                    Next
                End If
            Else
                If GlobalVarable.ERRMailToList <> "" Then
                    arrTO = GlobalVarable.ERRMailToList.Split(";")
                    For Each toPt As String In arrTO
                        If toPt.Trim = "" Then
                            Continue For
                        End If

                        Try
                            mail.To.Add(toPt)
                        Catch ex As Exception
                            strResult = "Error occur: " & ex.Message
                        End Try
                    Next
                End If

                If GlobalVarable.ERRMailCCList <> "" Then
                    arrCC = GlobalVarable.ERRMailCCList.Split(";")
                    For Each cc In arrCC
                        If cc.Trim = "" Then
                            Continue For
                        End If

                        Try
                            mail.CC.Add(cc)
                        Catch ex As Exception
                            strResult = "Error occur: " & ex.Message
                        End Try
                    Next
                End If
            End If

            mail.SubjectEncoding = System.Text.Encoding.UTF8
            mail.BodyEncoding = System.Text.Encoding.UTF8
            mail.Subject = strEmailSubject
            mail.Body = strEmailContent
            'Dim a As New System.Net.Mail.Attachment(g_ErrorLogFile)
            'mail.Attachments.Add(a)

            Dim smtp As New SmtpClient(GlobalVarable.MailServer, GlobalVarable.MailServerPort)

            smtp.UseDefaultCredentials = False
            smtp.Send(mail)

        Catch ex As Exception
            strResult = "Error occur: " & ex.Message
        End Try

        mail = Nothing
        Return strResult

    End Function
End Module
