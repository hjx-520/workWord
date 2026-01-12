Imports System
Imports System.IO
Imports System.Configuration
Imports System.Text
Imports System.Data
Imports System.Data.OleDb
Imports System.Net.Mail
Imports System.Globalization
Imports IBM.WMQ

Module Main
    Private outMsg As String = ""
    Sub Main()

        Dim myResult As Boolean = True
        Dim myLogger As New Logger

        Dim aArgs() As String

        Dim myINIFile, myErrINI, myDSINI, myEmailINI As IniFile
        Dim myDB As New DBProcOLEDB
        Dim ErrMsg As String = ""
        Dim ErrEmailContent As String = ""
        Dim tmpFileName As String
        Dim tmpStr As String = ""
        Dim i, j As Long

        GlobalVarable.ProgramStatus = True

        'GlobalVarable.AppPath = "D:\Development2\ARR_MQ\"
        GlobalVarable.AppPath = My.Application.Info.DirectoryPath & "\"

        GlobalVarable.ErrLogFile = GlobalVarable.AppPath & GlobalVarable.AppName & ".Log"
        GlobalVarable.INI_DataSource = ConfigurationManager.AppSettings("INI_DataSource")
        GlobalVarable.INI_Email = ConfigurationManager.AppSettings("INI_Email")
        GlobalVarable.INI_ErrCode = GlobalVarable.AppPath & ConfigurationManager.AppSettings("INI_ErrCode")

        'Setup and assign values to variables
        'argument checking
        'aArgs = System.Environment.GetCommandLineArgs
        aArgs = {"1", "GenRPAFile.ini"}
        Console.WriteLine(aArgs)
        If aArgs.Count < 2 Then
            myResult = WriteToLog("NOSECNAME", "Missing or Invalid parameter!")
            myResult = WriteToLog("NOSECNAME", "Example: GenRPAFile GenRPAFIle.ini")
            Exit Sub
        End If

        'GlobalVarable.INI_App = Path.GetFullPath(aArgs(1))
        GlobalVarable.INI_App = GlobalVarable.AppPath & aArgs(1)

        Console.WriteLine(GlobalVarable.INI_App)
        Console.WriteLine(GlobalVarable.AppPath)

        myINIFile = New IniFile(GlobalVarable.INI_App)
        myErrINI = New IniFile(GlobalVarable.INI_ErrCode)

        'If Not myResult Then
        ' myLogger.AppendToFile(GlobalVarable.ErrLogFile, myLogger.FormatString(True, GlobalVarable.AppName, "SetGlobalVariable", "SetGlobalVariable error, please check!"))
        'Exit Sub
        'End If

        myResult = WriteToLog("GENERAL", "ERR00100101")

        If Not File.Exists(GlobalVarable.INI_App) Then
            myResult = WriteToLog("INITIALIZATION", "ERR00200104")
        Else
            myResult = WriteToLog("INITIALIZATION", "ERR00200105")
            GlobalVarable.Output_File_Path = Trim(myINIFile.GetString("GENERAL", "OUTPUT_PATH", ""))
            GlobalVarable.Output_File_Path = IIf(Right(GlobalVarable.Output_File_Path, 1) = "\", GlobalVarable.Output_File_Path, GlobalVarable.Output_File_Path & "\")

            If Not Directory.Exists(GlobalVarable.Output_File_Path) Then
                myLogger.AppendToFile(GlobalVarable.ErrLogFile, myLogger.FormatString(True, GlobalVarable.AppName, "SetGlobalVariable", "Director of OUTPUT_PATH not exist, please check!"))
                Exit Sub
            End If

            GlobalVarable.TXT_File_Name = Trim(myINIFile.GetString("FILE", "TXT_FILE_NAME", ""))

            tmpStr = ""
            tmpStr = IIf(Trim(myINIFile.GetString("FILE", "CSV_DELIMITER", "")) = "", ",", Trim(myINIFile.GetString("FILE", "CSV_DELIMITER", "")))
            If tmpStr = "" Then
                myLogger.AppendToFile(GlobalVarable.ErrLogFile, myLogger.FormatString(True, GlobalVarable.AppName, "SetGlobalVariable", "CSV_Delimiter cannot be blanked , please check!"))
                Exit Sub
            Else
                GlobalVarable.CSV_DELIMTER = tmpStr
            End If

            tmpStr = ""
            tmpStr = IIf(Trim(myINIFile.GetString("FILE", "FILE_FORMAT_TXT", "")) = "", "N", UCase(Trim(myINIFile.GetString("FILE", "FILE_FORMAT_TXT", ""))))
            If (Len(tmpStr) > 2) Or (Not {"Y", "N"}.Contains(tmpStr)) Then
                myLogger.AppendToFile(GlobalVarable.ErrLogFile, myLogger.FormatString(True, GlobalVarable.AppName, "SetGlobalVariable", "The value of FILE_FORMAT is Y/N, please check!"))
                Exit Sub
            Else
                GlobalVarable.FILE_FORMAT_TXT = tmpStr
            End If

            tmpStr = ""
            tmpStr = IIf(Trim(myINIFile.GetString("FILE", "FILE_FORMAT_CSV", "")) = "", "N", UCase(Trim(myINIFile.GetString("FILE", "FILE_FORMAT_CSV", ""))))
            If (Len(tmpStr) > 2) Or (Not {"Y", "N"}.Contains(tmpStr)) Then
                myLogger.AppendToFile(GlobalVarable.ErrLogFile, myLogger.FormatString(True, GlobalVarable.AppName, "SetGlobalVariable", "The value of FILE_FORMAT is Y/N, please check!"))
                Exit Sub
            Else
                GlobalVarable.FILE_FORMAT_CSV = tmpStr
            End If

            GlobalVarable.TXT_Col_Width = Trim(myINIFile.GetString("FILE", "TXT_COL_WIDTH", "")).Split(",")
            If GlobalVarable.TXT_Col_Width.Count = 0 Then
                myLogger.AppendToFile(GlobalVarable.ErrLogFile, myLogger.FormatString(True, GlobalVarable.AppName, "SetGlobalVariable", "TXT_COL_WIDTH cannot be blanked , please check!"))
                Exit Sub
            End If
            GlobalVarable.TXT_Col_Type = Trim(myINIFile.GetString("FILE", "TXT_COL_TYPE", "")).Split(",")
            If GlobalVarable.TXT_Col_Type.Count <> GlobalVarable.TXT_Col_Width.Count Then
                myLogger.AppendToFile(GlobalVarable.ErrLogFile, myLogger.FormatString(True, GlobalVarable.AppName, "SetGlobalVariable", "Number of TXT_COL_WIDTH <> TXT_COL_TYPE, please check!"))
                Exit Sub
            End If
            For i = 0 To (GlobalVarable.TXT_Col_Type.Count - 1)
                If Not {"STRING", "DATE", "NUMBER"}.Contains(GlobalVarable.TXT_Col_Type(i)) Then
                    myLogger.AppendToFile(GlobalVarable.ErrLogFile, myLogger.FormatString(True, GlobalVarable.AppName, "SetGlobalVariable", "Wrong Data Type in TXT_COL_TYPE, please check!"))
                    Exit Sub
                End If
                If Not IsNumeric(GlobalVarable.TXT_Col_Width(i)) Then
                    myLogger.AppendToFile(GlobalVarable.ErrLogFile, myLogger.FormatString(True, GlobalVarable.AppName, "SetGlobalVariable", "Wrong Data Type in TXT_COL_TYPE, please check!"))
                    Exit Sub
                End If
            Next

            GlobalVarable.TXT_FILLER_WIDTH = IIf(Trim(myINIFile.GetString("FILE", "TXT_FILLER_WIDTH", "")) = "", 0, Val(myINIFile.GetString("FILE", "TXT_FILLER_WIDTH", "")))
            If Not IsNumeric(GlobalVarable.TXT_FILLER_WIDTH) Then
                myLogger.AppendToFile(GlobalVarable.ErrLogFile, myLogger.FormatString(True, GlobalVarable.AppName, "SetGlobalVariable", "TXT_FILLER_WIDTH should be numberic, please check!"))
                Exit Sub
            End If

            GlobalVarable.CSV_Col_Header = IIf(Trim(myINIFile.GetString("FILE", "CSV_COL_HEADER", "")) = "", "N", UCase(Trim(myINIFile.GetString("FILE", "CSV_COL_HEADER", ""))))
            GlobalVarable.CSV_Col_Header_Caption = Trim(myINIFile.GetString("FILE", "CSV_COL_HEADER_CAPTION", "")).Split(",")

            tmpStr = IIf(Trim(myINIFile.GetString("FILE", "HEADER_TRAILER_TXT", "")) = "", "N", UCase(Trim(myINIFile.GetString("FILE", "HEADER_TRAILER_TXT", ""))))
            If Len(tmpStr) > 2 Then
                myLogger.AppendToFile(GlobalVarable.ErrLogFile, myLogger.FormatString(True, GlobalVarable.AppName, "SetGlobalVariable", "The value of Header / Trailer Flag is Y/N, please check!"))
                Exit Sub
            Else
                GlobalVarable.Header_Trailer_TXT_Flag = tmpStr
            End If

            tmpStr = ""
            tmpStr = IIf(Trim(myINIFile.GetString("FILE", "HEADER_TRAILER_CSV", "")) = "", "N", UCase(Trim(myINIFile.GetString("FILE", "HEADER_TRAILER_CSV", ""))))
            If Len(tmpStr) > 2 Then
                myLogger.AppendToFile(GlobalVarable.ErrLogFile, myLogger.FormatString(True, GlobalVarable.AppName, "SetGlobalVariable", "The value of Header / Trailer Flag is Y/N, please check!"))
                Exit Sub
            Else
                GlobalVarable.Header_Trailer_CSV_Flag = tmpStr
            End If

            GlobalVarable.DATA_FILENAME_DATE_FORMAT_S = IIf(Right(Trim(myINIFile.GetString("GENERAL", "DATA_FILE_DATE_FORMAT_S", "")), 1) = "", "yyMMdd", Trim(myINIFile.GetString("GENERAL", "DATA_FILE_DATE_FORMAT_S", "")))
            GlobalVarable.DATA_FILENAME_DATE_FORMAT_L = IIf(Right(Trim(myINIFile.GetString("GENERAL", "DATA_FILE_DATE_FORMAT_L", "")), 1) = "", "yyyyMMdd", Trim(myINIFile.GetString("GENERAL", "DATA_FILE_DATE_FORMAT_L", "")))
            GlobalVarable.DATA_RETENTION_PERIOD = Val(IIf(myINIFile.GetString("GENERAL", "DATA_RETENTION_PERIOD", "") = "", "1000", myINIFile.GetString("GENERAL", "DATA_RETENTION_PERIOD", "")))
            GlobalVarable.MailEnable = IIf(Trim(myINIFile.GetString("MAIL", "MAILENABLE", "")) = "", "N", Trim(myINIFile.GetString("MAIL", "MAILENABLE", "")))
            GlobalVarable.MailServer = IIf(Trim(myINIFile.GetString("MAIL", "MAILSERVER", "")) = "", "SMTPEx.intranet.hkbea.com", Trim(myINIFile.GetString("MAIL", "MAILSERVER", "")))
            GlobalVarable.MailServerPort = IIf(Trim(myINIFile.GetString("MAIL", "MAILPORT", "")) = "", "25", Trim(myINIFile.GetString("MAIL", "MAILPORT", "")))
            GlobalVarable.MailFrom = IIf(Trim(myINIFile.GetString("MAIL", "MAILFROM", "")) = "", "TFS_Scheduler@hkbea.com", Trim(myINIFile.GetString("MAIL", "MAILFROM", "")))
            GlobalVarable.SUCCESSMailToList = IIf(Trim(myINIFile.GetString("MAIL", "SUCCESSMAILTO", "")) = "", "hkg-tfs-sic@hkbea.com", Trim(myINIFile.GetString("MAIL", "SUCCESSMAILTO", "")))
            GlobalVarable.SUCCESSMailCCList = IIf(Trim(myINIFile.GetString("MAIL", "SUCCESSMAILCC", "")) = "", "", Trim(myINIFile.GetString("MAIL", "SUCCESSMAILCC", "")))
            GlobalVarable.MailSubject = IIf(Trim(myINIFile.GetString("MAIL", "SUBJECT", "")) = "", "", Trim(myINIFile.GetString("MAIL", "SUBJECT", "")))
            GlobalVarable.ERRMailToList = IIf(Trim(myINIFile.GetString("MAIL", "ERRMAILTO", "")) = "", "hkg-tfs-sic@hkbea.com", Trim(myINIFile.GetString("MAIL", "ERRMAILTO", "")))
            GlobalVarable.ERRMailCCList = IIf(Trim(myINIFile.GetString("MAIL", "ERRMAILCC", "")) = "", "", Trim(myINIFile.GetString("MAIL", "ERRMAILCC", "")))
            GlobalVarable.MailSubject = IIf(Trim(myINIFile.GetString("MAIL", "SUBJECT", "")) = "", "", Trim(myINIFile.GetString("MAIL", "SUBJECT", "")))
            GlobalVarable.MailErrSubject = IIf(Trim(myINIFile.GetString("MAIL", "ERRSUBJECT", "")) = "", "", Trim(myINIFile.GetString("MAIL", "ERRSUBJECT", "")))
            GlobalVarable.MailBody = IIf(Trim(myINIFile.GetString("MAIL", "BODY", "")) = "", "", Trim(myINIFile.GetString("MAIL", "BODY", "")))
            GlobalVarable.Attachment = IIf(Trim(myINIFile.GetString("MAIL", "ATTACHMENT", "")) = "", "TXT", UCase(Trim(myINIFile.GetString("MAIL", "ATTACHMENT", ""))))

            tmpStr = ""
            tmpStr = Trim(myINIFile.GetString("MQ", "MQ_SERVER", ""))
            If tmpStr = "" Then
                myLogger.AppendToFile(GlobalVarable.ErrLogFile, myLogger.FormatString(True, GlobalVarable.AppName, "SetGlobalVariable", "MQ_SERVER can not be blank, please check!"))
                Exit Sub
            Else
                GlobalVarable.MQ_SERVER = tmpStr
            End If

            tmpStr = ""
            tmpStr = Trim(myINIFile.GetString("MQ", "MQ_CHANNEL", ""))
            If tmpStr = "" Then
                myLogger.AppendToFile(GlobalVarable.ErrLogFile, myLogger.FormatString(True, GlobalVarable.AppName, "SetGlobalVariable", "MQ_CHANNEL can not be blank, please check!"))
                Exit Sub
            Else
                GlobalVarable.MQ_CHANNEL = tmpStr
            End If

            tmpStr = ""
            tmpStr = Trim(myINIFile.GetString("MQ", "MQ_PORT", ""))
            If IsNumeric(tmpStr) Then
                GlobalVarable.MQ_PORT = Int(Val(tmpStr))
            Else
                GlobalVarable.MQ_PORT = 1414
            End If

            tmpStr = ""
            tmpStr = Trim(myINIFile.GetString("MQ", "MQ_Manager", ""))
            If tmpStr = "" Then
                myLogger.AppendToFile(GlobalVarable.ErrLogFile, myLogger.FormatString(True, GlobalVarable.AppName, "SetGlobalVariable", "MQ_Manager can not be blank, please check!"))
                Exit Sub
            Else
                GlobalVarable.MQ_Manager = tmpStr
            End If

            tmpStr = ""
            tmpStr = Trim(myINIFile.GetString("MQ", "MQ_QUEUE_REQ", ""))
            If tmpStr = "" Then
                myLogger.AppendToFile(GlobalVarable.ErrLogFile, myLogger.FormatString(True, GlobalVarable.AppName, "SetGlobalVariable", "MQ_QUEUE_REQ can not be blank, please check!"))
                Exit Sub
            Else
                GlobalVarable.MQ_QUEUE_REQ = tmpStr
            End If

            tmpStr = ""
            tmpStr = Trim(myINIFile.GetString("MQ", "MQ_QUEUE_RPLY", ""))
            If tmpStr = "" Then
                myLogger.AppendToFile(GlobalVarable.ErrLogFile, myLogger.FormatString(True, GlobalVarable.AppName, "SetGlobalVariable", "MQ_QUEUE_RPLY can not be blank, please check!"))
                Exit Sub
            Else
                GlobalVarable.MQ_QUEUE_RPLY = tmpStr
            End If

            tmpStr = ""
            tmpStr = Trim(myINIFile.GetString("MQ", "MQ_TRANS_TYPE", ""))
            If tmpStr = "" Then
                myLogger.AppendToFile(GlobalVarable.ErrLogFile, myLogger.FormatString(True, GlobalVarable.AppName, "SetGlobalVariable", "MQ_TRANS_TYPE can not be blank, please check!"))
                Exit Sub
            Else
                GlobalVarable.MQ_TRANS_TYPE = tmpStr
            End If

            tmpStr = ""
            tmpStr = Trim(myINIFile.GetString("MQ", "MQ_MSG_Expiry", ""))
            If IsNumeric(tmpStr) Then
                GlobalVarable.MQ_MSG_Expiry = Int(Val(tmpStr))
            Else
                GlobalVarable.MQ_MSG_Expiry = 1000
            End If

            tmpStr = ""
            tmpStr = Trim(myINIFile.GetString("MQ", "MQ_Wait_Interval", ""))
            If IsNumeric(tmpStr) Then
                GlobalVarable.MQ_Wait_Interval = Int(Val(tmpStr))
            Else
                GlobalVarable.MQ_Wait_Interval = 10000
            End If


        End If


        Dim config As DataSourceConfig = New DataSourceConfig()

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
        'strDataSource = "(DESCRIPTION=(ADDRESS=(PROTOCOL=tcp)(HOST=10.129.84.144)(PORT=11521))(CONNECT_DATA=(SERVICE_NAME=TIIUTATR)))"
        strDataSource = config.RptDSN
        Console.WriteLine("Connect To: " & strDataSource & "-" & config.RptUserID)
        rptConnStr = myDB.GetConnectionStr(strDataSource, config.RptUserID, config.RptPwd)

        Dim myConn As New OleDbConnection
        Dim myAdapter As New OleDbDataAdapter
        Dim myComm As New OleDbCommand
        Dim myParam As New OleDb.OleDbParameter
        Dim myDataset As New DataSet
        Dim strSQL As String
        Dim tmpRunDate As String = ""
        Dim tmpDataStr As String = ""
        Dim strOSList_csv, strOSList_txt, tmpCSVString, tmpTXTString, tmpColHeader As String


        myResult = WriteToLog("GENERATEREPORT", "ERR00400101")

        '***************************************************
        '* Retrieve Outstanding List with ARR Rate from DB *
        '***************************************************
        Try
            myConn.ConnectionString = rptConnStr
            myConn.Open()

            Console.WriteLine("Get Run Date")
            strSQL = "select Run_Date from v_Get_Run_Date"
            myAdapter = New OleDbDataAdapter(strSQL, myConn)
            myAdapter.Fill(myDataset)
            If myDataset.Tables(0).Rows.Count > 0 Then
                GlobalVarable.RunDate = myDataset.Tables(0).Rows(0).Item(0)
            End If
            Console.WriteLine("Get Run Date Done")
            'strSQL = ""
            'strSQL = strSQL & "	insert into RPT_ARR_RPA ( "
            'strSQL = strSQL & "     BILLS_REF, "
            'strSQL = strSQL & "     BASE_RATE_CODE, "
            'strSQL = strSQL & "     BASE_CCY, "
            'strSQL = strSQL & "     START_DATE, "
            'strSQL = strSQL & "     END_DATE, "
            'strSQL = strSQL & "     BASE_RATE, "
            'strSQL = strSQL & "     ARR_RATE, "
            'strSQL = strSQL & "     SPREAD_RATE, "
            'strSQL = strSQL & "     DAYS_INT, "
            'strSQL = strSQL & "     TENOR_CODE, "
            'strSQL = strSQL & "     ODD_FLAG, "
            'strSQL = strSQL & "     CAS, "
            'strSQL = strSQL & "     ALL_IN_RATE, "
            'strSQL = strSQL & "     RPA_UPDATE_DATE, "
            'strSQL = strSQL & "     RPA_UPDATE_STATUS, "
            'strSQL = strSQL & "     NEXT_BUS_DATE, "
            'strSQL = strSQL & "     LASTCALENDARDATE ) "
            'strSQL = strSQL & " Select "
            'strSQL = strSQL & "     BILLS_REF, "
            'strSQL = strSQL & "     BASE_RATE_CODE, "
            'strSQL = strSQL & "     BASE_CCY, "
            'strSQL = strSQL & "     to_char(START_DATE, 'yyyymmdd'), "
            'strSQL = strSQL & "     to_char(END_DATE, 'yyyymmdd'), "
            'strSQL = strSQL & "     to_char(BASE_RATE,'99999999999.99999'), "
            'strSQL = strSQL & "     to_char(ARR_RATE,'99999999999.99999'), "
            'strSQL = strSQL & "     to_char(SPREAD_RATE,'99999999999.99999'), "
            'strSQL = strSQL & "     to_char(DAYS_INT,'9999'), "
            'strSQL = strSQL & "     TENOR_CODE, "
            'strSQL = strSQL & "     ODD, "
            'strSQL = strSQL & "     to_char(CAS,'99999999999.99999'), "
            'strSQL = strSQL & "     to_char(ALL_IN_RATE,'99999999999.99999'), "
            'strSQL = strSQL & "     to_char(RPA_UPDATE_DATE, 'yyyymmdd'), "
            'strSQL = strSQL & "     RPA_UPDATE_FLAG, "
            'strSQL = strSQL & "     to_char(NEXT_BUS_DATE, 'yyyymmdd'), "
            'strSQL = strSQL & "     to_char((select run_date from rpt_run_Date), 'yyyy-mm-dd')  "
            'strSQL = strSQL & " From V_ARR_OSLOANLIST_4 "

            'strSQL = "Select * from V_ARR_OSLOANLIST_4"
            'strSQL = " Exec sp_Gen_ARR_RPA (24)"
            'myAdapter = New OleDbDataAdapter(strSQL, myConn)
            'myAdapter.Fill(myDataset)

            Console.WriteLine("start Execute sp_Gen_ARR_RPA")
            myComm = myConn.CreateCommand
            myComm = New OleDbCommand("sp_Gen_ARR_RPA", myConn)
            myComm.CommandType = CommandType.StoredProcedure
            myParam = New OleDb.OleDbParameter
            myParam = myComm.Parameters.Add("@I_RETENT_PERIOD", OleDbType.Integer)
            myParam.Direction = ParameterDirection.Input
            myParam.Value = GlobalVarable.DATA_RETENTION_PERIOD
            'myComm.Parameters.Add(New OleDbParameter("@i_RETENT_PERIOD", OleDbType.Integer, 10, ParameterDirection.Input)).Value = GlobalVarable.DATA_RETENTION_PERIOD

            Try
                myComm.ExecuteNonQuery()
            Catch ex As Exception
                myResult = WriteToLog("NOSECNAME", ex.Message)
                GlobalVarable.ProgramStatus = False
            End Try
            'myComm.ExecuteNonQuery()
            'iCount = myComm.ExecuteNonQuery()

            Console.WriteLine("Execute sp_Gen_ARR_RPA done.")

            tmpFileName = GlobalVarable.TXT_File_Name
            If tmpFileName.IndexOf("<") <> -1 Then
                GlobalVarable.TXT_File_Name = Left(tmpFileName, tmpFileName.IndexOf("<")) & Date.Parse(GlobalVarable.RunDate).ToString("yyMMdd") & Right(tmpFileName, Len(tmpFileName) - (tmpFileName.IndexOf(">") + 1))
            End If

            GlobalVarable.CSV_File_Name = Trim(myINIFile.GetString("FILE", "CSV_FILE_NAME", ""))
            tmpFileName = GlobalVarable.CSV_File_Name
            If tmpFileName.IndexOf("<") <> -1 Then
                GlobalVarable.CSV_File_Name = Left(tmpFileName, tmpFileName.IndexOf("<")) & Date.Parse(GlobalVarable.RunDate).ToString("yyMMdd") & Right(tmpFileName, Len(tmpFileName) - (tmpFileName.IndexOf(">") + 1))
            End If

            If File.Exists(GlobalVarable.Output_File_Path & GlobalVarable.CSV_File_Name) Then
                File.Delete(GlobalVarable.Output_File_Path & GlobalVarable.CSV_File_Name)
            End If

            If File.Exists(GlobalVarable.Output_File_Path & GlobalVarable.TXT_File_Name) Then
                File.Delete(GlobalVarable.Output_File_Path & GlobalVarable.TXT_File_Name)
            End If

            myAdapter = Nothing
            myDataset = Nothing
            myDataset = New DataSet

            Console.WriteLine("start update CAS & RATE")

            'strSQL = "select BILLS_REF,START_DATE,END_DATE,BASE_RATE_CODE,BASE_CCY,ALL_IN_RATE,RPA_UPDATE_DATE,RPA_UPDATE_STATUS,TENOR_CODE,ODD_FLAG,CAS,Next_BUS_DATE, ARR_RATE, SPREAD_RATE from v_RPT_ARR_RPA"
            strSQL = "select * from RPT_ARR_RPA where LastCalendarDate = (Select Run_Date from v_Get_Run_Date)"

            Dim tmpStringLength As Integer = 0
            Dim tmpStringWholeLength As Integer = 0

            myAdapter = New OleDbDataAdapter(strSQL, myConn)
            myAdapter.Fill(myDataset)
            If myDataset.Tables(0).Rows.Count > 0 Then
                'add below by terry for IBOR 2.
                myResult = WriteToLog("NOSECNAME", "ARR.getSingleDataFromHost: Total count=" & myDataset.Tables(0).Rows.Count.ToString)
                Dim IRow As Integer = 0
                For IRow = 0 To myDataset.Tables(0).Rows.Count - 1
                    strSQL = ""
                    If myDataset.Tables(0).Rows(IRow).Item("ODD_FLAG") = "Y" Then
                        'outMsg = "BSERATEOUT                                                                                                                                00000000010000000.0817000                                                                                                         OD1 00008470{                                                 "
                        myResult = WriteToLog("NOSECNAME", "BASE_CCY: " & myDataset.Tables(0).Rows(IRow).Item("BASE_CCY") & "; StartDate: " & myDataset.Tables(0).Rows(IRow).Item("START_DATE") & "; EndDate: " & myDataset.Tables(0).Rows(IRow).Item("END_DATE") & "; NextBusDate: " & myDataset.Tables(0).Rows(IRow).Item("NEXT_BUS_DATE"))
                        'getSingleDataFromHost(myDataset.Tables(0).Rows(IRow).Item("BASE_CCY"), myDataset.Tables(0).Rows(IRow).Item("START_DATE"), myDataset.Tables(0).Rows(IRow).Item("END_DATE"), myDataset.Tables(0).Rows(IRow).Item("BASE_RATE_CODE"), iRow, myDataset.Tables(0).Rows(IRow).Item("NEXT_BUS_DATE"))
                        getSingleDataFromHost(myDataset.Tables(0).Rows(IRow).Item("BASE_CCY"), myDataset.Tables(0).Rows(IRow).Item("START_DATE"), myDataset.Tables(0).Rows(IRow).Item("END_DATE"), myDataset.Tables(0).Rows(IRow).Item("BASE_RATE_CODE"), myDataset.Tables(0).Rows(IRow).Item("MQ_SEQNO_2"), myDataset.Tables(0).Rows(IRow).Item("NEXT_BUS_DATE"))

                        Dim strCAS As String = "0"
                        Dim strRATE As String = "0"
                        'Console.WriteLine("get one MQ msg done. outMsg=" & outMsg)

                        strCAS = outMsg.Substring(273, 9).Trim
                        Console.WriteLine("CAS=" & strCAS)
                        strRATE = outMsg.Substring(148, 15).Trim
                        myResult = WriteToLog("NOSECNAME", "MQ: " & outMsg)
                        myResult = WriteToLog("NOSECNAME", "ARR.getSingleDataFromHost: CODE =" & myDataset.Tables(0).Rows(IRow).Item("BASE_RATE_CODE") & " CAS=" & strCAS & " BASE RATE=" & strRATE)
                        'strSQL = "update RPT_ARR_RPA set CAS = '" & strCAS & "' , BASE_RATE = '" & strRATE & "' WHERE BILLS_REF = '" & myDataset.Tables(0).Rows(IRow).Item("BILLS_REF") & "' AND LastCalendarDate = to_char((select run_date from rpt_run_Date), 'yyyy-mm-dd')"
                        'strSQL = "update RPT_ARR_RPA set CAS = '" & strCAS & "' , ARR_RATE = '" & strRATE & "', ALL_IN_RATE = " & Val(myDataset.Tables(0).Rows(IRow).Item("SPREAD_RATE")) + Val(strCAS) + Val(strRATE) & " WHERE TRIM(BILLS_REF) = '" & Trim(myDataset.Tables(0).Rows(IRow).Item("BILLS_REF")) & "' AND LastCalendarDate = (select run_date from v_Get_Run_Date)"

                        strSQL = strSQL & "Update RPT_ARR_RPA "
                        strSQL = strSQL & "Set CAS = '" & strCAS & "' , "
                        strSQL = strSQL & "ARR_RATE = '" & strRATE & "', "
                        strSQL = strSQL & "ALL_IN_RATE = " & Val(myDataset.Tables(0).Rows(IRow).Item("SPREAD_RATE")) + Val(strCAS) + Val(strRATE) & ", "
                        strSQL = strSQL & "MQ_REQ = '" & "BASE_CCY: " & myDataset.Tables(0).Rows(IRow).Item("BASE_CCY") & "; StartDate: " & myDataset.Tables(0).Rows(IRow).Item("START_DATE") & "; EndDate: " & myDataset.Tables(0).Rows(IRow).Item("END_DATE") & "; NextBusDate: " & myDataset.Tables(0).Rows(IRow).Item("NEXT_BUS_DATE") & "', "
                        strSQL = strSQL & "MQ_REPLY = '" & outMsg & "' "
                        strSQL = strSQL & " WHERE TRIM(BILLS_REF) = '" & Trim(myDataset.Tables(0).Rows(IRow).Item("BILLS_REF")) & "' AND LastCalendarDate = (select run_date from v_Get_Run_Date)"

                        myResult = WriteToLog("NOSECNAME", "ARR.getSingleDataFromHost: " & strSQL)
                        'ErrorHandling.WriteLog("ARR.getSingleDataFromHost", strSQL)
                        myComm = myConn.CreateCommand
                        myComm.CommandType = CommandType.Text
                        myComm.CommandText = strSQL
                        myComm.ExecuteNonQuery()
                        'Console.WriteLine("update one record.")

                        'strSQL = "update RPT_ARR_RPA SET ALL_IN_RATE = TO_NUMBER(ARR_RATE) + SPREAD_RATE + TO_NUMBER(CAS) WHERE TRIM(BILLS_REF) = '" & Trim(myDataset.Tables(0).Rows(IRow).Item("BILLS_REF")) & "' AND LastCalendarDate = (select run_date from v_Get_Run_Date)"
                        'myResult = WriteToLog("NOSECNAME", "ARR.getSingleDataFromHost: " & strSQL)
                        'myComm = myConn.CreateCommand
                        'myComm.CommandType = CommandType.Text
                        'myComm.CommandText = strSQL
                        'myComm.ExecuteNonQuery()
                    End If
                Next

                myDataset.Reset()

                'strSQL = "select BILLS_REF,START_DATE,END_DATE,BASE_RATE_CODE,BASE_CCY,ALL_IN_RATE,RPA_UPDATE_DATE,RPA_UPDATE_STATUS,TENOR_CODE,ODD_FLAG,CAS,Next_BUS_DATE from v_RPT_ARR_RPA"
                strSQL = "select * from v_RPT_ARR_RPA"

                myAdapter = New OleDbDataAdapter(strSQL, myConn)
                myAdapter.Fill(myDataset)

                If myDataset.Tables(0).Columns.Count > 0 Then
                    For j = 0 To (myDataset.Tables(0).Columns.Count - 1)
                        tmpStringWholeLength = tmpStringWholeLength + Val(GlobalVarable.TXT_Col_Width(j))
                    Next
                    tmpStringWholeLength = tmpStringWholeLength + ((myDataset.Tables(0).Columns.Count - 1) * GlobalVarable.TXT_FILLER_WIDTH)
                End If
                'Create Header and Trailer Strings
                tmpStringLength = Len("00" & DateTime.Parse(GlobalVarable.RunDate).ToString(GlobalVarable.DATA_FILENAME_DATE_FORMAT_L) & DateTime.Parse(GlobalVarable.RunDate).ToString(GlobalVarable.DATA_FILENAME_DATE_FORMAT_L) & "TFS" & "RPA")
                GlobalVarable.TXT_File_Header = "00" & DateTime.Parse(GlobalVarable.RunDate).ToString(GlobalVarable.DATA_FILENAME_DATE_FORMAT_L) & DateTime.Parse(GlobalVarable.RunDate).ToString(GlobalVarable.DATA_FILENAME_DATE_FORMAT_L) & "TFS" & "RPA" & Space(tmpStringWholeLength - tmpStringLength)
                GlobalVarable.CSV_File_Header = "00" & DateTime.Parse(GlobalVarable.RunDate).ToString(GlobalVarable.DATA_FILENAME_DATE_FORMAT_L) & DateTime.Parse(GlobalVarable.RunDate).ToString(GlobalVarable.DATA_FILENAME_DATE_FORMAT_L) & "TFS" & "RPA" & Space(tmpStringWholeLength - tmpStringLength)
                tmpStringLength = Len("99" & Right("000000000000000" & myDataset.Tables(0).Rows.Count, 15))
                GlobalVarable.TXT_File_Trailer = "99" & Right("000000000000000" & myDataset.Tables(0).Rows.Count + 2, 15) & Space(tmpStringWholeLength - tmpStringLength)

                GlobalVarable.CSV_File_Trailer = "99" & Right("000000000000000" & myDataset.Tables(0).Rows.Count + 2, 15) & Space(tmpStringWholeLength - tmpStringLength)

                tmpColHeader = ""
                For i = 0 To (myDataset.Tables(0).Rows.Count - 1)
                    strOSList_csv = ""
                    strOSList_txt = ""
                    tmpCSVString = ""
                    tmpTXTString = ""

                    If myDataset.Tables(0).Columns.Count > 0 Then
                        For j = 0 To (myDataset.Tables(0).Columns.Count - 1)

                            '*******************************
                            '* Create Column Header string *
                            '*******************************
                            If i = 0 Then
                                If j = (myDataset.Tables(0).Columns.Count - 1) Then
                                    tmpColHeader = tmpColHeader & Chr(34) + Trim(GlobalVarable.CSV_Col_Header_Caption(j)) & Chr(34)
                                Else
                                    tmpColHeader = tmpColHeader & Chr(34) + Trim(GlobalVarable.CSV_Col_Header_Caption(j)) & Chr(34) & GlobalVarable.CSV_DELIMTER
                                End If
                            End If

                            If IsDBNull(myDataset.Tables(0).Rows(i).ItemArray(j)) Then
                                tmpCSVString = Chr(34) & Chr(34)
                                tmpTXTString = Space(GlobalVarable.TXT_Col_Width(j))
                            Else
                                '***************************
                                '* Formating Column String *
                                '***************************
                                Select Case UCase(GlobalVarable.TXT_Col_Type(j))
                                    Case "STRING"
                                        tmpCSVString = Chr(34) & Trim(myDataset.Tables(0).Rows(i).ItemArray(j)) & Chr(34)
                                        tmpTXTString = Left((Trim(myDataset.Tables(0).Rows(i).ItemArray(j)) & Space(GlobalVarable.TXT_Col_Width(j))), GlobalVarable.TXT_Col_Width(j))
                                    Case "NUMBER"
                                        tmpTXTString = Right((Space(GlobalVarable.TXT_Col_Width(j)) & FormatNumber(myDataset.Tables(0).Rows(i).ItemArray(j), 5)), GlobalVarable.TXT_Col_Width(j))
                                        tmpCSVString = Chr(34) & FormatNumber(myDataset.Tables(0).Rows(i).ItemArray(j), 5) & Chr(34)
                                    Case "DATE"
                                        'Below is used to retrieve from V_ARR_OSLOANLIST directly
                                        'tmpTXTString = DateTime.Parse(myDataset.Tables(0).Rows(i).ItemArray(j)).ToString(GlobalVarable.DATA_FILENAME_DATE_FORMAT_L)
                                        'tmpCSVString = Chr(34) + DateTime.Parse(myDataset.Tables(0).Rows(i).ItemArray(j)).ToString(GlobalVarable.DATA_FILENAME_DATE_FORMAT_L) & Chr(34)

                                        'Below two line is used to retrieve from RPT_ARR_RPA Table because the already in yyyymmdd format
                                        tmpTXTString = myDataset.Tables(0).Rows(i).ItemArray(j).ToString
                                        tmpCSVString = Chr(34) & myDataset.Tables(0).Rows(i).ItemArray(j).ToString & Chr(34)
                                End Select
                            End If


                            '************************************
                            '* Line up all columns into one row *
                            '************************************
                            If j = (myDataset.Tables(0).Columns.Count - 1) Then
                                strOSList_csv = strOSList_csv & tmpCSVString
                                strOSList_txt = strOSList_txt & tmpTXTString
                            Else
                                strOSList_csv = strOSList_csv & tmpCSVString & GlobalVarable.CSV_DELIMTER
                                strOSList_txt = strOSList_txt & tmpTXTString & Space(GlobalVarable.TXT_FILLER_WIDTH)
                            End If

                        Next

                        '*********************************
                        '* Append Header to Output Files *
                        '*********************************
                        If i = 0 Then
                            If GlobalVarable.Header_Trailer_TXT_Flag = "Y" Then
                                myResult = WriteTextToFile(GlobalVarable.TXT_File_Header, GlobalVarable.Output_File_Path & GlobalVarable.TXT_File_Name)
                            End If
                            If GlobalVarable.Header_Trailer_CSV_Flag = "Y" Then
                                myResult = WriteTextToFile(GlobalVarable.TXT_File_Header, GlobalVarable.Output_File_Path & GlobalVarable.CSV_File_Name)
                            End If
                            If GlobalVarable.CSV_Col_Header = "Y" Then
                                myResult = WriteTextToFile(tmpColHeader, GlobalVarable.Output_File_Path & GlobalVarable.CSV_File_Name)
                            End If
                        End If

                        '*******************************
                        '* Append Data to Output Files *
                        '*******************************
                        myResult = WriteTextToFile(strOSList_csv, GlobalVarable.Output_File_Path & GlobalVarable.CSV_File_Name)
                        myResult = WriteTextToFile(strOSList_txt, GlobalVarable.Output_File_Path & GlobalVarable.TXT_File_Name)

                        'Append Trailer to Output Files
                        If i = (myDataset.Tables(0).Rows.Count - 1) Then
                            If GlobalVarable.Header_Trailer_TXT_Flag = "Y" Then
                                myResult = WriteTextToFile(GlobalVarable.TXT_File_Trailer, GlobalVarable.Output_File_Path & GlobalVarable.TXT_File_Name)
                            End If
                            If GlobalVarable.Header_Trailer_CSV_Flag = "Y" Then
                                myResult = WriteTextToFile(GlobalVarable.TXT_File_Trailer, GlobalVarable.Output_File_Path & GlobalVarable.CSV_File_Name)
                            End If
                        End If
                    End If
                Next
            Else
                myResult = WriteTextToFile("""NO DATA""", GlobalVarable.Output_File_Path & GlobalVarable.CSV_File_Name)
                myResult = WriteTextToFile("NO DATA", GlobalVarable.Output_File_Path & GlobalVarable.TXT_File_Name)
            End If
            myResult = WriteToLog("GENERATEREPORT", "ERR00400102")
        Catch ex As Exception
            myResult = WriteToLog("GENERATEREPORT", "ERR00400103")
            myResult = WriteToLog("NOSECNAME", ex.Message)
            ErrEmailContent = ErrEmailContent & "Error: " & ex.Message & ", please check!" & vbNewLine
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
            GlobalVarable.ProgramStatus = False
            ErrEmailContent = ErrEmailContent & "Error: " & ex.Message & ", please check!" & vbNewLine
        End Try


        '************************************
        '* Send out Status Result via Email *
        '************************************

        If GlobalVarable.MailEnable = "Y" Then
            myResult = WriteToLog("EMAIL", "ERR00600101")
            If GlobalVarable.ProgramStatus Then
                If GlobalVarable.Attachment = "TXT" Then
                    ErrMsg = SendMail(GlobalVarable.MailSubject, GlobalVarable.MailBody, GlobalVarable.Output_File_Path & GlobalVarable.TXT_File_Name, True)
                Else
                    ErrMsg = SendMail(GlobalVarable.MailSubject, GlobalVarable.MailBody, GlobalVarable.Output_File_Path & GlobalVarable.CSV_File_Name, True)
                End If
                If ErrEmailContent <> "" Then
                    ErrMsg = SendMail(GlobalVarable.MailErrSubject, ErrEmailContent, GlobalVarable.ErrLogFile, False)
                End If

            Else
                ErrMsg = SendMail(GlobalVarable.MailErrSubject, ErrEmailContent, GlobalVarable.ErrLogFile, False)
            End If
            myResult = WriteToLog("EMAIL", "ERR00600103")
        End If

        myResult = WriteToLog("GENERAL", "ERR00100102")

    End Sub
    Public Function SetGlobalVariable() As Boolean

        Try

            'GlobalVarable.AppPath = "D:\Development\ARR\"
            GlobalVarable.AppPath = My.Application.Info.DirectoryPath & "\"
            'GlobalVarable.INI_DataFormat = ConfigurationManager.AppSettings("INI_DataFormat")
            GlobalVarable.ErrLogFile = GlobalVarable.AppPath & GlobalVarable.AppName & ".Log"
            GlobalVarable.INI_DataSource = ConfigurationManager.AppSettings("INI_DataSource")
            GlobalVarable.INI_Email = ConfigurationManager.AppSettings("INI_Email")
            GlobalVarable.INI_ErrCode = GlobalVarable.AppPath & ConfigurationManager.AppSettings("INI_ErrCode")

            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

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

    Public Function WriteTextToFile(strContent As String, strOutputFileName As String) As Boolean

        Dim ErrMsg As String = ""
        Dim myLogger As New Logger

        Try
            myLogger.AppendToFile(strOutputFileName, strContent)
            Return True
        Catch ex As Exception
            ErrMsg = ex.Message
            Return False
        Finally
            'Console.WriteLine(ErrMsg)
        End Try
        myLogger = Nothing

    End Function

    Public Function SendMail(ByVal strEmailSubject As String, ByVal strEmailContent As String, ByVal strAttachFileName As String, ByVal SuccessStatusFlag As Boolean) As String
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
            Dim a As New System.Net.Mail.Attachment(strAttachFileName)
            mail.Attachments.Add(a)

            Dim smtp As New SmtpClient(GlobalVarable.MailServer, GlobalVarable.MailServerPort)

            smtp.UseDefaultCredentials = False
            smtp.Send(mail)

        Catch ex As Exception
            strResult = "Error occur: " & ex.Message
        End Try

        mail = Nothing
        Return strResult

    End Function


    Public Function getSingleDataFromHost_Ori(ByVal strccy As String, ByVal strSDate As String, ByVal strEDate As String, ByVal strcode As String, ByVal strEnqID As Int32, ByVal strNDate As String) As Boolean

        Dim strbuff As String
        Dim myResult As Boolean

        Try

            strbuff = formatMQMessage(strccy, strSDate, strEDate, strcode, strEnqID, strNDate)
            If sendRequestToHost(strbuff) Then
                myResult = WriteToLog("NOSECNAME", "ARR.getSingleDataFromHost: " & strbuff)
                Return True
            Else
                Return False
                Exit Function
            End If
        Catch ex As Exception
            ErrorMessage = "Err Code: " & CStr(Err.Number) & ", Err Desc: " & Err.Description
            myResult = WriteToLog("NOSECNAME", "ARR.getSingleDataFromHost: " & ErrorMessage)
            Return False
            Exit Function
        End Try

    End Function

    Public Function getSingleDataFromHost(ByVal strccy As String, ByVal strSDate As String, ByVal strEDate As String, ByVal strcode As String, ByVal strEnqID As String, ByVal strNDate As String) As Boolean

        Dim strbuff As String
        Dim myResult As Boolean

        Try

            strbuff = formatMQMessage(strccy, strSDate, strEDate, strcode, strEnqID, strNDate)
            If sendRequestToHost(strbuff) Then
                myResult = WriteToLog("NOSECNAME", "ARR.getSingleDataFromHost: " & strbuff)
                Return True
            Else
                Return False
                Exit Function
            End If
        Catch ex As Exception
            ErrorMessage = "Err Code: " & CStr(Err.Number) & ", Err Desc: " & Err.Description
            myResult = WriteToLog("NOSECNAME", "ARR.getSingleDataFromHost: " & ErrorMessage)
            Return False
            Exit Function
        End Try

    End Function

    Public Function sendRequestToHost(ByRef strCommArea As String) As Boolean

        Dim strErrMessage As String
        Dim myResult As Boolean

        Dim mqQMgr As MQQueueManager = Nothing    '* MQQueueManager instance
        'ErrorHandling.WriteLogByFile("dim MQQueueManager. Start.") 'add by terry on 20170420
        myResult = WriteToLog("NOSECNAME", "ARR.getSingleDataFromHost: " & "dim MQQueueManager. Start.")
        Dim mqQueue As MQQueue = Nothing          '* MQQueue instance
        'ErrorHandling.WriteLogByFile("dim MQQueue. Start.") 'add by terry on 20170420
        myResult = WriteToLog("NOSECNAME", "ARR.getSingleDataFromHost: " & "dim MQQueue. Start.")
        Dim mqMsg As MQMessage = Nothing          '* MQMessage instance
        'ErrorHandling.WriteLogByFile("dim MQMessage. Start.") 'add by terry on 20170420
        myResult = WriteToLog("NOSECNAME", "ARR.getSingleDataFromHost: " & "dim MQMessage. Start.")
        Dim mqPutMsgOpts As MQPutMessageOptions   '* MQPutMessageOptions instance
        'ErrorHandling.WriteLogByFile("dim MQPutMessageOptions. Start.") 'add by terry on 20170420
        myResult = WriteToLog("NOSECNAME", "ARR.getSingleDataFromHost: " & "dim MQPutMessageOptions. Start.")
        Dim mqGetMsgOpts As MQGetMessageOptions   '* MQGetMessageOptions instance
        'ErrorHandling.WriteLogByFile("dim MQGetMessageOptions. Start.") 'add by terry on 20170420
        myResult = WriteToLog("NOSECNAME", "ARR.getSingleDataFromHost: " & "dim MQGetMessageOptions. Start.")

        Dim byteMsgID As Byte()                   '* Message ID
        Dim strMsgID As String = ""

        Try
            'ErrorHandling.WriteLogByFile("Send MQ message. Start.") 'add by terry on 20170420
            myResult = WriteToLog("NOSECNAME", "Send MQ message. Start.")

            mqQMgr = New MQQueueManager(GlobalVarable.MQ_Manager)
        Catch mqe As IBM.WMQ.MQException
            strErrMessage = "Fail to connect " + GlobalVarable.MQ_Manager + ", ErrorCode: " + CStr(mqe.Reason) + ", ErrorDesc: " + mqe.Message
            'ErrorHandling.WriteLog("CMHOnline.sendRequestToHost", strErrMessage)
            myResult = WriteToLog("NOSECNAME", "CMHOnline.sendRequestToHost: " & strErrMessage)
            'MsgBox(strErrMessage)
            Return False
            Exit Function
        End Try

        Try
            mqQueue = mqQMgr.AccessQueue(GlobalVarable.MQ_QUEUE_REQ, MQC.MQOO_OUTPUT + MQC.MQOO_FAIL_IF_QUIESCING)

        Catch mqe As IBM.WMQ.MQException
            strErrMessage = "fail to access " + GlobalVarable.MQ_QUEUE_REQ + " , ErrorCode: " + CStr(mqe.Reason) + ", ErrorDesc: " + mqe.Message
            'ErrorHandling.WriteLog("CMHOnline.sendRequestToHost", strErrMessage)
            myResult = WriteToLog("NOSECNAME", "CMHOnline.sendRequestToHost: " & strErrMessage)
            'MsgBox(strErrMessage)
            Return False
            Exit Function
        End Try

        Console.WriteLine("Input CommArea Len:" + CStr(Len(strCommArea)))
        Console.WriteLine(strCommArea)
        Console.WriteLine("Connect MQ done.")

        mqMsg = New MQMessage

        mqMsg.WriteString(strCommArea)
        mqMsg.Format = MQC.MQFMT_STRING
        'mqMsg.Write(Encoding.Unicode.GetBytes(strCommArea))
        'mqMsg.Write(Text.Encoding.GetEncoding("UTF-16LE").GetBytes(strCommArea))
        mqPutMsgOpts = New MQPutMessageOptions

        Try
            mqQueue.Put(mqMsg, mqPutMsgOpts)
        Catch mqe As IBM.WMQ.MQException
            strErrMessage = "fail to put message, ErrorCode: " + CStr(mqe.Reason) + ", ErrorDesc: " + mqe.Message
            'ErrorHandling.WriteLog("CMHOnline.sendRequestToHost", strErrMessage)
            myResult = WriteToLog("NOSECNAME", "CMHOnline.sendRequestToHost: " & strErrMessage)
            MsgBox(strErrMessage)
            Return False
            Exit Function
        End Try
        'ErrorHandling.WriteLogByFile("Send MQ message. End.") 'add by terry on 20170420
        myResult = WriteToLog("NOSECNAME", "Send MQ message. End.")
        byteMsgID = mqMsg.MessageId

        Dim enc As New System.Text.ASCIIEncoding
        strMsgID = enc.GetString(byteMsgID)

        'ErrorHandling.WriteLogByFile("Get MQ message. Start.") 'add by terry on 20170420
        myResult = WriteToLog("NOSECNAME", "Get MQ message. Start.")
        Try
            mqQueue = mqQMgr.AccessQueue(GlobalVarable.MQ_QUEUE_RPLY, MQC.MQOO_INPUT_AS_Q_DEF + MQC.MQOO_FAIL_IF_QUIESCING)
        Catch mqe As IBM.WMQ.MQException
            strErrMessage = "fail to access " + GlobalVarable.MQ_QUEUE_RPLY + ", ErrorCode: " + CStr(mqe.Reason) + ", ErrorCode: " + mqe.Message
            'ErrorHandling.WriteLog("CMHOnline.sendRequestToHost", strErrMessage)
            myResult = WriteToLog("NOSECNAME", "CMHOnline.sendRequestToHost: " & strErrMessage)
            MsgBox(strErrMessage)
            Return False
            Exit Function
        End Try

        mqMsg = New MQMessage
        mqGetMsgOpts = New MQGetMessageOptions
        mqGetMsgOpts.WaitInterval = GlobalVarable.MQ_Wait_Interval
        mqGetMsgOpts.Options += MQC.MQGMO_WAIT

        Try
            mqQueue.Get(mqMsg, mqGetMsgOpts)
            strCommArea = mqMsg.ReadString(mqMsg.MessageLength)
            'strCommArea = Encoding.Unicode.GetString(mqMsg.ReadBytes((Len(strCommArea) + 300) * 2))
            'strCommArea = Encoding.Unicode.GetString(mqMsg.ReadBytes(Len(strCommArea)))
            outMsg = strCommArea 'add below by terry for IBOR 2. on 20220330

            Console.WriteLine("Return CommArea Len:" + CStr(Len(strCommArea)))
            Console.WriteLine(strCommArea)
        Catch mqe As IBM.WMQ.MQException
            strErrMessage = "fail to get Msg" + ", ErrorCode: " + CStr(mqQueue.ReasonCode) + ", ErrorCode: " + mqe.Message
            'ErrorHandling.WriteLog("CMHOnline.sendRequestToHost", strErrMessage)
            myResult = WriteToLog("NOSECNAME", "CMHOnline.sendRequestToHost: " & strErrMessage)
            MsgBox(strErrMessage)
            Return False
            Exit Function
        End Try
        'ErrorHandling.WriteLogByFile("Get MQ message. End.") 'add by terry on 20170420
        myResult = WriteToLog("NOSECNAME", "Get MQ message. End.")

        Try
            mqQueue.Close()
            mqQueue = Nothing
        Catch mqe As IBM.WMQ.MQException
            strErrMessage = "Close Queue" + ", RC: " + CStr(mqe.Reason) + ", Error: " + mqe.Message
            'ErrorHandling.WriteLog("CMHOnline.sendRequestToHost", strErrMessage)
            myResult = WriteToLog("NOSECNAME", "CMHOnline.sendRequestToHost: " & strErrMessage)
            MsgBox(strErrMessage)
            Return False
            Exit Function
        End Try

        If Not (mqQMgr Is Nothing) Then
            If mqQMgr.IsConnected Then
                mqQMgr.Disconnect()
            End If
        End If
        Return True
    End Function


    Private Function formatMQMessage_Ori(ByVal strccy As String, ByVal strSDate As String, ByVal strEDate As String, ByVal strcode As String, ByVal strEnqID As Int32, ByVal strNDate As String) As String

        Dim strBuff As String
        Dim strCurDate As String
        Dim myResult As Boolean

        strBuff = ""
        Try
            '"BASERATEQ                                                                                                                                 0000561141B.00202111010020211201USD"
            strCurDate = Format(Now, "yyyy-MM-dd")

            strBuff = "BASERATEQ2"
            strBuff = strBuff + Space(128)
            strBuff = strBuff + (strEnqID + 1).ToString("0000000000")
            strBuff = strBuff + strcode
            strBuff = strBuff + "00" + strSDate.Replace("-", "").Replace("/", "")
            strBuff = strBuff + "00" + strNDate.Replace("-", "").Replace("/", "")
            strBuff = strBuff + strccy
            strBuff = strBuff + "00" + strSDate.Replace("-", "").Replace("/", "")
            strBuff = strBuff + "00" + strEDate.Replace("-", "").Replace("/", "")


            'ErrorHandling.WriteLog("formatMQMessage", strBuff)
            myResult = WriteToLog("NOSECNAME", "formatMQMessage: " & strBuff)
        Catch ex As Exception
            ErrorMessage = "Err Code: " & CStr(Err.Number) & ", Err Desc: " & Err.Description
            'ErrorHandling.WriteLog("ARR.formatMQMessage", ErrorMessage)
            myResult = WriteToLog("NOSECNAME", "formatMQMessage: " & ErrorMessage)
        End Try

        Return strBuff

    End Function

    Private Function formatMQMessage(ByVal strccy As String, ByVal strSDate As String, ByVal strEDate As String, ByVal strcode As String, ByVal strEnqID As String, ByVal strNDate As String) As String

        Dim strBuff As String
        Dim strCurDate As String
        Dim myResult As Boolean

        strBuff = ""
        Try
            '"BASERATEQ                                                                                                                                 0000561141B.00202111010020211201USD"
            strCurDate = Format(Now, "yyyy-MM-dd")

            strBuff = "BASERATEQ2"
            strBuff = strBuff + Space(128)
            strBuff = strBuff + strEnqID
            strBuff = strBuff + strcode
            strBuff = strBuff + "00" + strSDate.Replace("-", "").Replace("/", "")
            strBuff = strBuff + "00" + strNDate.Replace("-", "").Replace("/", "")
            strBuff = strBuff + strccy
            strBuff = strBuff + "00" + strSDate.Replace("-", "").Replace("/", "")
            strBuff = strBuff + "00" + strEDate.Replace("-", "").Replace("/", "")


            'ErrorHandling.WriteLog("formatMQMessage", strBuff)
            myResult = WriteToLog("NOSECNAME", "formatMQMessage: " & strBuff)
        Catch ex As Exception
            ErrorMessage = "Err Code: " & CStr(Err.Number) & ", Err Desc: " & Err.Description
            'ErrorHandling.WriteLog("ARR.formatMQMessage", ErrorMessage)
            myResult = WriteToLog("NOSECNAME", "formatMQMessage: " & ErrorMessage)
        End Try

        Return strBuff

    End Function



    Private Function myTrim(ByVal strValue As String)
        strValue = Replace(strValue, Chr(0), Space(1))
        strValue = Replace(strValue, "'", "''")
        strValue = Replace(strValue, "&", "&'||'")
        strValue = Trim(strValue)

        Return strValue
    End Function

    Private Function CheckEngStr(ByVal strValue As String) As Boolean

        Dim ChkResult As Boolean
        Dim myResult As Boolean

        ChkResult = True
        For i = 0 To strValue.Length - 1
            If Not Char.IsLetter(strValue.Chars(i)) And Not Char.IsNumber(strValue.Chars(i)) And Not Char.IsSymbol(strValue.Chars(i)) And Not Char.IsSeparator(strValue.Chars(i)) And Not Char.IsPunctuation(strValue.Chars(i)) And Not Char.IsSymbol(strValue.Chars(i)) Then
                ChkResult = False
                'ErrorHandling.WriteLog("CheckEngStr: ", strValue.Chars(i)) ' Cyn
                myResult = WriteToLog("NOSECNAME", "CheckEngStr: " & strValue.Chars(i))
                Return ChkResult
            End If
        Next

        Return ChkResult
    End Function

    Private Function formatInt(ByVal IntIn As Int32, ByVal totalLength As Int16) As String
        Dim strOut As String

        strOut = IntIn.ToString("0000000000")


        Return ""
    End Function
End Module
