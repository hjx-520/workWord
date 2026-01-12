Imports log4net
Imports System.IO
Imports BEA.TFS.CommonUtility
Imports BEA.TFS.Common.ExtensionMethods

Public Class Utility

    Public Shared _logger As ILog = LogManager.GetLogger(GetType(Utility).ToString)

    Public Shared Function GetConnectionStr(dataSource As String, userID As String, password As String) As String
        Return $"Provider=OraOLEDB.Oracle;Data Source={dataSource};User Id={userID};Password={password};"
    End Function

    ' TODO: [REMOVE COMMENT AFTER REVIEW] Made utility methods that causes data changes to take IDbTransaction as a required parameter.
    ' TODO: [REMOVE COMMENT AFTER REVIEW] Made utility methods that involves only queries to take IDbTransaction as an optional parameter.
    ' TODO: Use RptExcpt entity CRUD methods to replace.
    Friend Shared Sub DeleteRptExcpt(conn As IDbConnection, tran As IDbTransaction)
        Using cmd As IDbCommand = conn.CreateCommand
            cmd.ExecuteNonQuery(
                AppConst.DELETE_RPT_EXCPT_SQL,
                New IDbDataParameter() {},
                tran
            )
        End Using
        _logger.Debug($"Delete Table RPT_EXCEPT is Done")
    End Sub

    ' TODO: Use RptExcpt entity CRUD methods to replace.
    Friend Shared Sub InsertRptExcpt(errData As Forf_Upload_Prop.TB_RPT_EXCPT, conn As IDbConnection, tran As IDbTransaction)
        Using cmd As IDbCommand = conn.CreateCommand
            cmd.ExecuteNonQuery(
                AppConst.INSERT_INTO_RPT_EXCPT,
                New IDbDataParameter() {
                    cmd.CreateParameter("BILLS_REF", errData.BillsRef),
                    cmd.CreateParameter("ERROR_MESSAGE", errData.ErrorMessage)
                },
                tran
            )
        End Using
    End Sub

    Friend Shared Sub InsertForfaitingMaster(conn As IDbConnection, tran As IDbTransaction)
        DeleteRptForfMaster(conn, tran)
        DataPrepareRptForfMaster(conn, tran)
    End Sub

    ' TODO: Use RptForfMaster entity CRUD methods to replace.
    Friend Shared Sub DeleteRptForfMaster(conn As IDbConnection, tran As IDbTransaction)
        Using cmd As IDbCommand = conn.CreateCommand
            cmd.ExecuteNonQuery(
                AppConst.DELETE_RPT_FORF_MASTER_SQL,
                New IDbDataParameter() {},
                tran
            )
        End Using
        _logger.Debug($"Delete Table RPT_FORF_MASTER is Done")
    End Sub

    ' TODO: Use RptForfMaster entity CRUD methods to replace.
    Friend Shared Sub UpdateRptForfMaster(conn As IDbConnection, tran As IDbTransaction)
        Using cmd As IDbCommand = conn.CreateCommand
            cmd.ExecuteNonQuery(
                AppConst.UPDATE_RPT_FORF_MASTER_SQL,
                New IDbDataParameter() {},
                tran
            )
        End Using
        _logger.Debug($"Update Table RPT_FORF_MASTER is Done")
    End Sub

    ' TODO: Use RptForfMaster entity CRUD methods to replace.
    Friend Shared Sub InsertRptForfMaster(recordData As Forf_Upload_Prop.TB_RPT_FORF_MASTER, conn As IDbConnection, tran As IDbTransaction)
        Using cmd As IDbCommand = conn.CreateCommand
            cmd.ExecuteNonQuery(
                AppConst.INSERT_RPT_FORF_MASTER_SQL,
                New IDbDataParameter() {
                    cmd.CreateParameter("BILLS_REF", recordData.BillsRef),
                    cmd.CreateParameter("TENOR", recordData.Tenor),
                    cmd.CreateParameter("MODEL_TYPE", recordData.ModelType),
                    cmd.CreateParameter("CURRENCY", recordData.BillsCcy),
                    cmd.CreateParameter("LOAN_PRINCIPAL", recordData.OsBalFcy),
                    cmd.CreateParameter("LOAN_START_DATE", recordData.CREATE_DATE),
                    cmd.CreateParameter("LOAN_DUE_DATE", recordData.DUE_DATE),
                    cmd.CreateParameter("INTEREST_RATE", recordData.INT_RATE)
                },
                tran
            )
        End Using
    End Sub

    ' TODO: Create entity for RptForfMaster.
    Friend Shared Sub DataPrepareRptForfMaster(conn As IDbConnection, tran As IDbTransaction)
        Using cmd As IDbCommand = conn.CreateCommand,
                rs As IDataReader = cmd.ExecuteReader(
                    AppConst.SELECT_DATAPREPARE_RPT_FORF_MASTER_SQL,
                    New IDbDataParameter() {},
                    tran
                )
            While rs.Read

                ' TODO: [REMOVE COMMENT AFTER REVIEW] Always instantiate a new object at the beginning of every iteration.
                Dim entity As New Forf_Upload_Prop.TB_RPT_FORF_MASTER


                entity.BillsRef = rs.GetString("BILLS_REF")



                entity.Tenor = rs.GetString("TENOR")



                entity.ModelType = rs.GetString("MODEL_TYPE")


                entity.BillsCcy = rs.GetString("BILLS_CCY")


                entity.OsBalFcy = rs.GetNullableDecimal("OS_BAL_FCY")


                entity.CREATE_DATE = rs.GetNullableDateTime("CREATE_DATE")


                entity.DUE_DATE = rs.GetNullableDateTime("DUE_DATE")


                entity.INT_RATE = rs.GetNullableDecimal("INT_RATE")


                InsertRptForfMaster(entity, conn, tran)

            End While
        End Using
    End Sub

    Public Shared Function IsExistDetail(billsRef As String, conn As IDbConnection, Optional tran As IDbTransaction = Nothing) As Boolean
        Using cmd As IDbCommand = conn.CreateCommand,
                rs As IDataReader = cmd.ExecuteReader(
                    AppConst.SELECT_BILLS_REF_FROM_RPT_IFP_BILLS_DETAIL_MONTHEND,
                    New IDbDataParameter() {
                        cmd.CreateParameter("BILLS_REF", billsRef)
                    },
                    tran
                )
            Return rs.Read
        End Using
    End Function

    Public Shared Function IsExistForf(Bill_Ref As String, conn As IDbConnection, Optional tran As IDbTransaction = Nothing) As Boolean
        Using cmd As IDbCommand = conn.CreateCommand,
                rs As IDataReader = cmd.ExecuteReader(
                    AppConst.SELECT_RPT_FORF_UPLOAD_BY_BILLS_REF,
                    New IDbDataParameter() {
                        cmd.CreateParameter("BILLS_REF", Bill_Ref)
                    },
                    tran
                )
            Return rs.Read
        End Using
    End Function

    ' TODO: Use RptForfUpload entity CRUD methods to replace.
    Friend Shared Sub UpdateRptForfUpload(billsRef As String, conn As IDbConnection, tran As IDbTransaction)
        Using cmd As IDbCommand = conn.CreateCommand
            cmd.ExecuteNonQuery(
                AppConst.UPDATE_RPT_FORF_UPLOAD_VALID_IND,
                New IDbDataParameter() {
                    cmd.CreateParameter("BILLS_REF", billsRef)
                },
                tran
            )
        End Using

        _logger.Debug($"Update Table RPT_FORF_UPLOAD is Done")
    End Sub

    ' For tenor_code
    Public Shared Function IsTenorCodeNull(billsRef As String, conn As IDbConnection, tran As IDbTransaction) As Boolean
        Using cmd As IDbCommand = conn.CreateCommand,
                rs As IDataReader = cmd.ExecuteReader(
                    AppConst.SELECT_FROM_RPT_IFP_BILLS_DETAIL_MONTHEND_BY_TENOR_CODE,
                    New IDbDataParameter() {
                        cmd.CreateParameter("BILLS_REF", billsRef)
                    },
                    tran
                )
            Return rs.Read
        End Using
    End Function

    Friend Shared Sub ValidateRepriceForPeriod(conn As IDbConnection, tran As IDbTransaction)
        Using cmd As IDbCommand = conn.CreateCommand,
                rs As IDataReader = cmd.ExecuteReader(
                    AppConst.SELECT_FROM_RPT_REPRICE_UPLOAD_GROUP_BY_BILLS_REF,
                    New IDbDataParameter() {},
                    tran
                )
            While rs.Read
                If Not String.IsNullOrEmpty(rs.GetString("BILLS_REF")) Then
                    Dim billsRef As String = rs.GetString("BILLS_REF").Trim
                    ' TODO: [REMOVE COMMENT AFTER REVIEW] Refactored away method IsContinuePeriod due to unclear boundary of responsibilities. If it checks something, it should be made a Function returning Boolean; if it has side effects, those should be separated.
                    ' For continue period case 3
                    ' Check if start date is equals to previous record last year
                    ' Only one bills_ref pass in SQL and group by function using in SQL
                    ' Only one record or no record return from SQL
                    Using cmdContinuePeriod As IDbCommand = conn.CreateCommand,
                rsContinuePeriod As IDataReader = cmdContinuePeriod.ExecuteReader(
                    AppConst.SELECT_IS_CONTINUE_PERIOD_SQL,
                    New IDbDataParameter() {
                        cmdContinuePeriod.CreateParameter("BILLS_REF", billsRef)
                    },
                    tran
                )
                        If rsContinuePeriod.Read Then
                            ' TODO: Convert class to data entity.
                            Dim entity As New Forf_Upload_Prop.TB_RPT_EXCPT With {
                    .BillsRef = rsContinuePeriod.GetString("BILLS_REF").Trim,
                    .ErrorMessage = "Reprice period end date is not equal to start date of next reprice period"
                }
                            InsertRptExcpt(entity, conn, tran)
                            UpdateRptForfUpload(billsRef, conn, tran)
                        End If
                    End Using
                End If
            End While
        End Using
    End Sub

    Friend Shared Sub RunProcess(loaderPath As String, Optional args As String = Nothing)
        Dim proc As Process = Nothing
        proc = New Process()
        Try
            proc.StartInfo.WorkingDirectory = Path.GetDirectoryName(loaderPath)
            proc.StartInfo.FileName = Path.GetFileName(loaderPath)
            proc.StartInfo.Arguments = args
            proc.StartInfo.CreateNoWindow = False
            proc.Start()
            proc.WaitForExit()
            _logger.Info(loaderPath & " Finished")
        Catch ex As Exception
            _logger.Info(loaderPath & " Run with error " & proc.ExitCode.ToString())
        End Try
    End Sub

    Friend Shared Sub ValidateFullSet(conn As IDbConnection, tran As IDbTransaction)
        ' Check if any bill_ref in table rpt_forf_refn_monthend and rpt_bills_detail_monthend
        ' But not in table rpt_forf_upload
        Using cmd As IDbCommand = conn.CreateCommand,
                rs As IDataReader = cmd.ExecuteReader(
                    AppConst.SELECT_FROM_RPT_FORF_REFN_MONTHEND_RPT_BILLS_DETAIL_MONTHEND_RPT_FORF_UPLOAD,
                    New IDbDataParameter() {},
                    tran
                )
            While rs.Read
                Dim errData As New Forf_Upload_Prop.TB_RPT_EXCPT With {
                    .BillsRef = rs.GetString("BILLS_REF").Trim,
                    .ErrorMessage = "Bills reference was not found in upload file"
                }
                InsertRptExcpt(errData, conn, tran)
            End While
        End Using
    End Sub

    Friend Shared Sub ValidateCurve(conn As IDbConnection, tran As IDbTransaction)

        Using cmd As IDbCommand = conn.CreateCommand,
                rs As IDataReader = cmd.ExecuteReader(
                    AppConst.SELECT_FROM_RPT_FORF_UPLOAD_RPT_IFP_BILLS_DETAIL_MONTHEND,
                    New IDbDataParameter() {},
                    tran
                )
            While rs.Read
                Dim billsRef As String = rs.GetString("BILLS_REF").Trim
                Dim tenor As String = rs.GetString("TENOR").Trim
                Dim modelType As String = rs.GetString("MODEL_TYPE", "0").Trim
                Dim billsCcy As String = rs.GetString("BILLS_CCY").Trim

                Dim errData As New Forf_Upload_Prop.TB_RPT_EXCPT With {
                    .BillsRef = billsRef
                }

                If modelType = "2" OrElse modelType = "4" Then
                    If GetCurveName(billsCcy, tenor, "F", conn, tran) = Nothing Then
                        errData.ErrorMessage = "Curve is not found"
                        InsertRptExcpt(errData, conn, tran)
                        UpdateRptForfUpload(billsRef, conn, tran)
                    End If
                End If

                If modelType <> Nothing AndAlso GetCurveName(billsCcy, tenor, "D", conn, tran) = Nothing Then
                    errData.ErrorMessage = "Curve is not found"
                    InsertRptExcpt(errData, conn, tran)
                    UpdateRptForfUpload(billsRef, conn, tran)
                End If
                ' For model 2 or 4
            End While
        End Using

    End Sub

    Public Shared Function GetCurveName(ccy As String, tenorCode As String, usage As Char, conn As IDbConnection, Optional tran As IDbTransaction = Nothing) As String

        Using cmd As IDbCommand = conn.CreateCommand
            GetCurveName = cmd.ExecuteScalar(
                AppConst.SELECT_FROM_RPT_CURVE_USAGE,
                New IDbDataParameter() {
                    cmd.CreateParameter("currencyCode", ccy),
                    cmd.CreateParameter("usage", usage),
                    cmd.CreateParameter("skipTenorCodeCondition", If(usage = "D"c, 1, 0)),
                    cmd.CreateParameter("tenorCode", tenorCode)
                },
                tran
            )
        End Using

        Return GetCurveName
    End Function

    Friend Shared Sub GenRptSwapCurveCsv(billsRef As String, swapCurveDate As Date, swapCurveDatFile As String, swapCurveJPYDatFile As String, swapCurveCsvFile As String, conn As IDbConnection, Optional tran As IDbTransaction = Nothing)
        Dim hDateJpy As Date = Nothing
        Dim countJpy As Integer = 0
        Dim k As Integer = 0
        Dim decCheckSum As Decimal
        Dim decCalCheckSum As Decimal

        Using sr As New StreamReader(swapCurveDatFile),
                sw As New StreamWriter(swapCurveCsvFile, True)
            Dim hDate As Date = Nothing

            Dim lineCount As Integer = 0
            Dim recordCount As Integer = 0
            While Not sr.EndOfStream

                Dim line As String = sr.ReadLine.PadRight(84)
                If line.Trim.StartsWith("00") Then

                    hDate = Date.ParseExact(line.Substring(2, 8), "yyyyMMdd", Globalization.CultureInfo.InvariantCulture)

                ElseIf line.Trim.StartsWith("99") Then

                    recordCount = line.Substring(2, 12)
                    decCheckSum = Convert.ToDecimal(line.Substring(14, 16))

                ElseIf line.Trim.StartsWith("05") Then

                    Dim curveName As String = line.Substring(2, 16)
                    Dim ccy As String = line.Substring(18, 3)
                    Dim startDate As String = line.Substring(21, 4) & line.Substring(25, 2) & line.Substring(27, 2)
                    Dim endDate As String = line.Substring(29, 4) & line.Substring(33, 2) & line.Substring(35, 2)
                    Dim discountFactor As String = line.Substring(37, 47)
                    Dim content As String = $"""{curveName}"",""{ccy}"",""{startDate}"",""{endDate}"",""{discountFactor}"""
                    lineCount += 1


                    decCalCheckSum += Convert.ToDecimal(discountFactor)
                    sw.WriteLine(content)

                End If

            End While

            If recordCount <> lineCount OrElse hDate <> swapCurveDate Then

                InsertExcp(billsRef, $"KPS curve is not valid! Record Count: {recordCount} : {lineCount},  Curve Date: {hDate.ToString("yyyyMMdd")} : {swapCurveDate}", conn, tran)
                Throw New Exception($"KPS to TFS yield curve interface file is not valid!")
            End If

        End Using


        If decCheckSum <> decCalCheckSum Then

            InsertExcp(billsRef, $"Check Sum of KPS curve is not correct: {decCheckSum} : {decCalCheckSum}", conn, tran)
            Throw New Exception($"Check Sum of KPS curve Is Not correct: {decCheckSum} : {decCalCheckSum}")
        End If

        ' JPY
        Using sr As New StreamReader(swapCurveJPYDatFile),
                sw As New StreamWriter(swapCurveCsvFile, True)
            Dim hDate As Date = Nothing

            Dim lineCount As Integer = 0
            Dim recordCount As Integer = 0
            While Not sr.EndOfStream

                Dim line As String = sr.ReadLine.PadRight(47)
                If line.Trim.StartsWith("00") Then

                    hDate = Date.ParseExact(line.Substring(2, 8), "yyyyMMdd", Globalization.CultureInfo.InvariantCulture)

                ElseIf line.Trim.StartsWith("99") Then

                    recordCount = line.Substring(2, 12)

                ElseIf line.Trim.StartsWith("05") Then

                    Dim curveName As String = line.Substring(2, 16)
                    Dim ccy As String = line.Substring(18, 3)
                    Dim startDate As String = line.Substring(21, 4) & line.Substring(25, 2) & line.Substring(27, 2)
                    Dim endDate As String = line.Substring(29, 4) & line.Substring(33, 2) & line.Substring(35, 2)
                    Dim discountFactor As String = line.Substring(37, 11)
                    Dim content As String = $"""{curveName}"",""{ccy}"",""{startDate}"",""{endDate}"",""{discountFactor}"""
                    lineCount += 1
                    sw.WriteLine(content)

                End If

            End While

            If recordCount <> lineCount OrElse hDate <> swapCurveDate Then
                InsertExcp(billsRef, $"TCL to TFS JPY yield curve interface file is not valid! {recordCount} {lineCount} {hDate} {swapCurveDate}", conn, tran)

                Throw New Exception($"TCL to TFS JPY yield curve interface file is not valid!{recordCount} {lineCount} {hDate} {swapCurveDate}")
            End If

        End Using
    End Sub

    Friend Shared Sub RemoveKPSFiles(swapCurveFtpPath As String)

        ' TODO: [REMOVE COMMENT AFTER REVIEW] Added file pattern filter to avoid deleting irrelevant files.
        For Each fi As FileInfo In New DirectoryInfo(Path.GetDirectoryName(swapCurveFtpPath)).GetFiles("KPSTOTFS*.DAT")
            If fi.FullName <> swapCurveFtpPath Then
                File.Delete(fi.FullName)
            End If
        Next

    End Sub

    Friend Shared Sub RemoveIS9Files(provisionFtpPath As String)

        ' TODO: [REMOVE COMMENT AFTER REVIEW] Added file pattern filter to avoid deleting irrelevant files.
        For Each fi As FileInfo In New DirectoryInfo(Path.GetDirectoryName(provisionFtpPath)).GetFiles("IS9TOTFS_PROVISION*.DAT")
            If fi.FullName <> provisionFtpPath Then
                File.Delete(fi.FullName)
            End If
        Next

    End Sub

    Friend Shared Sub ValidateReprice(conn As IDbConnection, tran As IDbTransaction)
        Using cmd As IDbCommand = conn.CreateCommand,
                rs As IDataReader = cmd.ExecuteReader(
                    AppConst.SELECT_FROM_RPT_REPRICE_UPLOAD,
                    New IDbDataParameter() {},
                    tran
                )
            Dim prevBillsRef As String = Nothing
            Dim billsRefError As Boolean = False
            Dim dateError As Boolean = False

            While rs.Read
                Dim billsRef As String = rs.GetString("BILLS_REF")

                ' Check if new Bills_type
                If prevBillsRef <> billsRef Then
                    dateError = False
                    billsRefError = False
                End If

                Dim startDate As Date? = rs.GetNullableDateTime("START_DATE")

                Dim endDate As Date? = rs.GetNullableDateTime("END_DATE")

                Dim errData As New Forf_Upload_Prop.TB_RPT_EXCPT With {
                    .BillsRef = billsRef
                }

                If Not (IsExistDetail(billsRef, conn, tran) AndAlso IsExistForf(billsRef, conn, tran)) AndAlso Not billsRefError Then
                    errData.ErrorMessage = "Bills reference was not found in RPT_FORF_UPLOAD or RPT_IFP_BILLS_DETAIL_MONTHEND"
                    billsRefError = True
                    InsertRptExcpt(errData, conn, tran)
                    UpdateRptForfUpload(billsRef, conn, tran)
                End If

                If startDate.HasValue AndAlso endDate.HasValue AndAlso endDate <= startDate And Not dateError Then
                    errData.ErrorMessage = "Reprice Start Date >= End Date"
                    dateError = True
                    InsertRptExcpt(errData, conn, tran)
                    UpdateRptForfUpload(billsRef, conn, tran)
                End If

                prevBillsRef = billsRef
            End While
        End Using
    End Sub

    Friend Shared Sub ValidateRepriceForOthers(conn As IDbConnection, tran As IDbTransaction)
        Using cmd As IDbCommand = conn.CreateCommand,
                rs As IDataReader = cmd.ExecuteReader(
                    AppConst.SELECT_FROM_RPT_FORF_UPLOAD_BY_MODEL_TYPE,
                    New IDbDataParameter() {},
                    tran
                )
            While rs.Read
                Dim billsRef As String = rs.GetString("BILLS_REF")
                Dim modelType As String = rs.GetString("MODEL_TYPE")

                Dim errData As New Forf_Upload_Prop.TB_RPT_EXCPT With {
                    .BillsRef = billsRef
                }

                Select Case modelType
                    Case "2", "4"
                        If Not ExistsInReprice(billsRef, conn, tran) Then
                            errData.ErrorMessage = "Bills reference " & billsRef & " has no reprice data"
                            InsertRptExcpt(errData, conn, tran)
                            UpdateRptForfUpload(billsRef, conn, tran)
                        End If
                    ' Case 6
                    Case "1", "3"
                        If ExistsInReprice(billsRef, conn, tran) Then
                            errData.ErrorMessage = "Bills reference " & billsRef & " of model 1 or 3 has reprice data inputted"
                            InsertRptExcpt(errData, conn, tran)
                            UpdateRptForfUpload(billsRef, conn, tran)
                        End If
                End Select

                ' Case 7
                If modelType = "2" AndAlso Not EndDatesAgree(billsRef, conn, tran) Then
                    errData.ErrorMessage = "Bills reference " & billsRef & " Reprice period end date does not match the due date of the loan"
                    InsertRptExcpt(errData, conn, tran)
                    UpdateRptForfUpload(billsRef, conn, tran)
                End If

                ' Case 8
                If IsExistDetail(billsRef, conn, tran) AndAlso IsTenorCodeNull(billsRef, conn, tran) Then
                    errData.ErrorMessage = "Tenor code is null"
                    InsertRptExcpt(errData, conn, tran)
                    UpdateRptForfUpload(billsRef, conn, tran)
                End If
            End While
        End Using
    End Sub

    Public Shared Function ExistsInReprice(ByVal Bill_Ref As String, conn As IDbConnection, tran As IDbTransaction) As Boolean
        Using cmd As IDbCommand = conn.CreateCommand,
                rs As IDataReader = cmd.ExecuteReader(
                    AppConst.SELECT_IS_EXIST_IN_REPRICE_SQL,
                    New IDbDataParameter() {
                        cmd.CreateParameter("BILLS_REF", Bill_Ref)
                    },
                    tran
                )
            Return rs.Read
        End Using
    End Function

    Public Shared Function EndDatesAgree(billsRef As String, conn As IDbConnection, Optional tran As IDbTransaction = Nothing) As Boolean
        Using cmdEndDate As IDbCommand = conn.CreateCommand,
                cmdDueDate As IDbCommand = conn.CreateCommand
            Dim endDate As Date? = cmdEndDate.ExecuteScalar(
                AppConst.SELECT_END_DATE_FROM_RPT_REPRICE_UPLOAD,
                New IDbDataParameter() {
                    cmdEndDate.CreateParameter("BILLS_REF", billsRef)
                },
                tran
            )
            Dim dueDate As Date? = cmdDueDate.ExecuteScalar(
                AppConst.SELECT_DUE_DATE_FROM_RPT_IFP_BILLS_DETAIL_MONTHEND,
                New IDbDataParameter() {
                    cmdEndDate.CreateParameter("BILLS_REF", billsRef)
                },
                tran
            )

            Return endDate.HasValue AndAlso dueDate.HasValue AndAlso endDate = dueDate
        End Using
    End Function

    Friend Shared Sub ValidateForfaiting(conn As IDbConnection, tran As IDbTransaction)
        DeleteRptExcpt(conn, tran)

        Using cmd As IDbCommand = conn.CreateCommand,
                rs As IDataReader = cmd.ExecuteReader(
                    AppConst.SELECT_FROM_RPT_FORF_UPLOAD,
                    New IDbDataParameter() {},
                    tran
                )
            While rs.Read
                Dim billsRef As String = rs.GetString("BILLS_REF")
                Dim modelType As String = rs.GetString("MODEL_TYPE")

                Dim errData As New Forf_Upload_Prop.TB_RPT_EXCPT With {
                    .BillsRef = billsRef
                }

                If Not IsExistDetail(billsRef, conn, tran) Then
                    errData.ErrorMessage = "Bills reference was not found in RPT_IFP_BILLS_DETAIL_MONTHEND"
                    InsertRptExcpt(errData, conn, tran)
                    UpdateRptForfUpload(billsRef, conn, tran)
                End If

                If Not IsNumeric(modelType) OrElse Not (modelType > 0 AndAlso modelType < 5) Then
                    errData.ErrorMessage = "MODEL TYPE " & modelType & " is invalid"
                    InsertRptExcpt(errData, conn, tran)
                    UpdateRptForfUpload(billsRef, conn, tran)
                End If
            End While
        End Using
    End Sub

    Friend Shared Sub DeleteRptIs9Provision(conn As IDbConnection, tran As IDbTransaction)
        Using cmd As IDbCommand = conn.CreateCommand
            cmd.ExecuteNonQuery(
                AppConst.DELETE_FROM_RPT_IS9_PROVISION,
                New IDbDataParameter() {},
                tran
            )
        End Using
        _logger.Debug($"Delete Table RPT_IS9_PROVISION is Done")
    End Sub

    Friend Shared Sub InsertRptIs9Provision(conn As IDbConnection, tran As IDbTransaction)
        Using cmd As IDbCommand = conn.CreateCommand
            cmd.ExecuteNonQuery(
                AppConst.INSERT_INTO_RPT_IS9_PROVISION_FROM_RPT_IS9_PROVISION_STG,
                New IDbDataParameter() {},
                tran
            )
        End Using
        _logger.Debug($"Insert Table RPT_IS9_PROVISION is Done")
    End Sub

    Friend Shared Sub UpdateRptIS9Provision(conn As IDbConnection, tran As IDbTransaction)
        DeleteRptIs9Provision(conn, tran)
        InsertRptIs9Provision(conn, tran)
        _logger.Debug($"Update RPT_IS9_PROVISION is Done")
    End Sub

    Friend Shared Sub UpdateSwapCurveStartDate(conn As IDbConnection, tran As IDbTransaction)
        Using cmd As IDbCommand = conn.CreateCommand
            cmd.ExecuteNonQuery(
                AppConst.UPDATE_RPT_SWAP_CURVE_SQL,
                New IDbDataParameter() {},
                tran
            )
        End Using

        Using cmd As IDbCommand = conn.CreateCommand
            cmd.ExecuteNonQuery(
                AppConst.UPDATE_RPT_SWAP_CURVE_CURRENCY_SQL,
                New IDbDataParameter() {},
                tran
            )
        End Using
        _logger.Debug($"Update Table RPT_SWAP_CURVE is Done")
    End Sub

    Friend Shared Sub InsertExcp(billsRef As String, msg As String, conn As IDbConnection, Optional tran As IDbTransaction = Nothing)
        If conn IsNot Nothing Then

            Using cmd As IDbCommand = conn.CreateCommand
                cmd.ExecuteNonQuery(
                    AppConst.INSERT_INTO_RPT_EXCPT,
                    New IDbDataParameter() {
                        cmd.CreateParameter("BILLS_REF", billsRef),
                        cmd.CreateParameter("ERROR_MESSAGE", msg)
                    },
                tran
                )
            End Using
        End If
    End Sub

    Friend Shared Sub CheckUploadCurveIsNotExist(billsRef As String, conn As IDbConnection, Optional tran As IDbTransaction = Nothing)
        Dim result As String = ""
        Using cmd As IDbCommand = conn.CreateCommand,
                    rs As IDataReader = cmd.ExecuteReader(
                        AppConst.SELECT_RPT_SWAP_CURVE_SQL,
                        cmd.CreateParameter()
                    )
            While rs.Read()
                Dim name As String = rs.GetString("CURVE_NAME")
                If Not String.IsNullOrEmpty(name) Then
                    If result <> "" Then
                        result = result & ","
                    End If
                    result = result & name
                End If
            End While
        End Using

        If result.Length >= 100 Then
            result = result.Substring(0, 98)
        End If
        If result.Length > 0 Then
            InsertExcp(billsRef, "KPS curve is not defined: " & result, conn, tran)
            Throw New Exception("KPS curve is not defined: " & result)
        End If
    End Sub
End Class