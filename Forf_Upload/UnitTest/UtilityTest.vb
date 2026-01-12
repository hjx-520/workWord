Imports System.IO
Imports Forf_Upload
Imports BEA.TFS.CommonUtility
' TODO: [REMOVE COMMENT AFTER REVIEW] Renamed file and class to UtilityTest for consistency with other projects.
<TestClass()> Public Class UtilityTest

    <TestMethod()> Public Sub Test_CopyCsvForLoadToDB()
        Directory.CreateDirectory("folder1")
        Directory.CreateDirectory("folder2")

        Dim runCtlDirPath As String = "folder1\csvFeeIncomePath.csv"
        Dim csvFeeIncomePath As String = "folder2\csvFeeIncomePath.csv"
        Dim expectedFileToBeGenerated As String = "folder1\csvFeeIncomePath.csv"

        File.WriteAllText(csvFeeIncomePath, "TestCopyCsvForLoadToDB")

        Try
            File.Delete(expectedFileToBeGenerated)
        Catch ex As Exception
        End Try

        File.Copy(csvFeeIncomePath, runCtlDirPath)

        Dim actual As String = File.ReadAllText(expectedFileToBeGenerated)
        Dim expected As String = "TestCopyCsvForLoadToDB"
        Assert.AreEqual(expected, actual)
    End Sub

    ' TODO: There is no testing file valid after CR, please test at SIT
    <TestMethod()> Public Sub Test_ConvertFile()
        'Directory.CreateDirectory("folder")

        'Dim csvBCRecoursePath As String = "folder\csvBCRecoursePath2.txt"
        'Dim txtBCRecoursePath As String = "folder\txtBCRecoursePath1.txt"
        'Dim JpyBCRecoursePath As String = "folder\txtBCRecoursePath3.txt"

        'File.WriteAllLines(txtBCRecoursePath, {
        '    "0020181231KPSTFS",
        '    "05CAD-3M-BS-D     CAD2017123120180101 0.99997247",
        '    "99000000000001"
        '})

        'File.WriteAllLines(JpyBCRecoursePath, {
        '    "0020181231TCLTFS",
        '    "05JPY-3M          JPY201902082019051301.00026349",
        '    "99000000000001"
        '})

        'Try
        '    File.Delete(csvBCRecoursePath)
        'Catch ex As Exception
        'End Try

        'Utility.GenRptSwapCurveCsv("EmptyBillsRef", New Date(2018, 12, 31), txtBCRecoursePath, JpyBCRecoursePath, csvBCRecoursePath, Nothing, Nothing)

        'Dim actual As String = File.ReadAllText(csvBCRecoursePath)
        'Dim expected As String = """CAD-3M-BS-D     "",""CAD"",""20171231"",""20180101"","" 0.99997247                                    """ _
        '    & vbNewLine & """JPY-3M          "",""JPY"",""20190208"",""20190513"",""01.00026349""" & vbNewLine
        'Assert.AreEqual(expected, actual)
    End Sub
    <TestMethod()> Public Sub Test_ValidateForfaiting()
        Dim connStub As DbConnectionStub = ConnectionStubBuilder.NewInstance _
            .AddQueryRtnVal() _
            .AddQueryRtnVal() _
                .WithFields(New Dictionary(Of String, Type) From {
                    {"BILLS_REF", Type.GetType("System.String")},
                    {"MODEL_TYPE", Type.GetType("System.String")}
                }) _
                .AddRow(New List(Of Object) From {
                    "bills_ref",
                    "model_type"
                }) _
            .AddQueryRtnVal() _
            .AddQueryRtnVal() _
            .AddQueryRtnVal() _
            .AddQueryRtnVal() _
            .AddQueryRtnVal() _
            .Build()
        Dim tranStub As DbTransactionStub = connStub.BeginTransaction()
        Dim sqlExpectation As CalledSqlExpectation = New CalledSqlExpectation(connStub)
        Dim Forf_Upload_Util As Utility = New Utility
        Dim paramExpectation As DbParamExpectation = New DbParamExpectation(connStub)

        Dim bill_ref As String = "bills_ref"

        Utility.ValidateForfaiting(connStub, tranStub)

        sqlExpectation.AssertEqualsTo(New List(Of String) From {
            AppConst.DELETE_RPT_EXCPT_SQL,
            AppConst.SELECT_FROM_RPT_FORF_UPLOAD,
            AppConst.SELECT_BILLS_REF_FROM_RPT_IFP_BILLS_DETAIL_MONTHEND,
            AppConst.INSERT_INTO_RPT_EXCPT,
            AppConst.UPDATE_RPT_FORF_UPLOAD_VALID_IND,
             AppConst.INSERT_INTO_RPT_EXCPT,
            AppConst.UPDATE_RPT_FORF_UPLOAD_VALID_IND
        })

        paramExpectation.ClearExpectedParam()
        paramExpectation.AddParam("BILLS_REF", DbType.String, bill_ref)
        paramExpectation.AssertEqualsToCmd(2)


        paramExpectation.ClearExpectedParam()
        paramExpectation.AddParam("BILLS_REF", DbType.String, bill_ref)
        paramExpectation.AddParam("ERROR_MESSAGE", DbType.String, "Bills reference was not found in RPT_IFP_BILLS_DETAIL_MONTHEND")

        paramExpectation.AssertEqualsToCmd(3)

        paramExpectation.ClearExpectedParam()
        paramExpectation.AddParam("BILLS_REF", DbType.String, bill_ref)
        paramExpectation.AssertEqualsToCmd(4)

        paramExpectation.ClearExpectedParam()
        paramExpectation.AddParam("BILLS_REF", DbType.String, bill_ref)
        paramExpectation.AddParam("ERROR_MESSAGE", DbType.String, "MODEL TYPE model_type is invalid")
        paramExpectation.AssertEqualsToCmd(5)

        paramExpectation.ClearExpectedParam()
        paramExpectation.AddParam("BILLS_REF", DbType.String, bill_ref)
        paramExpectation.AssertEqualsToCmd(6)
    End Sub


    <TestMethod()> Public Sub Test_ValidateRepriceForOthers_modeltype_2()
        Dim connStub As DbConnectionStub = ConnectionStubBuilder.NewInstance _
                        .AddQueryRtnVal() _
                .WithFields(New Dictionary(Of String, Type) From {
                    {"BILLS_REF", Type.GetType("System.String")},
                    {"MODEL_TYPE", Type.GetType("System.String")}
                }) _
                .AddRow(New List(Of Object) From {
                    "bills_ref",
                    "2"
                }) _
            .AddQueryRtnVal() _
            .AddQueryRtnVal() _
            .AddQueryRtnVal() _
            .AddQueryRtnVal() _
            .AddQueryRtnVal() _
             .AddQueryRtnVal() _
             .AddQueryRtnVal() _
             .AddQueryRtnVal() _
            .Build()
        Dim tranStub As DbTransactionStub = connStub.BeginTransaction()
        Dim sqlExpectation As CalledSqlExpectation = New CalledSqlExpectation(connStub)
        Dim Forf_Upload_Util As Utility = New Utility
        Dim paramExpectation As DbParamExpectation = New DbParamExpectation(connStub)

        Dim bill_ref As String = "bills_ref"

        Utility.ValidateRepriceForOthers(connStub, tranStub)

        sqlExpectation.AssertEqualsTo(New List(Of String) From {
            AppConst.SELECT_FROM_RPT_FORF_UPLOAD_BY_MODEL_TYPE,
            AppConst.SELECT_IS_EXIST_IN_REPRICE_SQL,
            AppConst.INSERT_INTO_RPT_EXCPT,
            AppConst.UPDATE_RPT_FORF_UPLOAD_VALID_IND,
            AppConst.SELECT_END_DATE_FROM_RPT_REPRICE_UPLOAD,
            AppConst.SELECT_DUE_DATE_FROM_RPT_IFP_BILLS_DETAIL_MONTHEND,
            AppConst.INSERT_INTO_RPT_EXCPT,
            AppConst.UPDATE_RPT_FORF_UPLOAD_VALID_IND,
            AppConst.SELECT_BILLS_REF_FROM_RPT_IFP_BILLS_DETAIL_MONTHEND})

        paramExpectation.ClearExpectedParam()
        paramExpectation.AddParam("BILLS_REF", DbType.String, bill_ref)
        paramExpectation.AssertEqualsToCmd(1)

        paramExpectation.ClearExpectedParam()
        paramExpectation.AddParam("BILLS_REF", DbType.String, bill_ref)
        paramExpectation.AddParam("ERROR_MESSAGE", DbType.String, "Bills reference bills_ref has no reprice data")

        paramExpectation.AssertEqualsToCmd(2)

        paramExpectation.ClearExpectedParam()
        paramExpectation.AddParam("BILLS_REF", DbType.String, bill_ref)
        paramExpectation.AssertEqualsToCmd(3)


        paramExpectation.ClearExpectedParam()
        paramExpectation.AddParam("BILLS_REF", DbType.String, bill_ref)
        paramExpectation.AssertEqualsToCmd(4)

        paramExpectation.ClearExpectedParam()
        paramExpectation.AddParam("BILLS_REF", DbType.String, bill_ref)
        paramExpectation.AssertEqualsToCmd(5)

        paramExpectation.ClearExpectedParam()
        paramExpectation.AddParam("BILLS_REF", DbType.String, bill_ref)
        paramExpectation.AddParam("ERROR_MESSAGE", DbType.String, "Bills reference bills_ref Reprice period end date does not match the due date of the loan")

        paramExpectation.AssertEqualsToCmd(6)

        paramExpectation.ClearExpectedParam()
        paramExpectation.AddParam("BILLS_REF", DbType.String, bill_ref)
        paramExpectation.AssertEqualsToCmd(7)


        paramExpectation.ClearExpectedParam()
        paramExpectation.AddParam("BILLS_REF", DbType.String, bill_ref)
        paramExpectation.AssertEqualsToCmd(8)



    End Sub

    <TestMethod()> Public Sub Test_ValidateRepriceForOthers_ModelType3()
        Dim connStub As DbConnectionStub = ConnectionStubBuilder.NewInstance _
                        .AddQueryRtnVal() _
                .WithFields(New Dictionary(Of String, Type) From {
                    {"BILLS_REF", Type.GetType("System.String")},
                    {"MODEL_TYPE", Type.GetType("System.String")}
                }) _
                .AddRow(New List(Of Object) From {
                    "bills_ref",
                    "3"
                }) _
            .AddQueryRtnVal() _
            .AddQueryRtnVal() _
            .Build()
        Dim tranStub As DbTransactionStub = connStub.BeginTransaction()
        Dim sqlExpectation As CalledSqlExpectation = New CalledSqlExpectation(connStub)
        Dim Forf_Upload_Util As Utility = New Utility
        Dim paramExpectation As DbParamExpectation = New DbParamExpectation(connStub)

        Dim bill_ref As String = "bills_ref"

        Utility.ValidateRepriceForOthers(connStub, tranStub)

        sqlExpectation.AssertEqualsTo(New List(Of String) From {
            AppConst.SELECT_FROM_RPT_FORF_UPLOAD_BY_MODEL_TYPE,
            AppConst.SELECT_IS_EXIST_IN_REPRICE_SQL,
            AppConst.SELECT_BILLS_REF_FROM_RPT_IFP_BILLS_DETAIL_MONTHEND})

        paramExpectation.ClearExpectedParam()
        paramExpectation.AddParam("BILLS_REF", DbType.String, bill_ref)
        paramExpectation.AssertEqualsToCmd(1)


        paramExpectation.ClearExpectedParam()
        paramExpectation.AddParam("BILLS_REF", DbType.String, bill_ref)
        paramExpectation.AssertEqualsToCmd(2)


    End Sub

    <TestMethod()> Public Sub Test_ValidateReprice()
        Dim connStub As DbConnectionStub = ConnectionStubBuilder.NewInstance _
            .AddQueryRtnVal() _
                .WithFields(New Dictionary(Of String, Type) From {
                    {"BILLS_REF", Type.GetType("System.String")},
                    {"START_DATE", Type.GetType("System.Date")},
                    {"END_DATE", Type.GetType("System.Date")}
                }) _
                .AddRow(New List(Of Object) From {
                    "bills_ref",
                    Date.Parse("2018-01-01"),
                    Date.Parse("2018-01-01")}) _
            .AddQueryRtnVal() _
            .AddQueryRtnVal() _
            .AddQueryRtnVal() _
            .AddQueryRtnVal() _
             .AddQueryRtnVal() _
            .Build()
        Dim tranStub As DbTransactionStub = connStub.BeginTransaction()
        Dim sqlExpectation As CalledSqlExpectation = New CalledSqlExpectation(connStub)
        Dim Forf_Upload_Util As Utility = New Utility
        Dim paramExpectation As DbParamExpectation = New DbParamExpectation(connStub)

        Dim bill_ref As String = "bills_ref"

        Utility.ValidateReprice(connStub, tranStub)

        sqlExpectation.AssertEqualsTo(New List(Of String) From {
            AppConst.SELECT_FROM_RPT_REPRICE_UPLOAD,
            AppConst.SELECT_BILLS_REF_FROM_RPT_IFP_BILLS_DETAIL_MONTHEND,
            AppConst.INSERT_INTO_RPT_EXCPT,
            AppConst.UPDATE_RPT_FORF_UPLOAD_VALID_IND,
             AppConst.INSERT_INTO_RPT_EXCPT,
            AppConst.UPDATE_RPT_FORF_UPLOAD_VALID_IND
        })

        paramExpectation.ClearExpectedParam()
        paramExpectation.AddParam("BILLS_REF", DbType.String, bill_ref)
        paramExpectation.AssertEqualsToCmd(1)

        paramExpectation.ClearExpectedParam()
        paramExpectation.AddParam("BILLS_REF", DbType.String, bill_ref)
        paramExpectation.AddParam("ERROR_MESSAGE", DbType.String, "Bills reference was not found in RPT_FORF_UPLOAD or RPT_IFP_BILLS_DETAIL_MONTHEND")

        paramExpectation.AssertEqualsToCmd(2)

        paramExpectation.ClearExpectedParam()
        paramExpectation.AddParam("BILLS_REF", DbType.String, bill_ref)
        paramExpectation.AssertEqualsToCmd(3)

        paramExpectation.ClearExpectedParam()
        paramExpectation.AddParam("BILLS_REF", DbType.String, bill_ref)
        paramExpectation.AddParam("ERROR_MESSAGE", DbType.String, "Reprice Start Date >= End Date")
        paramExpectation.AssertEqualsToCmd(4)

        paramExpectation.ClearExpectedParam()
        paramExpectation.AddParam("BILLS_REF", DbType.String, bill_ref)
        paramExpectation.AssertEqualsToCmd(5)
    End Sub
    <TestMethod()> Public Sub Test_ValidateRepriceForPeriod()
        Dim connStub As DbConnectionStub = ConnectionStubBuilder.NewInstance _
            .AddQueryRtnVal() _
                .WithFields(New Dictionary(Of String, Type) From {
                    {"BILLS_REF", Type.GetType("System.String")}
                }) _
                .AddRow(New List(Of Object) From {
                    "bills_ref"
                }) _
            .AddQueryRtnVal() _
                .WithFields(New Dictionary(Of String, Type) From {
                    {"BILLS_REF", Type.GetType("System.String")}
                }) _
                .AddRow(New List(Of Object) From {
                    "bills_ref"
                }) _
            .AddQueryRtnVal() _
            .AddQueryRtnVal() _
            .Build()
        Dim tranStub As DbTransactionStub = connStub.BeginTransaction()
        Dim sqlExpectation As CalledSqlExpectation = New CalledSqlExpectation(connStub)
        Dim Forf_Upload_Util As Utility = New Utility
        Dim paramExpectation As DbParamExpectation = New DbParamExpectation(connStub)

        Dim bill_ref As String = "bills_ref"

        Utility.ValidateRepriceForPeriod(connStub, tranStub)

        sqlExpectation.AssertEqualsTo(New List(Of String) From {
            AppConst.SELECT_FROM_RPT_REPRICE_UPLOAD_GROUP_BY_BILLS_REF,
            AppConst.SELECT_IS_CONTINUE_PERIOD_SQL,
            AppConst.INSERT_INTO_RPT_EXCPT,
            AppConst.UPDATE_RPT_FORF_UPLOAD_VALID_IND})

        paramExpectation.ClearExpectedParam()
        paramExpectation.AddParam("BILLS_REF", DbType.String, bill_ref)
        paramExpectation.AssertEqualsToCmd(1)

        paramExpectation.ClearExpectedParam()
        paramExpectation.AddParam("BILLS_REF", DbType.String, bill_ref)
        paramExpectation.AddParam("ERROR_MESSAGE", DbType.String, "Reprice period end date is not equal to start date of next reprice period")
        paramExpectation.AssertEqualsToCmd(2)

        paramExpectation.ClearExpectedParam()
        paramExpectation.AddParam("BILLS_REF", DbType.String, bill_ref)
        paramExpectation.AssertEqualsToCmd(3)


    End Sub

    <TestMethod()> Public Sub Test_ValidateFullSet()
        Dim connStub As DbConnectionStub = ConnectionStubBuilder.NewInstance _
            .AddQueryRtnVal() _
                .WithFields(New Dictionary(Of String, Type) From {
                    {"BILLS_REF", Type.GetType("System.String")}
                }) _
                .AddRow(New List(Of Object) From {
                    "bills_ref"
                }) _
             .AddQueryRtnVal() _
            .Build()
        Dim tranStub As DbTransactionStub = connStub.BeginTransaction()
        Dim sqlExpectation As CalledSqlExpectation = New CalledSqlExpectation(connStub)
        Dim Forf_Upload_Util As Utility = New Utility
        Dim paramExpectation As DbParamExpectation = New DbParamExpectation(connStub)

        Dim bill_ref As String = "bills_ref"

        Utility.ValidateFullSet(connStub, tranStub)

        sqlExpectation.AssertEqualsTo(New List(Of String) From {
            AppConst.SELECT_FROM_RPT_FORF_REFN_MONTHEND_RPT_BILLS_DETAIL_MONTHEND_RPT_FORF_UPLOAD,
            AppConst.INSERT_INTO_RPT_EXCPT})


        paramExpectation.ClearExpectedParam()
        paramExpectation.AddParam("BILLS_REF", DbType.String, bill_ref)
        paramExpectation.AddParam("ERROR_MESSAGE", DbType.String, "Bills reference was not found in upload file")
        paramExpectation.AssertEqualsToCmd(1)



    End Sub

    <TestMethod()> Public Sub Test_ValidateCurve()
        Dim connStub As DbConnectionStub = ConnectionStubBuilder.NewInstance _
            .AddQueryRtnVal() _
                .WithFields(New Dictionary(Of String, Type) From {
                    {"BILLS_REF", Type.GetType("System.String")},
                    {"TENOR", Type.GetType("System.String")},
                     {"MODEL_TYPE", Type.GetType("System.String")},
                      {"BILLS_CCY", Type.GetType("System.String")}}) _
                .AddRow(New List(Of Object) From {
                    "bills_ref",
                    "123",
                    "2",
                    "A"}) _
             .AddQueryRtnVal() _
             .AddQueryRtnVal() _
             .AddQueryRtnVal() _
             .AddQueryRtnVal() _
             .AddQueryRtnVal() _
             .AddQueryRtnVal() _
            .Build()
        Dim tranStub As DbTransactionStub = connStub.BeginTransaction()
        Dim sqlExpectation As CalledSqlExpectation = New CalledSqlExpectation(connStub)
        Dim Forf_Upload_Util As Utility = New Utility
        Dim paramExpectation As DbParamExpectation = New DbParamExpectation(connStub)

        Dim bill_ref As String = "bills_ref"
        Dim BILLS_CCY = "A"
        Dim TENOR = "123"
        Utility.ValidateCurve(connStub, tranStub)

        sqlExpectation.AssertEqualsTo(New List(Of String) From {
            AppConst.SELECT_FROM_RPT_FORF_UPLOAD_RPT_IFP_BILLS_DETAIL_MONTHEND,
            AppConst.SELECT_FROM_RPT_CURVE_USAGE,
            AppConst.INSERT_INTO_RPT_EXCPT,
            AppConst.UPDATE_RPT_FORF_UPLOAD_VALID_IND,
            AppConst.SELECT_FROM_RPT_CURVE_USAGE,
            AppConst.INSERT_INTO_RPT_EXCPT,
            AppConst.UPDATE_RPT_FORF_UPLOAD_VALID_IND})

        paramExpectation.ClearExpectedParam()
        paramExpectation.AddParam("currencyCode", DbType.String, BILLS_CCY)
        paramExpectation.AddParam("usage", DbType.AnsiStringFixedLength, "F"c)
        paramExpectation.AddParam("skipTenorCodeCondition", DbType.Int32, 0)
        paramExpectation.AddParam("tenorCode", DbType.String, TENOR)
        paramExpectation.AssertEqualsToCmd(1)



        paramExpectation.ClearExpectedParam()
        paramExpectation.AddParam("BILLS_REF", DbType.String, bill_ref)
        paramExpectation.AddParam("ERROR_MESSAGE", DbType.String, "Curve is not found")
        paramExpectation.AssertEqualsToCmd(2)

        paramExpectation.ClearExpectedParam()
        paramExpectation.AddParam("BILLS_REF", DbType.String, bill_ref)
        paramExpectation.AssertEqualsToCmd(3)


        paramExpectation.ClearExpectedParam()
        paramExpectation.AddParam("currencyCode", DbType.String, BILLS_CCY)
        paramExpectation.AddParam("usage", DbType.AnsiStringFixedLength, "D"c)
        paramExpectation.AddParam("skipTenorCodeCondition", DbType.Int32, 1)
        paramExpectation.AddParam("tenorCode", DbType.String, TENOR)
        paramExpectation.AssertEqualsToCmd(4)


        paramExpectation.ClearExpectedParam()
        paramExpectation.AddParam("BILLS_REF", DbType.String, bill_ref)
        paramExpectation.AddParam("ERROR_MESSAGE", DbType.String, "Curve is not found")
        paramExpectation.AssertEqualsToCmd(5)

        paramExpectation.ClearExpectedParam()
        paramExpectation.AddParam("BILLS_REF", DbType.String, bill_ref)
        paramExpectation.AssertEqualsToCmd(6)

    End Sub


    <TestMethod()> Public Sub Test_InsertForfaitingMaster()
        Dim connStub As DbConnectionStub = ConnectionStubBuilder.NewInstance _
            .AddQueryRtnVal() _
            .AddQueryRtnVal() _
                .WithFields(New Dictionary(Of String, Type) From {
                    {"BILLS_REF", Type.GetType("System.String")},
                    {"TENOR", Type.GetType("System.String")},
                    {"MODEL_TYPE", Type.GetType("System.String")},
                    {"BILLS_CCY", Type.GetType("System.Decimal")},
                    {"OS_BAL_FCY", Type.GetType("System.Decimal")},
                    {"CREATE_DATE", Type.GetType("System.Date")},
                    {"DUE_DATE", Type.GetType("System.Date")},
                    {"INT_RATE", Type.GetType("System.Decimal")}
                }) _
                .AddRow(New List(Of Object) From {
                    "bills_ref",
                    "tenor",
                    "model_type",
                    "currency",
                     Decimal.Parse("2"),
                    DateTime.Parse("2018-01-01"),
                    DateTime.Parse("2018-01-01"),
                    Decimal.Parse("3")
                }) _
             .AddQueryRtnVal() _
            .Build()
        Dim tranStub As DbTransactionStub = connStub.BeginTransaction()
        Dim sqlExpectation As CalledSqlExpectation = New CalledSqlExpectation(connStub)
        Dim Forf_Upload_Util As Utility = New Utility
        Dim paramExpectation As DbParamExpectation = New DbParamExpectation(connStub)

        Dim bill_ref As String = "bills_ref"
        Dim tenor As String = "tenor"
        Dim Model_type As String = "model_type"
        Dim Bills_CCY As String = "currency"
        Dim OS_BAL_FCY As Decimal = 2
        Dim CREATE_DATE As DateTime = Date.Parse("2018-01-01")
        Dim DUE_DATE As DateTime = Date.Parse("2018-01-01")
        Dim INT_RATE As Decimal = 3

        Utility.InsertForfaitingMaster(connStub, tranStub)

        sqlExpectation.AssertEqualsTo(New List(Of String) From {
            AppConst.DELETE_RPT_FORF_MASTER_SQL,
            AppConst.SELECT_DATAPREPARE_RPT_FORF_MASTER_SQL,
            AppConst.INSERT_RPT_FORF_MASTER_SQL})


        paramExpectation.ClearExpectedParam()
        paramExpectation.AddParam("BILLS_REF", DbType.String, bill_ref)
        paramExpectation.AddParam("TENOR", DbType.String, tenor)
        paramExpectation.AddParam("MODEL_TYPE", DbType.String, Model_type)
        paramExpectation.AddParam("CURRENCY", DbType.String, Bills_CCY)
        paramExpectation.AddParam("LOAN_PRINCIPAL", DbType.Decimal, OS_BAL_FCY)
        paramExpectation.AddParam("LOAN_START_DATE", DbType.DateTime, CREATE_DATE)
        paramExpectation.AddParam("LOAN_DUE_DATE", DbType.DateTime, DUE_DATE)
        paramExpectation.AddParam("INTEREST_RATE", DbType.Decimal, INT_RATE)

        paramExpectation.AssertEqualsToCmd(2)



    End Sub

    <TestMethod()> Public Sub Test_ResolvedFileName()
        Dim SWAPCURVEFTPFILE As String = "D:\cygwinhome\BIL\0101411\Enq\RCV\KPSTOTFS_SWAPCURVE_D<yyMMdd>.DAT"
        Dim SWAPCURVEJPYFTPFILE As String = "D:\cygwinhome\BIL\0101411\Enq\RCV\TCLTOTFS_YIELDCURVE_D<yyMMdd>.DAT"
        Dim PROVISIONFTPFILE As String = "D:\cygwinhome\BIL\0101411\Enq\RCV\IS9TOTFS_PROVISION_D<yyMMdd>.DAT"

        Dim swapCurveDate As Date = New Date(2019, 4, 30)
        Dim actualSWAPCURVEFTPFILE As String = FileUtils.ResolveFileName(SWAPCURVEFTPFILE, swapCurveDate)
        Dim actualCURVEJPYFTPFILE As String = FileUtils.ResolveFileName(SWAPCURVEJPYFTPFILE, swapCurveDate)
        Dim actualPROVISIONFTPFILE As String = FileUtils.ResolveFileName(PROVISIONFTPFILE, swapCurveDate)

        Dim expectedSWAPCURVEFTPFILE As String = "D:\cygwinhome\BIL\0101411\Enq\RCV\KPSTOTFS_SWAPCURVE_D190430.DAT"
        Dim expectedSWAPCURVEJPYFTPFILE As String = "D:\cygwinhome\BIL\0101411\Enq\RCV\TCLTOTFS_YIELDCURVE_D190430.DAT"
        Dim expectedPROVISIONFTPFILE As String = "D:\cygwinhome\BIL\0101411\Enq\RCV\IS9TOTFS_PROVISION_D190430.DAT"

        Assert.AreEqual(expectedSWAPCURVEFTPFILE, actualSWAPCURVEFTPFILE)
        Assert.AreEqual(expectedSWAPCURVEJPYFTPFILE, actualCURVEJPYFTPFILE)
        Assert.AreEqual(expectedPROVISIONFTPFILE, actualPROVISIONFTPFILE)

    End Sub

End Class