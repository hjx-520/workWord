Public Class CMHCHCCU
    Private Local_Operating_Date As String
    Private Local_Entity_Code As String
    Private Local_Branch_Code As String
    Private Local_Operating_Terminal As String
    Private Local_User_Code As String
    Private Local_Error_RC As String
    Private Local_Error_Application_RC As String
    Private Local_Error_Application_Reference As String
    Private Local_Error_Program As String
    Private Local_Error_Reference As String
    Private Local_Error_CICS As String
    Private Local_Error_EIBFN As String
    Private Local_Error_EIBRSRCE As String
    Private Local_Error_EIBRCODE As String
    Private Local_Error_EIBRESP1 As String
    Private Local_Error_EIBRESP2 As String
    Private Local_Error_Object As String
    Private Local_Error_SQLCODE As String
    Private Local_Error_SQLERRM As String
    Private Local_Error_DLI_Code As String
    Private Local_Error_DLI_Description As String
    Private Local_Error_Filler As String
    Private Local_Filler As String

    Private errMessage As String
    Private ErrorHandling As New Logger

    Public Property Operating_Date() As String
        Get
            Return myTrim(Local_Operating_Date)
        End Get
        Set(ByVal value As String)
            Local_Operating_Date = formatString(Value, 10)
        End Set
    End Property

    Public Property Entity_Code() As String
        Get
            Return myTrim(Local_Entity_Code)
        End Get
        Set(ByVal value As String)
            Local_Entity_Code = formatString(Value, 4)
        End Set
    End Property

    Public Property Branch_Code() As String
        Get
            Return myTrim(Local_Branch_Code)
        End Get
        Set(ByVal value As String)
            Local_Branch_Code = formatString(Value, 4)
        End Set
    End Property

    Public Property Operating_Terminal() As String
        Get
            Return myTrim(Local_Operating_Terminal)
        End Get
        Set(ByVal value As String)
            Local_Operating_Terminal = formatString(Value, 4)
        End Set
    End Property

    Public Property User_Code() As String
        Get
            Return myTrim(Local_User_Code)
        End Get
        Set(ByVal value As String)
            Local_User_Code = formatString(Value, 8)
        End Set
    End Property

    Public Property Error_RC() As String
        Get
            Return myTrim(Local_Error_RC)
        End Get
        Set(ByVal value As String)
            Local_Error_RC = formatString(Value, 2)
        End Set
    End Property

    Public Property Error_Application_RC() As String
        Get
            Return myTrim(Local_Error_Application_RC)
        End Get
        Set(ByVal value As String)
            Local_Error_Application_RC = formatString(Value, 2)
        End Set
    End Property

    Public Property Error_Application_Reference() As String
        Get
            Return myTrim(Local_Error_Application_Reference)
        End Get
        Set(ByVal value As String)
            Local_Error_Application_Reference = formatString(Value, 8)
        End Set
    End Property

    Public Property Error_Program() As String
        Get
            Return myTrim(Local_Error_Program)
        End Get
        Set(ByVal value As String)
            Local_Error_Program = formatString(Value, 8)
        End Set
    End Property

    Public Property Error_Reference() As String
        Get
            Return myTrim(Local_Error_Reference)
        End Get
        Set(ByVal value As String)
            Local_Error_Reference = formatString(Value, 100)
        End Set
    End Property

    Public Property Error_CICS() As String
        Get
            Return myTrim(Local_Error_CICS)
        End Get
        Set(ByVal value As String)
            Local_Error_CICS = formatString(Value, 8)
        End Set
    End Property

    Public Property Error_EIBFN() As String
        Get
            Return myTrim(Local_Error_EIBFN)
        End Get
        Set(ByVal value As String)
            Local_Error_EIBFN = formatString(Value, 2)
        End Set
    End Property

    Public Property Error_EIBRSRCE() As String
        Get
            Return myTrim(Local_Error_EIBRSRCE)
        End Get
        Set(ByVal value As String)
            Local_Error_EIBRSRCE = formatString(Value, 8)
        End Set
    End Property

    Public Property Error_EIBRCODE() As String
        Get
            Return myTrim(Local_Error_EIBRCODE)
        End Get
        Set(ByVal value As String)
            Local_Error_EIBRCODE = formatString(Value, 6)
        End Set
    End Property

    Public Property Error_EIBRESP1() As Double
        Get
            Return Val(Mid(Local_Error_EIBRESP1, 2))
        End Get
        Set(ByVal value As Double)
            Local_Error_EIBRESP1 = formatNumber(value, 9, 0, False, True)
        End Set
    End Property

    Public Property Error_EIBRESP2() As Double
        Get
            Return Val(Mid(Local_Error_EIBRESP2, 2))
        End Get
        Set(ByVal value As Double)
            Local_Error_EIBRESP2 = formatNumber(value, 9, 0, False, True)
        End Set
    End Property

    Public Property Error_Object() As String
        Get
            Return myTrim(Local_Error_Object)
        End Get
        Set(ByVal value As String)
            Local_Error_Object = formatString(Value, 8)
        End Set
    End Property

    Public Property Error_SQLCODE() As Double
        Get
            Return Val(Mid(Local_Error_SQLCODE, 2))
        End Get
        Set(ByVal value As Double)
            Local_Error_SQLCODE = formatNumber(value, 4, 0, False, True)
        End Set
    End Property

    Public Property Error_SQLERRM() As String
        Get
            Return myTrim(Local_Error_SQLERRM)
        End Get
        Set(ByVal value As String)
            Local_Error_SQLERRM = formatString(Value, 70)
        End Set
    End Property

    Public Property Error_DLI_Code() As String
        Get
            Return myTrim(Local_Error_DLI_Code)
        End Get
        Set(ByVal value As String)
            Local_Error_DLI_Code = formatString(Value, 2)
        End Set
    End Property

    Public Property Error_DLI_Description() As String
        Get
            Return myTrim(Local_Error_DLI_Description)
        End Get
        Set(ByVal value As String)
            Local_Error_DLI_Description = formatString(Value, 90)
        End Set
    End Property

    Public Property Error_Filler() As String
        Get
            Return myTrim(Local_Error_Filler)
        End Get
        Set(ByVal value As String)
            Local_Error_Filler = formatString(Value, 61)
        End Set
    End Property

    Public Property Filler() As String
        Get
            Return myTrim(Local_Filler)
        End Get
        Set(ByVal value As String)
            Local_Filler = formatString(Value, 270)
        End Set
    End Property

    Public Function formatCommArea() As String
        Dim strBuff As String
        Dim myLogger As New Logger
        strBuff = ""
        Try
            strBuff = strBuff + Local_Operating_Date
            strBuff = strBuff + Local_Entity_Code
            strBuff = strBuff + Local_Branch_Code
            strBuff = strBuff + Local_Operating_Terminal
            strBuff = strBuff + Local_User_Code

            strBuff = strBuff + Local_Error_RC
            strBuff = strBuff + Local_Error_Application_RC
            strBuff = strBuff + Local_Error_Application_Reference
            strBuff = strBuff + Local_Error_Program
            strBuff = strBuff + Local_Error_Reference
            strBuff = strBuff + Local_Error_CICS
            strBuff = strBuff + Local_Error_EIBFN
            strBuff = strBuff + Local_Error_EIBRSRCE

            strBuff = strBuff + Local_Error_EIBRCODE
            strBuff = strBuff + Local_Error_EIBRESP1
            strBuff = strBuff + Local_Error_EIBRESP2
            strBuff = strBuff + Local_Error_Object
            strBuff = strBuff + Local_Error_SQLCODE
            strBuff = strBuff + Local_Error_SQLERRM
            strBuff = strBuff + Local_Error_DLI_Code
            strBuff = strBuff + Local_Error_DLI_Description
            strBuff = strBuff + Local_Error_Filler

            strBuff = strBuff + Local_Filler
        Catch ex As Exception
            errMessage = "Err Code: " & CStr(Err.Number) & ", Err Desc: " & Err.Description
            'ErrorHandling.WriteLog("CMHCHCCU.formatCommArea", errMessage)
            myLogger.AppendToFile(GlobalVarable.ErrLogFile, myLogger.FormatString(True, GlobalVarable.AppName, "CMHCHCCU.formatCommArea", errMessage))
            'MsgBox(errMessage)
        End Try
        myLogger = Nothing
        Return strBuff
    End Function

    Public Sub returnCommArea(ByVal strBuff As String)

        Dim myLogger As New Logger

        Try
            Local_Error_RC = Mid(strBuff, 31, 2)
            Local_Error_Application_RC = Mid(strBuff, 33, 2)
            Local_Error_Application_Reference = Mid(strBuff, 35, 8)
            Local_Error_Program = Mid(strBuff, 43, 8)
            Local_Error_Reference = Mid(strBuff, 51, 100)
            Local_Error_CICS = Mid(strBuff, 151, 8)
            Local_Error_EIBFN = Mid(strBuff, 159, 2)
            Local_Error_EIBRSRCE = Mid(strBuff, 161, 8)
            Local_Error_EIBRCODE = Mid(strBuff, 169, 6)
            Local_Error_EIBRESP1 = Mid(strBuff, 175, 10)
            Local_Error_EIBRESP2 = Mid(strBuff, 185, 10)
            Local_Error_Object = Mid(strBuff, 195, 8)
            Local_Error_SQLCODE = Mid(strBuff, 203, 5)
            Local_Error_SQLERRM = Mid(strBuff, 208, 70)
            Local_Error_DLI_Code = Mid(strBuff, 278, 2)
            Local_Error_DLI_Description = Mid(strBuff, 280, 90)
            Local_Error_Filler = Mid(strBuff, 370, 61)
            Local_Filler = Mid(431, 270)
        Catch ex As Exception
            errMessage = "Err Code: " & CStr(Err.Number) & ", Err Desc: " & Err.Description
            'ErrorHandling.WriteLog("CMHCHCCU.returnCommArea", errMessage)
            myLogger.AppendToFile(GlobalVarable.ErrLogFile, myLogger.FormatString(True, GlobalVarable.AppName, "CMHCHCCU.returnCommArea", errMessage))
            'MsgBox(errMessage)
        End Try
        myLogger = Nothing

    End Sub

    Private Function formatString(ByVal strInStr As String, ByVal intLen As Integer) As String

        Dim myLogger As New Logger
        Try
            strInStr = Trim(strInStr)
            If strInStr.Length < intLen Then
                strInStr = strInStr + Space(intLen - strInStr.Length)
            End If
            If strInStr.Length > intLen Then
                strInStr = Left(strInStr, intLen)
            End If
        Catch ex As Exception
            errMessage = "Err Code: " & CStr(Err.Number) & ", Err Desc: " & Err.Description
            'ErrorHandling.WriteLog("CMHCHCCU.formatString", errMessage)
            myLogger.AppendToFile(GlobalVarable.ErrLogFile, myLogger.FormatString(True, GlobalVarable.AppName, "CMHCHCCU.formatString", errMessage))
            'MsgBox(errMessage)
        End Try
        myLogger = Nothing
        Return strInStr
    End Function

    Private Function formatNumber(ByVal dblInNum As Double, ByVal intInteger As Integer, _
                                ByVal intDec As Integer, ByVal blnShowDec As Boolean, ByVal blnShowPos As Boolean) As String
        Dim strFormat As String
        Dim i As Integer
        Dim intTotalLen As Integer
        Dim dblValue As Double
        Dim strValue As String
        Dim myLogger As New Logger

        strValue = ""
        Try
            strFormat = ""
            For i = 1 To intInteger Step 1
                strFormat = strFormat + "0"
            Next
            intTotalLen = intInteger
            If blnShowDec And intDec > 0 Then
                strFormat = strFormat + "."
                intTotalLen = intTotalLen + 1
            End If
            For i = 1 To intDec Step 1
                strFormat = strFormat + "0"
            Next
            intTotalLen = intTotalLen + intDec
            dblValue = Math.Abs(dblInNum)
            strValue = Format(dblValue, strFormat)
            If Len(strValue) > intTotalLen Then
                strValue = Right(strValue, intTotalLen)
            End If

            If blnShowPos Then
                If dblInNum < 0 Then
                    strValue = "-" + strValue
                Else
                    strValue = "+" + strValue
                End If
            End If
        Catch ex As Exception
            errMessage = "Err Code: " & CStr(Err.Number) & ", Err Desc: " & Err.Description
            'ErrorHandling.WriteLog("CMHCHCCU.formatNumber", errMessage)
            myLogger.AppendToFile(GlobalVarable.ErrLogFile, myLogger.FormatString(True, GlobalVarable.AppName, "CMHCHCCU.formatNumber", errMessage))
            'MsgBox(errMessage)
        End Try
        myLogger = Nothing
        Return strValue
    End Function

    Private Function myTrim(ByVal strValue As String)
        strValue = Replace(strValue, Chr(0), Space(1))
        strValue = Trim(strValue)
        Return strValue
    End Function
End Class
