Public Class CMHIQHDU
    Private Local_CMHIQHDE_MSG_INFO As String
    Private Local_CMHIQHDE_REQ_APPL_ID As String
    Private Local_CMHIQHDE_REQ_MOD_ID As String
    Private Local_CMHIQHDE_REQ_ID As String
    Private Local_CMHIQHDE_INTF_REQ_NO As String
    Private Local_CMHIQHDE_SRV_ID As String
    Private Local_CMHIQHDE_TX_REF As String
    Private Local_CMHIQHDE_RESEND_IND As String
    Private Local_CMHIQHDE_RESERVE_1 As String
    Private Local_CMHIQHDE_RESERVE_2 As String
    Private Local_CMHIQHDE_RESERVE_3 As String

    Private Local_CMHOPHDE_RT_CODE As String
    Private Local_CMHOPHDE_RT_MSG_ID As String
    Private Local_CMHOPHDE_RT_MSG_TEXT As String
    Private Local_CMHOPHDE_RESERVE As String
    Private Local_CMHOPHDE_MSG_LEN As String

    Private errMessage As String
    Private ErrorHandling As New Logger

    Public Property CMHIQHDE_MSG_INFO() As String
        Get
            Return myTrim(Local_CMHIQHDE_MSG_INFO)
        End Get
        Set(ByVal value As String)
            Local_CMHIQHDE_MSG_INFO = formatString(value, 8)
        End Set
    End Property

    Public Property CMHIQHDE_REQ_APPL_ID() As String
        Get
            Return myTrim(Local_CMHIQHDE_REQ_APPL_ID)
        End Get
        Set(ByVal value As String)
            Local_CMHIQHDE_REQ_APPL_ID = formatString(Value, 3)
        End Set
    End Property

    Public Property CMHIQHDE_REQ_MOD_ID() As String
        Get
            Return myTrim(Local_CMHIQHDE_REQ_MOD_ID)
        End Get
        Set(ByVal value As String)
            Local_CMHIQHDE_REQ_MOD_ID = formatString(Value, 3)
        End Set
    End Property

    Public Property CMHIQHDE_REQ_ID() As String
        Get
            Return myTrim(Local_CMHIQHDE_REQ_ID)
        End Get
        Set(ByVal value As String)
            Local_CMHIQHDE_REQ_ID = formatString(Value, 3)
        End Set
    End Property

    Public Property CMHIQHDE_INTF_REQ_NO() As String
        Get
            Return myTrim(Local_CMHIQHDE_INTF_REQ_NO)
        End Get
        Set(ByVal value As String)
            Local_CMHIQHDE_INTF_REQ_NO = formatString(Value, 7)
        End Set
    End Property

    Public Property CMHIQHDE_SRV_ID() As String
        Get
            Return myTrim(Local_CMHIQHDE_SRV_ID)
        End Get
        Set(ByVal value As String)
            Local_CMHIQHDE_SRV_ID = formatString(Value, 20)
        End Set
    End Property

    Public Property CMHIQHDE_TX_REF() As String
        Get
            Return myTrim(Local_CMHIQHDE_TX_REF)
        End Get
        Set(ByVal value As String)
            Local_CMHIQHDE_TX_REF = formatString(Value, 50)
        End Set
    End Property

    Public Property CMHIQHDE_RESEND_IND() As String
        Get
            Return myTrim(Local_CMHIQHDE_RESEND_IND)
        End Get
        Set(ByVal value As String)
            Local_CMHIQHDE_RESEND_IND = formatString(Value, 1)
        End Set
    End Property

    Public Property CMHIQHDE_RESERVE_1() As String
        Get
            Return myTrim(Local_CMHIQHDE_RESERVE_1)
        End Get
        Set(ByVal value As String)
            Local_CMHIQHDE_RESERVE_1 = formatString(Value, 700)
        End Set
    End Property

    Public Property CMHIQHDE_RESERVE_2() As String
        Get
            Return myTrim(Local_CMHIQHDE_RESERVE_2)
        End Get
        Set(ByVal value As String)
            Local_CMHIQHDE_RESERVE_2 = formatString(Value, 500)
        End Set
    End Property

    Public Property CMHIQHDE_RESERVE_3() As String
        Get
            Return myTrim(Local_CMHIQHDE_RESERVE_3)
        End Get
        Set(ByVal value As String)
            Local_CMHIQHDE_RESERVE_3 = formatString(Value, 205)
        End Set
    End Property

    Public Property CMHOPHDE_RT_CODE() As String
        Get
            Return myTrim(Local_CMHOPHDE_RT_CODE)
        End Get
        Set(ByVal value As String)
            Local_CMHOPHDE_RT_CODE = formatString(Value, 2)
        End Set
    End Property

    Public Property CMHOPHDE_RT_MSG_ID() As String
        Get
            Return myTrim(Local_CMHOPHDE_RT_MSG_ID)
        End Get
        Set(ByVal value As String)
            Local_CMHOPHDE_RT_MSG_ID = formatString(Value, 8)
        End Set
    End Property

    Public Property CMHOPHDE_RT_MSG_TEXT() As String
        Get
            Return myTrim(Local_CMHOPHDE_RT_MSG_TEXT)
        End Get
        Set(ByVal value As String)
            Local_CMHOPHDE_RT_MSG_TEXT = formatString(Value, 250)
        End Set
    End Property

    Public Property CMHOPHDE_RESERVE() As String
        Get
            Return myTrim(Local_CMHOPHDE_RESERVE)
        End Get
        Set(ByVal value As String)
            Local_CMHOPHDE_RESERVE = formatString(Value, 34)
        End Set
    End Property

    Public Property CMHOPHDE_MSG_LEN() As Double
        Get
            Return Val(Local_CMHOPHDE_MSG_LEN)
        End Get
        Set(ByVal value As Double)
            Local_CMHOPHDE_MSG_LEN = formatNumber(value, 6, 0, False, False)
        End Set
    End Property

    Public Function formatCommArea() As String
        Dim strBuff As String

        strbuff = ""
        strbuff = strbuff + Local_CMHIQHDE_MSG_INFO
        strbuff = strbuff + Local_CMHIQHDE_REQ_APPL_ID
        strbuff = strbuff + Local_CMHIQHDE_REQ_MOD_ID
        strbuff = strbuff + Local_CMHIQHDE_REQ_ID
        strbuff = strbuff + Local_CMHIQHDE_INTF_REQ_NO
        strbuff = strbuff + Local_CMHIQHDE_SRV_ID
        strbuff = strbuff + Local_CMHIQHDE_TX_REF
        strbuff = strbuff + Local_CMHIQHDE_RESEND_IND
        strbuff = strbuff + Local_CMHIQHDE_RESERVE_1
        strbuff = strbuff + Local_CMHIQHDE_RESERVE_2
        strbuff = strbuff + Local_CMHIQHDE_RESERVE_3

        Return strbuff
    End Function

    Public Function returnCommArea(ByVal strBuff As String) As Boolean
        Dim strErrorMessage As String
        Dim myLogger = New Logger

        Try
            strErrorMessage = ""
            Local_CMHOPHDE_RT_CODE = Mid(strBuff, 1501, 2)
            Local_CMHOPHDE_RT_MSG_ID = Mid(strBuff, 1503, 8)
            Local_CMHOPHDE_RT_MSG_TEXT = Mid(strBuff, 1511, 250)
            Local_CMHOPHDE_RESERVE = Mid(strBuff, 1761, 34)
            Local_CMHOPHDE_MSG_LEN = Mid(strBuff, 1795, 6)

            If Local_CMHOPHDE_RT_CODE <> "00" And
                Trim(Local_CMHOPHDE_RT_MSG_ID) <> "CMH2100E" And
                Trim(Local_CMHOPHDE_RT_MSG_ID) <> "CMH2101W" Then
                strErrorMessage = "RC:" + Local_CMHOPHDE_RT_MSG_ID + ", MSG:" + Local_CMHOPHDE_RT_MSG_TEXT
                MsgBox(strErrorMessage)
                Return False
                Exit Function
            End If
        Catch ex As Exception
            errMessage = "Err Code: " & CStr(Err.Number) & ", Err Desc: " & Err.Description
            'ErrorHandling.WriteLog("CMHIQHDU.formatCommArea", errMessage)
            myLogger.AppendToFile(GlobalVarable.ErrLogFile, myLogger.FormatString(True, GlobalVarable.AppName, "CMHIQHDU.formatCommArea", errMessage))
            'MsgBox(errMessage)
            Return False
        End Try
        myLogger = Nothing
        Return True
    End Function

    Private Function formatString(ByVal strInStr As String, ByVal intLen As Integer) As String

        Dim myLogger = New Logger

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
            'ErrorHandling.WriteLog("CMHIQHDU.formatString", errMessage)
            myLogger.AppendToFile(GlobalVarable.ErrLogFile, myLogger.FormatString(True, GlobalVarable.AppName, "CMHIQHDU.formatString", errMessage))
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
        Dim myLogger = New Logger

        strValue = ""
        Try
            strFormat = ""
            For i = 1 To intInteger Step 1
                strFormat = strFormat + "0"
            Next
            intTotalLen = intInteger
            If blnShowDec Then
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
            'ErrorHandling.WriteLog("CMHIQHDU.formatNumber", errMessage)
            myLogger.AppendToFile(GlobalVarable.ErrLogFile, myLogger.FormatString(True, GlobalVarable.AppName, "CMHIQHDU.formatNumber", errMessage))
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
