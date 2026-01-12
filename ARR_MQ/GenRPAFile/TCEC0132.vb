Public Class TCEC0132
    Private I_Caller_Application_ID As String
    Private I_Caller_Program_ID As String
    Private I_Calling_Type As String
    Private I_Entity As String
    Private I_Rate_Code As String
    Private I_Convention_Type As String
    Private I_K_Days As String
    Private I_Period_End_Date As String
    Private I_Period_Start_Date As String
    Private I_Days_of_Interest_Period As String

    Private I_Choice_CAR_Or_SAR As String
    Private I_Choice_CAS_Floor_Legacy_Contracts As String
    Private I_input_Filler As String

    Private o_Return_Code As String
    Private o_SQL_RETURN_CODE As String
    Private o_SQL_RETURN_MESSAGE_LENGTH As String
    Private o_SQL_RETURN_MESSAGE_DATA As String
    Private o_DB2_TABLE_NAME As String
    Private o_FILLER1 As String
    Private o_Value_of_CAR_Unadjusted As String
    Private o_Value_of_SAR_Unadjusted As String
    Private o_Value_of_CAR_Adjusted As String
    Private o_Value_of_SAR_Adjusted As String

    Private o_Value_of_Floor_Rate As String
    Private o_Floor_Rate_Applied_Indicator As String
    Private o_Value_of_CAR_CAS As String
    Private o_Value_of_SAR_CAS As String
    Private o_Floor_Rate_CAS As String
    Private o_Value_of_CAS_for_Odd_Tenor As String
    Private o_Adjusted_CAS_Rate_Table As String
    Private o_Filler2

    Private errMessage As String
    Private ErrorHandling As New Logger
    Public Property Caller_Application_ID() As String
        Get
            Return myTrim(I_Caller_Application_ID)
        End Get
        Set(ByVal value As String)
            I_Caller_Application_ID = formatString(value, 3)
        End Set
    End Property
    Public Property Caller_Program_ID() As String
        Get
            Return myTrim(I_Caller_Program_ID)
        End Get
        Set(ByVal value As String)
            I_Caller_Program_ID = formatString(value, 8)
        End Set
    End Property
    Public Property Calling_Type() As String
        Get
            Return myTrim(I_Calling_Type)
        End Get
        Set(ByVal value As String)
            I_Calling_Type = formatString(value, 1)
        End Set
    End Property
    Public Property Entity() As String
        Get
            Return myTrim(I_Entity)
        End Get
        Set(ByVal value As String)
            I_Entity = formatString(value, 4)
        End Set
    End Property
    Public Property Rate_Code() As String
        Get
            Return myTrim(I_Rate_Code)
        End Get
        Set(ByVal value As String)
            I_Rate_Code = formatString(value, 8)
        End Set
    End Property
    Public Property Convention_Type() As String
        Get
            Return myTrim(I_Convention_Type)
        End Get
        Set(ByVal value As String)
            I_Convention_Type = formatString(value, 5)
        End Set
    End Property
    Public Property K_Days() As String
        Get
            Return myTrim(I_K_Days)
        End Get
        Set(ByVal value As String)
            I_K_Days = formatString(value, 3)
        End Set
    End Property
    Public Property Period_End_Date() As String
        Get
            Return myTrim(I_Period_End_Date)
        End Get
        Set(ByVal value As String)
            I_Period_End_Date = formatString(value, 10)
        End Set
    End Property
    Public Property Period_Start_Date() As String
        Get
            Return myTrim(I_Period_Start_Date)
        End Get
        Set(ByVal value As String)
            I_Period_Start_Date = formatString(value, 10)
        End Set
    End Property
    Public Property Days_of_Interest_Period() As String
        Get
            Return myTrim(I_Days_of_Interest_Period)
        End Get
        Set(ByVal value As String)
            I_Days_of_Interest_Period = formatString(value, 5)
        End Set
    End Property
    Public Property Choice_CAR_Or_SAR() As String
        Get
            Return myTrim(I_Choice_CAR_Or_SAR)
        End Get
        Set(ByVal value As String)
            I_Choice_CAR_Or_SAR = formatString(value, 1)
        End Set
    End Property
    Public Property Choice_CAS_Floor_Legacy_Contracts() As String
        Get
            Return myTrim(I_Choice_CAS_Floor_Legacy_Contracts)
        End Get
        Set(ByVal value As String)
            I_Choice_CAS_Floor_Legacy_Contracts = formatString(value, 3)
        End Set
    End Property
    Public Property input_Filler() As String
        Get
            Return myTrim(I_input_Filler)
        End Get
        Set(ByVal value As String)
            I_input_Filler = formatString(value, 50)
        End Set
    End Property


    Public Property Return_Code() As String
        Get
            Return myTrim(o_Return_Code)
        End Get
        Set(ByVal value As String)
            o_Return_Code = formatString(value, 2)
        End Set
    End Property
    Public Property SQL_RETURN_CODE() As String
        Get
            Return myTrim(o_SQL_RETURN_CODE)
        End Get
        Set(ByVal value As String)
            o_SQL_RETURN_CODE = formatString(value, 10)
        End Set
    End Property
    Public Property SQL_RETURN_MESSAGE_LENGTH() As String
        Get
            Return myTrim(o_SQL_RETURN_MESSAGE_LENGTH)
        End Get
        Set(ByVal value As String)
            o_SQL_RETURN_MESSAGE_LENGTH = formatString(value, 5)
        End Set
    End Property
    Public Property SQL_RETURN_MESSAGE_DATA() As String
        Get
            Return myTrim(o_SQL_RETURN_MESSAGE_DATA)
        End Get
        Set(ByVal value As String)
            o_SQL_RETURN_MESSAGE_DATA = formatString(value, 30)
        End Set
    End Property
    Public Property DB2_TABLE_NAME() As String
        Get
            Return myTrim(o_DB2_TABLE_NAME)
        End Get
        Set(ByVal value As String)
            o_DB2_TABLE_NAME = formatString(value, 16)
        End Set
    End Property
    Public Property FILLER1() As String
        Get
            Return myTrim(o_FILLER1)
        End Get
        Set(ByVal value As String)
            o_FILLER1 = formatString(value, 5)
        End Set
    End Property
    Public Property Value_of_CAR_Unadjusted() As String
        Get
            Return myTrim(o_Value_of_CAR_Unadjusted)
        End Get
        Set(ByVal value As String)
            o_Value_of_CAR_Unadjusted = formatString(value, 11)
        End Set
    End Property
    Public Property Value_of_SAR_Unadjusted() As String
        Get
            Return myTrim(o_Value_of_SAR_Unadjusted)
        End Get
        Set(ByVal value As String)
            o_Value_of_SAR_Unadjusted = formatString(value, 11)
        End Set
    End Property
    Public Property Value_of_CAR_Adjusted() As String
        Get
            Return myTrim(o_Value_of_CAR_Adjusted)
        End Get
        Set(ByVal value As String)
            o_Value_of_CAR_Adjusted = formatString(value, 11)
        End Set
    End Property
    Public Property Value_of_SAR_Adjusted() As String
        Get
            Return myTrim(o_Value_of_SAR_Adjusted)
        End Get
        Set(ByVal value As String)
            o_Value_of_SAR_Adjusted = formatString(value, 11)
        End Set
    End Property
    Public Property Value_of_Floor_Rate() As String
        Get
            Return myTrim(o_Value_of_Floor_Rate)
        End Get
        Set(ByVal value As String)
            o_Value_of_Floor_Rate = formatString(value, 11)
        End Set
    End Property
    Public Property Floor_Rate_Applied_Indicator() As String
        Get
            Return myTrim(o_Floor_Rate_Applied_Indicator)
        End Get
        Set(ByVal value As String)
            o_Floor_Rate_Applied_Indicator = formatString(value, 1)
        End Set
    End Property
    Public Property Value_of_CAR_CAS() As String
        Get
            Return myTrim(o_Value_of_CAR_CAS)
        End Get
        Set(ByVal value As String)
            o_Value_of_CAR_CAS = formatString(value, 11)
        End Set
    End Property
    Public Property Value_of_SAR_CAS() As String
        Get
            Return myTrim(o_Value_of_SAR_CAS)
        End Get
        Set(ByVal value As String)
            o_Value_of_SAR_CAS = formatString(value, 11)
        End Set
    End Property
    Public Property Floor_Rate_CAS() As String
        Get
            Return myTrim(o_Floor_Rate_CAS)
        End Get
        Set(ByVal value As String)
            o_Floor_Rate_CAS = formatString(value, 1)
        End Set
    End Property
    Public Property Value_of_CAS_for_Odd_Tenor() As String
        Get
            Return myTrim(o_Value_of_CAS_for_Odd_Tenor)
        End Get
        Set(ByVal value As String)
            o_Value_of_CAS_for_Odd_Tenor = formatString(value, 11)
        End Set
    End Property
    Public Property Adjusted_CAS_Rate_Table() As String
        Get
            Return myTrim(o_Adjusted_CAS_Rate_Table)
        End Get
        Set(ByVal value As String)
            o_Adjusted_CAS_Rate_Table = formatString(value, 161)
        End Set
    End Property
    Public Property Filler2() As String
        Get
            Return myTrim(o_Filler2)
        End Get
        Set(ByVal value As String)
            o_Filler2 = formatString(value, 116)
        End Set
    End Property

    Public Function formatCommArea() As String
        Dim strBuff As String
        Dim myLogger = New Logger

        strBuff = ""
        Try
            strBuff = strBuff + Caller_Application_ID
            strBuff = strBuff + Caller_Program_ID
            strBuff = strBuff + Calling_Type
            strBuff = strBuff + Entity
            strBuff = strBuff + Rate_Code
            strBuff = strBuff + Convention_Type
            strBuff = strBuff + K_Days
            strBuff = strBuff + Period_End_Date
            strBuff = strBuff + Period_Start_Date
            strBuff = strBuff + Days_of_Interest_Period

            strBuff = strBuff + Choice_CAR_Or_SAR
            strBuff = strBuff + Choice_CAS_Floor_Legacy_Contracts
            strBuff = strBuff + input_Filler

            strBuff = strBuff + Return_Code
            strBuff = strBuff + SQL_RETURN_CODE
            strBuff = strBuff + SQL_RETURN_MESSAGE_LENGTH
            strBuff = strBuff + SQL_RETURN_MESSAGE_DATA
            strBuff = strBuff + DB2_TABLE_NAME
            strBuff = strBuff + FILLER1
            strBuff = strBuff + Value_of_CAR_Unadjusted
            strBuff = strBuff + Value_of_SAR_Unadjusted
            strBuff = strBuff + Value_of_CAR_Adjusted
            strBuff = strBuff + Value_of_SAR_Adjusted

            strBuff = strBuff + Value_of_Floor_Rate
            strBuff = strBuff + Floor_Rate_Applied_Indicator
            strBuff = strBuff + Value_of_CAR_CAS
            strBuff = strBuff + Value_of_SAR_CAS
            strBuff = strBuff + Floor_Rate_CAS
            strBuff = strBuff + Value_of_CAS_for_Odd_Tenor
            strBuff = strBuff + Adjusted_CAS_Rate_Table
            strBuff = strBuff + Filler2

        Catch ex As Exception
            errMessage = "Err Code: " & CStr(Err.Number) & ", Err Desc: " & Err.Description
            'ErrorHandling.WriteLog("TCEC0132.formatCommArea", errMessage)
            myLogger.AppendToFile(GlobalVarable.ErrLogFile, myLogger.FormatString(True, GlobalVarable.AppName, "TCEC0132.formatCommArea", errMessage))
            'MsgBox(errMessage)
        End Try
        myLogger = Nothing
        Return strBuff
    End Function

    Public Sub returnCommArea(ByVal strBuff As String)
        Dim myLogger = New Logger

        Try
            Return_Code = Mid(strBuff, 112, 2)
            SQL_RETURN_CODE = Mid(strBuff, 114, 10)
            SQL_RETURN_MESSAGE_LENGTH = Mid(strBuff, 124, 5)
            SQL_RETURN_MESSAGE_DATA = Mid(strBuff, 129, 30)
            DB2_TABLE_NAME = Mid(strBuff, 159, 16)
            FILLER1 = Mid(strBuff, 175, 5)
            Value_of_CAR_Unadjusted = Mid(strBuff, 180, 11)
            Value_of_SAR_Unadjusted = Mid(strBuff, 191, 11)
            Value_of_CAR_Adjusted = Mid(strBuff, 202, 11)
            Value_of_SAR_Adjusted = Mid(strBuff, 213, 11)

            Value_of_Floor_Rate = Mid(strBuff, 224, 11)
            Floor_Rate_Applied_Indicator = Mid(strBuff, 235, 1)
            Value_of_CAR_CAS = Mid(strBuff, 236, 11)
            Value_of_SAR_CAS = Mid(strBuff, 247, 11)
            Floor_Rate_CAS = Mid(strBuff, 258, 1)
            Value_of_CAS_for_Odd_Tenor = Mid(strBuff, 259, 11)
            Adjusted_CAS_Rate_Table = Mid(strBuff, 270, 161)
            Filler2 = Mid(strBuff, 431, 116)


        Catch ex As Exception
            errMessage = "Err Code: " & CStr(Err.Number) & ", Err Desc: " & Err.Description
            'ErrorHandling.WriteLog("TCEC0132.returnCommArea", errMessage)
            myLogger.AppendToFile(GlobalVarable.ErrLogFile, myLogger.FormatString(True, GlobalVarable.AppName, "TCEC0132.returnCommArea", errMessage))
            'MsgBox(errMessage)
        End Try
        myLogger = Nothing
    End Sub
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
            'ErrorHandling.WriteLog("TCEC0132.formatString", errMessage)
            myLogger.AppendToFile(GlobalVarable.ErrLogFile, myLogger.FormatString(True, GlobalVarable.AppName, "TCEC0132.formatString", errMessage))
            'MsgBox(errMessage)
        End Try
        myLogger = Nothing
        Return strInStr
    End Function
    Private Function myTrim(ByVal strValue As String)
        strValue = Replace(strValue, Chr(0), Space(1))
        'strValue = Trim(strValue)
        Return strValue
    End Function
End Class
