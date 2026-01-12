Imports System.Text
Imports IBM.WMQ


Public Module MQUtilities

    Public Function MQGetMessage(mqQueue As MQQueue) As MQMessage
        Dim mqMessage = New MQMessage()

        Dim mqGetMessageOpts = New MQGetMessageOptions() With {
                .WaitInterval = MQC.MQWI_UNLIMITED,
                .Options = .Options + MQC.MQGMO_WAIT + MQC.MQGMO_ACCEPT_TRUNCATED_MSG + MQC.MQGMO_CONVERT
        }

        mqQueue.Get(mqMessage, mqGetMessageOpts)

        Return mqMessage

    End Function

    Public Sub MQPutMessage(mqQueue As MQQueue, mqMessage As MQMessage)
        Dim mqPutMessageOpts As MQPutMessageOptions = New MQPutMessageOptions()
        mqQueue.Put(mqMessage, mqPutMessageOpts)
    End Sub

    'convert bytes() to string contain only hexadecimal symbols 0-9 + A-F
    Public Function BytesToHexString(bytes As Byte()) As String
        Dim strTemp As New StringBuilder(bytes.Length * 2)
        For Each b As Byte In bytes
            Dim hex = Conversion.Hex(b)
            strTemp.Append(hex.PadLeft(2, "0"))
        Next
        Return strTemp.ToString()

    End Function
    'a reverse function of BytesToHexString
    Public Function HexStringToBytes(str As String) As Byte()

        If str.Length Mod 2 <> 0 Then
            Throw New Exception("Incorrect string format")
        End If

        Dim aryLength As Integer = str.Length / 2

        Dim bytes() As Byte = New Byte(aryLength - 1) {}
        For i = 0 To aryLength - 1
            bytes(i) = Byte.Parse(str(i * 2) & str(i * 2 + 1), Globalization.NumberStyles.HexNumber)
        Next

        Return bytes

    End Function

End Module
