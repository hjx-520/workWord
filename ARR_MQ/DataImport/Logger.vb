Public Class Logger

    '--------------------------------------------------------------------
    ' Generate the Directories
    ' Parameter filePath can be the pure path, or filename with absolutely path
    ' Return the error message if has
    '--------------------------------------------------------------------
    Function CrtFolder(ByVal folderPath As String) As String

        Dim errorMsg As String

        Dim fs As Object

        Dim nextIndex As Short
        Dim currentIndex As Short

        Dim filePath As String
        On Error GoTo CrtfolderErr

        fs = CreateObject("Scripting.FileSystemObject")

        currentIndex = 1
        nextIndex = 1
        Do While nextIndex > 0

            ' Search the highest path in the below the current level
            nextIndex = InStr(currentIndex, folderPath, "\")

            ' If there's a directory
            If nextIndex > 0 Then

                ' Get the directory
                filePath = Mid(folderPath, 1, nextIndex)

                ' Generate the directory
                'UPGRADE_WARNING: Couldn't resolve default property of object fs.FolderExists. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If Not fs.FolderExists(filePath) Then
                    'UPGRADE_WARNING: Couldn't resolve default property of object fs.CreateFolder. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    fs.CreateFolder(filePath)
                End If
            End If

            currentIndex = nextIndex + 1
        Loop
        'UPGRADE_NOTE: Object fs may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        fs = Nothing

        CrtFolder = ""
        Exit Function

CrtfolderErr:
        ' Set error message and return the error message.
        '   Because this function might be called in "WriteErrorLog" process,
        '   it can't write error log directly to avoid endless-looping
        'UPGRADE_NOTE: Object fs may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        fs = Nothing
        errorMsg = "CrtFolder() - " & ErrorToString()
        CrtFolder = errorMsg
    End Function

    Public Function AppendToFile(ByVal FileName As String, ByVal Content As String) As String
        Dim result As String
        Dim errorMsg As String
        Dim s As String
        Dim FileNumber As Object

        On Error GoTo AppendToFileErr
        result = ""

        ' If there's no warning log file name, do nothing
        If Trim(FileName) = "" Then
            'Upgrade By Alex.Wang, 
            AppendToFile = result
            Exit Function
        End If

        ' Create folder of warning log
        s = CrtFolder(FileName)
        ' If there's error while create folder, then stop
        If s <> "" Then
            errorMsg = s
            GoTo AppendToFileErr
        End If

        ' Generate the file, if needed.
        'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        If Dir(FileName) = "" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object FileNumber. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            FileNumber = FreeFile()
            'UPGRADE_WARNING: Couldn't resolve default property of object FileNumber. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            FileOpen(FileNumber, FileName, OpenMode.Output)
            'UPGRADE_WARNING: Couldn't resolve default property of object FileNumber. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            FileClose(FileNumber)
        End If

        On Error GoTo AppendToFileErr1

        ' Write content
        'UPGRADE_WARNING: Couldn't resolve default property of object FileNumber. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        FileNumber = FreeFile()
        'UPGRADE_WARNING: Couldn't resolve default property of object FileNumber. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        FileOpen(FileNumber, FileName, OpenMode.Append)
        'UPGRADE_WARNING: Couldn't resolve default property of object FileNumber. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(FileNumber, Content)
        'UPGRADE_WARNING: Couldn't resolve default property of object FileNumber. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        FileClose(FileNumber)
        AppendToFile = result
        Exit Function

AppendToFileErr:
        If (errorMsg = "") Then
            errorMsg = Err.Description
        End If

        'Set error message
        errorMsg = "AppendToFile() - " & errorMsg & Chr(13) & Chr(10) & "       FileName: " & FileName & ", Content: " & Content
        '    errorMsg = App.EXEName & " cause error " & errorMsg
        result = errorMsg
        AppendToFile = result
        Exit Function

AppendToFileErr1:  '' The lable that deal with error message, and close file handle.
        On Error Resume Next
        If (errorMsg = "") Then
            errorMsg = Err.Description
        End If

        'Set error message
        errorMsg = "AppendToFile() - " & errorMsg & Chr(13) & Chr(10) & "       FileName: " & FileName & ", Content: " & Content
        'UPGRADE_WARNING: Couldn't resolve default property of object FileNumber. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        FileClose(FileNumber)
        result = errorMsg
        AppendToFile = result

    End Function

    Public Function FormatString(ByVal LongDateTime As Boolean, ByVal ProcessName As String, ByVal ModuleName As String, ByVal Content As String) As String

        Dim OutputContent As String

        Try
            If LongDateTime Then
                OutputContent = Format(Now, "yyyy-MM-dd hh:mm:ss tt") & " " & Left(ProcessName & Space(20), 20) & Content
            Else
                OutputContent = Format(Now, "yyyy-MM-dd") & " " & Left(ProcessName & Space(20), 20) & Content
            End If
            Return OutputContent
        Catch ex As Exception
            Return Format(Now, "yyyy-MM-dd hh:mm:ss tt") & " " & Left(ProcessName & Space(20), 20) & ex.Message
        End Try

    End Function


End Class
