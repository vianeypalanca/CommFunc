Imports System.IO

Public Class clsText

    Public Function ReadText(ByVal FileName As String, Optional DisplayErr As Boolean = False)

        'FileName - file to be read
        'This fucntion reads FileName if it exist and put to all lines to DataArray
        'Error Handling - Shows a message box indicating that the parameter input does not exist

        Dim DataArray() As String

        If My.Computer.FileSystem.FileExists(FileName) Then
            DataArray = File.ReadAllLines(FileName)
        Else
            If DisplayErr = True Then MsgBox("""" & FileName & """" & " does not exist.")
        End If

        Return DataArray

    End Function

    Public Function ReadTextOneString(ByVal FileName As String) As String

        Dim objReader As New StreamReader(FileName)
        ReadTextOneString = objReader.ReadToEnd
        objReader = Nothing

    End Function

    Public Function SaveText(ByVal FileName As String, ByVal TextToSave As String, ByVal Append As Boolean)

        'FileName - file to be saved
        'TextToSave - what will be the text that will be put in FileName in string format
        'Append - choose if TextToSave is to be appended to FileName or not, True if Append, False if it overwrites the file indicated in FileName
        'This function saves the TextToSave string to FileName and append or overwrite depending on Append
        'Error Handling - retry the process

        On Error GoTo ErrHandler
ErrHandler:
        If My.Computer.FileSystem.FileExists(FileName) Then
            Dim objWriter As New System.IO.StreamWriter(FileName, Append)
            objWriter.WriteLine(TextToSave)
            objWriter.Close()
        Else
            System.IO.File.Create(FileName).Dispose()
            Dim objWriter As New System.IO.StreamWriter(FileName, Append)
            objWriter.WriteLine(TextToSave)
            objWriter.Close()
        End If

    End Function

    Public Function SaveArray(ByVal FileName As String, ByVal DataArray() As String, Optional ToDelete As Boolean = True)
        'Create new file named FileName

        If My.Computer.FileSystem.FileExists(FileName) And ToDelete = True Then
            IO.File.Delete(FileName)
        End If

        IO.File.WriteAllLines(FileName, DataArray)

    End Function

End Class
