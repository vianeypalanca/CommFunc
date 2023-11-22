Imports System.IO

Public Module FileFunctions

    Public Function IsFolderExisting(ByVal FolderPath As String) As Boolean

        'Type:
        'Boolean

        'Description:
        'IsFolderExisting function checks if the value for the "FolderPath" argument specified is available in the user's computer

        'Return values:
        'True - if the folder specified in the value for the "FolderPath" argument is available in the user's computer
        'False - if the folder specified in the value for the "FolderPath" argument is not available in the user's computer

        IsFolderExisting = False
        If (My.Computer.FileSystem.DirectoryExists(FolderPath)) Then IsFolderExisting = True

    End Function

    Public Function IsFileExisting(ByVal FileName As String) As Boolean

        'Type:
        'Boolean

        'Description:
        'IsFileExisting function checks if the value for the "FileName" argument specified is available in the user's computer

        'Return values:
        'True - if the file specified in the value for the "FileName" argument is available in the user's computer
        'False - if the file specified in the value for the "FileName" argument is not available in the user's computer

        IsFileExisting = False
        If My.Computer.FileSystem.FileExists(FileName) Then IsFileExisting = True

    End Function

    Public Function CreateFolder(ByVal FolderPath As String) As Boolean

        If IsFolderExisting(FolderPath) = False Then
            My.Computer.FileSystem.CreateDirectory(FolderPath)
        End If

    End Function

    Public Function CreateFile(ByVal FileName As String)

        If IsFileExisting(FileName) = False Then
            File.Create(FileName).Dispose()
        End If

    End Function

    Public Function CopyFile(ByVal SourceFileName As String, ByVal DestinationFileName As String, Optional Overwrite As Boolean = False) As Boolean

        CopyFile = False

        If IsFileExisting(SourceFileName) = True Then
            File.Copy(SourceFileName, DestinationFileName, Overwrite)
            CopyFile = True
        End If

    End Function

    Public Function DeleteFile(ByVal FileName As String) As Boolean

        DeleteFile = False

        If IsFileExisting(FileName) = True Then
            File.Delete(FileName)
            DeleteFile = True
        End If

    End Function

    Public Function GetFilesInFolder(ByVal FolderPath As String, Optional FileType As String = vbNullString, Optional IsGetSubFol As Boolean = False)

        Dim ReturnArr() As String

        ReturnArr = Nothing

        If IsFolderExisting(FolderPath) = True Then

            If FileType = vbNullString Then
                ReturnArr = Directory.GetFiles(FolderPath)
            Else
                ReturnArr = Directory.GetFileSystemEntries(FolderPath, "*." & FileType)
            End If

            If IsGetSubFol = True Then
                If ReturnArr Is Nothing Then ReDim ReturnArr(0)
                Call GetAllFilesInFolder(FolderPath, ReturnArr, FileType)
                If IsArrEmpty(ReturnArr) = False Then ReturnArr = RemoveArrayEmpty(ReturnArr)
            End If

        End If

        Return ReturnArr

    End Function

    Private Sub GetAllFilesInFolder(ByVal FolderPath As String, ByRef RetArr() As String,
                                   Optional FileType As String = vbNullString)

        Dim SubArr() As String

        If IsFolderExisting(FolderPath) = True Then
            For Each SubFolderStr As String In Directory.GetDirectories(FolderPath)

                SubArr = Nothing

                If FileType = vbNullString Then
                    SubArr = Directory.GetFiles(SubFolderStr)
                Else
                    SubArr = Directory.GetFileSystemEntries(SubFolderStr, "*." & FileType)
                End If

                If SubArr IsNot Nothing Then
                    ReDim Preserve RetArr(0 To UBound(RetArr) + (UBound(SubArr) + 1))
                    SubArr.CopyTo(RetArr, UBound(RetArr) - UBound(SubArr))
                    Call GetAllFilesInFolder(SubFolderStr, RetArr, FileType)
                End If

            Next SubFolderStr
        End If

    End Sub

    Public Function GetFileName(ByVal FileString As String, Optional IsExtension As Boolean = False) As String

        If FileString <> vbNullString Then
            If IsExtension = False Then
                GetFileName = Path.GetFileNameWithoutExtension(FileString)
            Else
                GetFileName = Path.GetFileName(FileString)
            End If

        Else
            GetFileName = ErrMsg
        End If

    End Function

    Public Function GetFolderName(ByVal FileString As String) As String

        If FileString <> vbNullString Then
            GetFolderName = Path.GetDirectoryName(FileString)
        Else
            GetFolderName = ErrMsg
        End If

    End Function

    Public Function GetFileType(ByVal FileString As String) As String

        If FileString <> vbNullString Then
            GetFileType = Path.GetExtension(FileString)
        Else
            GetFileType = ErrMsg
        End If

    End Function

    Public Function GetFileNameWithString(ByVal FolderPath As String, ByVal FileString As String) As String

        Dim Files() As String = GetFilesInFolder(FolderPath)
        GetFileNameWithString = vbNullString

        If Not Files Is Nothing Then
            For Each File As String In Files
                If UCase(File).Contains(UCase(FileString)) Then
                    GetFileNameWithString = File
                    Exit For
                End If
            Next File
        End If

    End Function

    Public Function GenerateRandomFileName(NumOfCharacters As Integer, ExtensionName As String) As String

        Dim PatterName() As String
        Dim FinalPatName As String = ""
        Dim Characs As String = "a b c d e f g h i j k l m n o p q r s t u v w x y z A B C D E F G H I J K L M N O P Q R S T U V W X Y Z 1 2 3 4 5 6 7 8 9 0"

        PatterName = Split(Characs, " ")
        Dim i1st As System.Random = New System.Random()


        For i As Integer = 0 To NumOfCharacters
            Dim rnd As Integer = i1st.Next(0, PatterName.Length)
            FinalPatName += PatterName(rnd)
        Next
        GenerateRandomFileName = FinalPatName + ExtensionName
        Return GenerateRandomFileName

    End Function

End Module
