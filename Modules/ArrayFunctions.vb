Public Module ArrayFunctions

    Public Function IsArrEmpty(ByRef DataArr() As String) As Boolean

        Dim Index As Long

        On Error GoTo ErrHandler

        IsArrEmpty = True
        Index = UBound(DataArr)
        IsArrEmpty = False

        Exit Function

ErrHandler:
        On Error GoTo - 1
    End Function

    Public Function IsArrEmpty(ByRef DataArr() As Double) As Boolean

        Dim Index As Long

        On Error GoTo ErrHandler

        IsArrEmpty = True
        Index = UBound(DataArr)
        IsArrEmpty = False

        Exit Function

ErrHandler:
        On Error GoTo - 1
    End Function

    Public Function ArrayRemoveEndEmpty(ByVal InputArr() As String)

        Dim ReturnArr() As String
        Dim TempString As String = Join(InputArr, "|")

    End Function

    Public Function SplitWithoutEmpty(ByVal InputString As String, Optional Delimeter As String = " ")

        Dim ReturnArr() As String
        Dim TempArr() As String
        Dim TempString As String

        TempArr = Split(InputString, Delimeter)
        TempString = JoinDeletedSpaces(TempArr, Delimeter)
        ReturnArr = Split(TempString, Delimeter)

        Return ReturnArr

    End Function

    Public Function GetDelimetedString(ByVal InputString As String, ByVal Delimeter As String, ByVal ArgumentNum As Long) As String

        Dim DataArr() As String

        DataArr = Split(InputString, Delimeter)

        If ArgumentNum > DataArr.Length Then
            GetDelimetedString = "ERROR_NULL"
            Exit Function
        Else
            GetDelimetedString = DataArr(ArgumentNum - 1)
        End If

    End Function

    Public Function SeparateCharacters(ByVal InputString As String)

        Dim CharArr() As Char
        Dim ReturnArr() As String

        CharArr = InputString.ToCharArray
        ReDim ReturnArr(UBound(CharArr))

        For x As Long = 0 To UBound(CharArr)
            ReturnArr(x) = CharArr(x)
        Next x

        Return ReturnArr

    End Function

    Public Function CheckList(ByVal InputString As String, ByVal List() As String, Optional CheckIndex As Boolean = False) As Boolean

        Dim TempIndex As Long

        CheckList = False

        'If CheckIndex = True Then
        '    If InputString.Contains(".") Then
        '        TempIndex = InputString.LastIndexOf(".")
        '        InputString = InputString.Remove(TempIndex)
        '    End If
        'End If

        'TempIndex = Array.IndexOf(List, InputString)

        For x As Long = 0 To UBound(List)
            If UCase(InputString) = UCase(List(x)) Then
                CheckList = True
                Exit For
            End If
        Next x

        'new
        If CheckList = False Then

            If InputString.Contains(".") Then
                TempIndex = InputString.LastIndexOf(".")
                InputString = InputString.Remove(TempIndex)
            End If

            For x As Long = 0 To UBound(List)

                If List(x).Contains(".") Then
                    TempIndex = List(x).LastIndexOf(".")
                    List(x) = List(x).Remove(TempIndex)
                End If

                If UCase(InputString) = UCase(List(x)) Then
                    CheckList = True
                    Exit For
                End If

            Next x

        End If
        'new

    End Function

    Public Function RemoveArrayEmpty(ByVal DataArr() As String)

        Dim RetArr() As String
        Dim RetCounter As Long = 0

        For LineCounter = 0 To UBound(DataArr)
            If DataArr(LineCounter) <> vbNullString Or DataArr(LineCounter) <> Nothing Then DataArr(LineCounter) = DataArr(LineCounter).Trim
            If DataArr(LineCounter) <> vbNullString Then
                ReDim Preserve RetArr(0 To RetCounter)
                RetArr(RetCounter) = DataArr(LineCounter)
                RetCounter = RetCounter + 1
            End If
        Next LineCounter

        Return RetArr

    End Function

    Public Function ItemExist(ByVal DataArr() As String, ByVal Value As String) As Boolean

        ItemExist = False

        'old
        If Not DataArr Is Nothing Then
            For x = 0 To UBound(DataArr)
                If UCase(Value) = UCase(DataArr(x)) Then ItemExist = True
            Next x
        End If
        ''old

    End Function

    Public Function IsItemExistingInArray(ByVal RefArr() As String, ByVal ItemToCheck As String, Optional IsCaseSensitive As Boolean = True) As Boolean

        IsItemExistingInArray = False

        If Not RefArr Is Nothing Then
            If IsCaseSensitive = True Then
                If RefArr.Contains(ItemToCheck) Then IsItemExistingInArray = True
            Else
                If RefArr.Contains(ItemToCheck, StringComparer.CurrentCultureIgnoreCase) Then IsItemExistingInArray = True
            End If
        End If

    End Function

    Public Function AddItemIfNotExistingInArray(ByVal ItemToCheck As String, ByRef RefArr() As String, Optional IsCaseSensitive As Boolean = True) As Boolean

        AddItemIfNotExistingInArray = False

        If IsItemExistingInArray(RefArr, ItemToCheck, IsCaseSensitive) = False Then

            If RefArr Is Nothing Then
                ReDim Preserve RefArr(0 To 0)
                RefArr(0) = ItemToCheck
            Else
                Dim ArrLen As Long = UBound(RefArr)
                ReDim Preserve RefArr(0 To ArrLen + 1)
                RefArr(UBound(RefArr)) = ItemToCheck
            End If

            'ArrayCounter += 1
            'ReDim Preserve RefArr(0 To ArrayCounter)
            'RefArr(ArrayCounter) = ItemToCheck
            AddItemIfNotExistingInArray = True

        End If

    End Function

    Public Function CheckItemIndexInArray(ByVal ItemToCheck As String, ByVal RefArr() As String) As Long

        CheckItemIndexInArray = -1

        For LineCounter As Long = 0 To UBound(RefArr)

            If UCase(ItemToCheck).Trim = UCase(RefArr(LineCounter)).Trim Then
                CheckItemIndexInArray = LineCounter
            End If

        Next LineCounter

    End Function

    Public Function GetDataFromArray(ByVal InputString As String, ByVal RefArr() As String, ByVal Delimeter As String, ByVal Position As Long) As String

        On Error GoTo ExitHere

        Dim TempArr() As String

        GetDataFromArray = "ERR"

        For x As Long = 0 To UBound(RefArr)
            TempArr = Split(RefArr(x), Delimeter)
            If UCase(InputString) = UCase(TempArr(Position - 1)) Then
                GetDataFromArray = RefArr(x)
                Exit For
            End If
        Next x

ExitHere:
    End Function

    Public Function SortArrayOfNumbers(ByVal DataArr() As String)

        For x As Long = 0 To UBound(DataArr)
            DataArr(x) = Format(CLng(DataArr(x)), "00000")
        Next x
        Array.Sort(DataArr)
        For x As Long = 0 To UBound(DataArr)
            DataArr(x) = Format(CLng(DataArr(x)), "0")
        Next x

        Return DataArr

    End Function

    Public Function RemoveDuplicateItem(ByVal DataArr() As String)

        Dim ReturnArr() As String

        If Not DataArr Is Nothing Then ReturnArr = DataArr.Distinct().ToArray

        Return ReturnArr

    End Function

    Public Function RemoveDuplicatesInArray(ByVal DataArr() As String)

        Dim ReturnArr() As String

        If Not DataArr Is Nothing Then ReturnArr = DataArr.Distinct().ToArray

        Return ReturnArr

    End Function

    Public Function RemoveSpaceBetween(ByVal InputString As String) As String

        Dim TempArr() As String
        RemoveSpaceBetween = InputString

        TempArr = Split(RemoveSpaceBetween, " ")
        If Not TempArr Is Nothing Then
            TempArr = RemoveArrayEmpty(TempArr)
            RemoveSpaceBetween = Join(TempArr, " ")
        End If

        TempArr = Split(RemoveSpaceBetween, vbNewLine)
        If Not TempArr Is Nothing Then
            TempArr = RemoveArrayEmpty(TempArr)
            RemoveSpaceBetween = Join(TempArr, " ")
        End If

        TempArr = Split(RemoveSpaceBetween, vbLf)
        If Not TempArr Is Nothing Then
            TempArr = RemoveArrayEmpty(TempArr)
            RemoveSpaceBetween = Join(TempArr, " ")
        End If

    End Function

    Public Function ChangeArrayItem(ByVal InputArr() As String, ByVal InputString As String, Optional Operation As String = "ADD")

        Dim ReturnArr() As String

        ReturnArr = InputArr

        If UCase(Operation) = "ADD" Then
            If ReturnArr Is Nothing Then ReDim Preserve ReturnArr(0 To 0)
            If ItemExist(InputArr, InputString) = False Then
                ReDim Preserve ReturnArr(0 To (UBound(ReturnArr) + 1))
                ReturnArr(UBound(ReturnArr)) = InputString
            End If
        ElseIf UCase(Operation) = "REMOVE" Then
            If InputArr Is Nothing Then Exit Function
            For x As Long = 0 To UBound(ReturnArr)
                If UCase(InputString) = UCase(ReturnArr(x)) Then
                    ReturnArr(x) = vbNullString
                End If
            Next x
        End If
        ReturnArr = RemoveArrayEmpty(ReturnArr)
        Return ReturnArr

    End Function

    Public Function CompareArrays(ByVal RefArr() As String, ByVal ArrToCheck() As String)

        Dim ReturnArr() As String
        Dim ReturnCount As Long = -1

        If Not RefArr Is Nothing And Not ArrToCheck Is Nothing Then
            ReturnArr = RefArr.Except(ArrToCheck).ToArray
        End If

        Return ReturnArr

    End Function

    Public Function RemoveNonDouble(ByVal InputArr() As String)

        Dim RetArr() As String

        For x As Long = 0 To UBound(InputArr)
            If Not IsNumeric(InputArr(x)) Then InputArr(x) = vbNullString
        Next x

        RetArr = RemoveArrayEmpty(InputArr)

        Return RetArr

    End Function

End Module
